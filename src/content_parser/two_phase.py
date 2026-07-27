"""Two-phase content generation (см. docs/IMPROVEMENT_PLAN.md, пункт 4).

Phase A: one call — brief + a compact menu of what the template offers →
an outline (ordered list of {role, theme, count}). The model plans structure,
code enforces the hard rules it doesn't reliably follow (observed with GigaChat:
"at most one divider" and "always fill the title" were both ignored when asked
in a single big prompt).

Phase B: one small call per block — role + theme + EXACT item count, known
because a concrete template slide was already chosen by the matcher. One
request carrying one requirement is the reliability unit this model can
actually deliver; the old single-shot parse_brief carried four at once and
statistically dropped some on every run."""

import json
import re

from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, call_with_model_fallback
from common.synthesis import SYNTHESIZABLE_TYPES
from template_spec.builder import describe_for_prompt

OUTLINE_ROLES = [
    "title", "section_divider", "bullet_list", "stats_kpi",
    "two_column_comparison", "image_caption", "closing",
]

MAX_SECTION_DIVIDERS = 1
MAX_IMAGE_SLIDES = 1
MAX_BLOCKS = 10

OUTLINE_PROMPT = """Ты планируешь структуру презентации по брифу.

В шаблоне презентации доступны такие виды слайдов:
__MENU__

Разбей бриф на 5-9 блоков. Для каждого блока укажи:
- "role" — один из: __ROLES__
- "theme" — тема блока, 3-7 слов (о чём конкретно этот слайд)
- "count" — сколько пунктов/цифр планируется (для bullet_list и stats_kpi;
  подбирай под доступные ёмкости из списка выше; для остальных ролей — null)

Первый блок — всегда role="title". ПОСЛЕДНИЙ блок — всегда role="closing"
(итог и призыв к действию), презентация обязана иметь финал. Где визуал усилит слайд (обзор продукта,
процесс, результат), добавь 1 блок role="image_caption" — под него зарезервируется
место под иллюстрацию.

Ответь ТОЛЬКО валидным JSON-массивом, без пояснений и markdown:
[{"role": "...", "theme": "...", "count": 3}, ...]

Бриф:
---
__BRIEF__
---
"""

BLOCK_PROMPTS = {
    "title": """Напиши контент титульного слайда презентации по брифу.
Тема слайда: __THEME__
Ответь ТОЛЬКО валидным JSON: {"title": "название продукта/презентации, до 8 слов", "subtitle": "подзаголовок, до 12 слов"}

Бриф:
---
__BRIEF__
---""",
    "section_divider": """Напиши заголовок слайда-перехода между разделами презентации.
Тема раздела: __THEME__
Ответь ТОЛЬКО валидным JSON: {"title": "короткая фраза, 1-4 слова"}""",
    "image_caption": """Слайд с иллюстрацией. Картинка ПОКА не генерируется — нужно
только описать, что на ней должно быть, чтобы зарезервировать место в композиции.
Тема слайда: __THEME__
Используй НАСТОЯЩЕЕ название продукта и конкретику из брифа. ЗАПРЕЩЕНЫ заглушки
вроде "Продукт X", "название продукта", "ваш продукт". Описание картинки должно быть
конкретной сценой (что именно видно на изображении), а не общими словами вроде
"иллюстрация продукта".
Ответь ТОЛЬКО валидным JSON: {"title": "заголовок слайда, до 7 слов", "image": "что изобразить, 3-8 слов"}

Бриф:
---
__BRIEF__
---""",
    "closing": """Напиши контент финального слайда презентации.
Тема: __THEME__
Призыв должен быть КОНКРЕТНЫМ и опираться на бриф: назови продукт и следующий шаг —
что именно предлагается сделать и кому. Соблюдай тон из брифа. ЗАПРЕЩЕНЫ пустые
лозунги без конкретики ("Действуйте немедленно", "Не упустите шанс", "Начните
трансформацию", "Не теряйте преимущество") и восклицательные знаки.
Ответь ТОЛЬКО валидным JSON: {"title": "финальная фраза/призыв, до 6 слов", "subtitle": "подпись, до 10 слов"}

Бриф:
---
__BRIEF__
---""",
    "bullet_list": """Напиши контент слайда-списка для презентации по брифу.
Тема слайда: __THEME__
Ответь ТОЛЬКО валидным JSON:
{"title": "заголовок слайда, до 7 слов", "bullets": [РОВНО __COUNT__ строк, каждая — короткая фраза до 60 символов]}

Бриф:
---
__BRIEF__
---""",
    "stats_kpi": """Напиши контент слайда с ключевыми цифрами для презентации по брифу.
Тема слайда: __THEME__
Ответь ТОЛЬКО валидным JSON:
{"title": "заголовок слайда, до 7 слов", "stats": [РОВНО __COUNT__ пар вида ["число", "подпись до 6 слов"]]}

Первый элемент пары — ТОЛЬКО величина: цифры со знаком/единицей и НИЧЕГО больше.
Правильно: "+25%", "-30 часов", "95%", "3 дня", "8 из 10".
НЕПРАВИЛЬНО: "+18% конверсии", "-20% затрат времени", "90% точность прогноза" —
название метрики идёт во ВТОРОЙ элемент (подпись), а не в число.

Бриф:
---
__BRIEF__
---""",
    "two_column_comparison": """Напиши контент слайда-сравнения (две колонки) для презентации по брифу.
Тема слайда: __THEME__
Ответь ТОЛЬКО валидным JSON:
{"title": "заголовок слайда, до 7 слов",
 "left_heading": "заголовок левой колонки, 1-3 слова",
 "left_points": [РОВНО __COUNT__ коротких фраз],
 "right_heading": "заголовок правой колонки, 1-3 слова",
 "right_points": [РОВНО __COUNT__ коротких фраз]}

Бриф:
---
__BRIEF__
---""",
}


THEME_JUNK = {"", "null", "none", "нет"}


def _validate_outline(outline):
    if not isinstance(outline, list) or not outline:
        raise ValueError("outline is not a non-empty list")
    for item in outline:
        if not isinstance(item, dict) or item.get("role") not in OUTLINE_ROLES:
            raise ValueError(f"bad outline item: {item!r}")
        theme = str(item.get("theme") or "").strip()
        # The model occasionally emits the literal string "null" as a theme.
        item["theme"] = theme if theme.lower() not in THEME_JUNK else item["role"]
    return outline


def _enforce_outline_rules(outline):
    """Hard rules the model has demonstrably ignored when merely asked."""
    result = []
    dividers = 0
    images = 0
    for item in outline[:MAX_BLOCKS]:
        if item["role"] == "section_divider":
            dividers += 1
            if dividers > MAX_SECTION_DIVIDERS:
                continue
        if item["role"] == "image_caption":
            images += 1
            if images > MAX_IMAGE_SLIDES:
                continue
        result.append(item)
    if not result or result[0]["role"] != "title":
        result.insert(0, {"role": "title", "theme": "титульный слайд", "count": None})

    # A deck must END. The prompt asks for a closing "if the brief suits one",
    # and the model takes that as optional — a real deck came back finishing on a
    # stats slide, no wrap-up, no call to action, while the brief explicitly
    # asked for one. Every presentation needs a last word, so make it a rule.
    # Trailing non-closing blocks stay; the closing is appended after them.
    if not any(item["role"] == "closing" for item in result):
        result = result[:MAX_BLOCKS - 1] if len(result) >= MAX_BLOCKS else result
        result.append({"role": "closing", "theme": "итог и призыв к действию", "count": None})

    # Asking for an image slide "where a visual helps" gets one only about half
    # the time — same lesson as the divider cap: state it as a rule, enforce it
    # in code. Placed just before the closing (a visual right before the CTA),
    # and only when there's room under MAX_BLOCKS.
    if images == 0 and len(result) < MAX_BLOCKS:
        # Theme wording matters twice over: the model ECHOES it into the title,
        # and a vague one starves the block. "визуальная иллюстрация продукта"
        # produced a generic "Продукт X — визуализация"; a concrete, natural
        # phrasing yields a grounded title ("Платформа «Поток» в действии").
        # Measured against the real model before settling on this wording.
        item = {"role": "image_caption", "theme": "как выглядит продукт в работе", "count": None}
        if result[-1]["role"] == "closing":
            result.insert(len(result) - 1, item)
        else:
            result.append(item)
    return result


def generate_outline(client, brief, spec, models=TEXT_MODELS, style_preamble=""):
    menu = describe_for_prompt(spec, SYNTHESIZABLE_TYPES)
    if style_preamble:
        # Template design brief (plan 9.4) — advice for tone/length; every
        # hard limit stays code-enforced below regardless of what it says.
        menu = f"{menu}\n\n{style_preamble}"
    prompt = (
        OUTLINE_PROMPT
        .replace("__MENU__", menu)
        .replace("__ROLES__", ", ".join(OUTLINE_ROLES))
        .replace("__BRIEF__", brief)
    )

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=1200)
        return _validate_outline(extract_json(result["choices"][0]["message"]["content"]))

    outline = call_with_model_fallback(call, models)
    return _enforce_outline_rules(outline)


# A KPI figure is a number plus at most a short unit: "+25%", "-30 часов",
# "8 из 10", "3 дня". Anything wordier is the metric NAME leaking into the
# number slot (where it renders as a wall of orange text instead of a punchy
# figure) - it belongs in the label underneath.
MAX_FIGURE_WORDS = 3
MAX_FIGURE_CHARS = 10


def _is_display_figure(num):
    text = str(num).replace(" ", " ").strip()
    if not text or not any(c.isdigit() for c in text):
        return False
    return len(text) <= MAX_FIGURE_CHARS and len(text.split()) <= MAX_FIGURE_WORDS


def _reject_wordy_figures(block, role):
    """Raise (-> retry) when a stat's number slot holds a phrase, not a figure.
    Seen on a real deck: "+18% konversii", "-20% zatrat vremeni", "90% tochnost
    prognoza" - the KPI row read as three orange sentences while the labels
    below repeated the same words. Checked only in the retrying path, so a
    stubborn model costs a couple of retries, never the whole run."""
    if role != "stats_kpi":
        return
    bad = [num for num, _ in block.get("stats", []) if not _is_display_figure(num)]
    if bad:
        raise ValueError(f"stat numbers must be bare figures, got {bad!r}")


def _validate_block(block, role, count):
    if not isinstance(block, dict):
        raise ValueError("block is not a dict")
    if not block.get("title"):
        raise ValueError("empty title")
    if role == "bullet_list":
        if len(block.get("bullets", [])) != count:
            raise ValueError(f"expected {count} bullets, got {len(block.get('bullets', []))}")
    elif role == "stats_kpi":
        stats = block.get("stats", [])
        if len(stats) != count or not all(isinstance(s, list) and len(s) == 2 for s in stats):
            raise ValueError(f"expected {count} [num, label] pairs")
    elif role == "two_column_comparison":
        if len(block.get("left_points", [])) != count or len(block.get("right_points", [])) != count:
            raise ValueError("column point counts don't match")
    elif role == "image_caption":
        if not str(block.get("image", "")).strip():
            raise ValueError("image_caption needs an 'image' description")
    return block


# Text budgets (plan 10б), enforced deterministically AFTER generation rather
# than via validation retries: length overruns are frequent and gradual (a
# retry often returns another overrun, burning calls), while trimming is
# lossless enough — a title cut to its first sentence stays a title. Numeric
# COUNTS keep using retries: a wrong count can't be repaired locally.
MAX_TITLE_WORDS = 9
MAX_BULLET_CHARS = 72


def _first_sentence(text):
    for sep in (". ", "! ", "? "):
        if sep in text:
            return text.split(sep, 1)[0].rstrip(".!?")
    return text


def _guard_widow(text):
    """Glue a bullet's last two words with a non-breaking space so a wrap can't
    strand the final word alone on its own line — the «сирота»/widow the eye
    catches on narrow two-column bullets («…принятия / решений»). Only for >=3
    words, so line 1 always keeps at least one word even if the glued pair wraps.
    No-op on text that has no space to replace. U+00A0 renders identically to a
    space and is honoured as non-breaking by LibreOffice/PowerPoint."""
    words = text.split()
    if len(words) < 3:
        return text
    return " ".join(words[:-1]) + " " + words[-1]


def _enforce_text_budgets(block, role):
    title = str(block.get("title") or "")
    if len(title.split()) > MAX_TITLE_WORDS:
        title = _first_sentence(title)
        if len(title.split()) > MAX_TITLE_WORDS:
            title = " ".join(title.split()[:MAX_TITLE_WORDS])
        block["title"] = title.rstrip(".,;: ")
    for field in ("bullets", "left_points", "right_points"):
        if field in block:
            trimmed = []
            for item in block[field]:
                item = str(item)
                if len(item) > MAX_BULLET_CHARS:
                    cut = item[:MAX_BULLET_CHARS].rsplit(" ", 1)[0]
                    item = cut.rstrip(".,;: ")
                trimmed.append(_guard_widow(item))
            block[field] = trimmed
    # stats_kpi labels widow too (narrow KPI columns): «Экономия времени на /
    # отчёты». Guard the LABEL (pair[1]); never touch the number (pair[0]).
    stats = block.get("stats")
    if isinstance(stats, list):
        block["stats"] = [
            [pair[0], _guard_widow(str(pair[1]))]
            if isinstance(pair, (list, tuple)) and len(pair) == 2 else pair
            for pair in stats
        ]
    return block


def stat_fingerprints(block):
    """Numbers and labels a stats_kpi block occupies, normalized for comparison.
    Blocks are generated one call at a time with no knowledge of each other, so
    two stat slides routinely came out with the SAME numbers — see
    _reject_duplicate_stats."""
    nums, labels = set(), set()
    for pair in block.get("stats", []) or []:
        if isinstance(pair, (list, tuple)) and len(pair) == 2:
            num, label = pair
            if str(num).strip():
                nums.add(str(num).strip().lower())
            norm = re.sub(r"\s+", " ", str(label).replace(" ", " ")).strip().lower()
            if norm:
                labels.add(norm)
    return nums, labels


def _reject_duplicate_stats(block, role, used_nums, used_labels):
    """Raise (-> retry) when a stat slide reuses a number or label another slide
    already showed. Observed on a real deck: slides 5 and 6 both read
    +25% / -40% / 92% with near-identical labels — visually one slide shown
    twice. The model can't know this on its own; each block is a separate call."""
    if role != "stats_kpi":
        return
    nums, labels = stat_fingerprints(block)
    clash = (nums & used_nums) | (labels & used_labels)
    if clash:
        raise ValueError(f"stats repeat what another slide already shows: {sorted(clash)}")


def generate_block(client, role, theme, brief, count=None, models=TEXT_MODELS,
                   style_preamble="", used_stats=None):
    """One block, one call, one requirement (the exact count) — sized for the
    already-chosen slide, so nothing needs resizing after the fact.
    style_preamble: the template design brief (plan 9.4), advisory only.
    used_stats: (numbers, labels) already shown on earlier stat slides — asked
    for in the prompt AND enforced by validation+retry, because asking alone
    doesn't hold (project rule: hard constraints go in code)."""
    theme_line = f"{theme}\n{style_preamble}" if style_preamble else theme
    used_nums, used_labels = used_stats or (set(), set())
    prompt = (
        BLOCK_PROMPTS[role]
        .replace("__THEME__", theme_line)
        .replace("__BRIEF__", brief)
        .replace("__COUNT__", str(count or 3))
    )
    if role == "stats_kpi" and (used_nums or used_labels):
        prompt += (
            "\n\nЭТИ цифры и метрики УЖЕ показаны на другом слайде — возьми ДРУГИЕ "
            "показатели, не повторяй ни числа, ни формулировки:\n"
            + "; ".join(sorted(used_nums | used_labels))
        )

    def make_call(enforce_unique):
        def call(model):
            result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=800)
            block = _validate_block(extract_json(result["choices"][0]["message"]["content"]),
                                    role, count or 3)
            if enforce_unique:
                _reject_wordy_figures(block, role)
                _reject_duplicate_stats(block, role, used_nums, used_labels)
            return block
        return call

    try:
        block = call_with_model_fallback(make_call(True), models)
    except ValueError:
        # Couldn't get distinct metrics after the retries: a duplicated stat
        # slide is a defect, a crashed iteration is worse. Take the block.
        block = call_with_model_fallback(make_call(False), models)
    block["type"] = role
    return _enforce_text_budgets(block, role)
