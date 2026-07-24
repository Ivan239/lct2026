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

from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, call_with_model_fallback
from common.synthesis import SYNTHESIZABLE_TYPES
from template_spec.builder import describe_for_prompt

OUTLINE_ROLES = [
    "title", "section_divider", "bullet_list", "stats_kpi",
    "two_column_comparison", "closing",
]

MAX_SECTION_DIVIDERS = 1
MAX_BLOCKS = 10

OUTLINE_PROMPT = """Ты планируешь структуру презентации по брифу.

В шаблоне презентации доступны такие виды слайдов:
__MENU__

Разбей бриф на 5-9 блоков. Для каждого блока укажи:
- "role" — один из: __ROLES__
- "theme" — тема блока, 3-7 слов (о чём конкретно этот слайд)
- "count" — сколько пунктов/цифр планируется (для bullet_list и stats_kpi;
  подбирай под доступные ёмкости из списка выше; для остальных ролей — null)

Первый блок — всегда role="title". Завершай блоком role="closing", если брифу
подходит финал с призывом к действию.

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
    "closing": """Напиши контент финального слайда презентации.
Тема: __THEME__
Ответь ТОЛЬКО валидным JSON: {"title": "финальная фраза/призыв, до 6 слов", "subtitle": "подпись, до 10 слов"}""",
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
{"title": "заголовок слайда, до 7 слов", "stats": [РОВНО __COUNT__ пар вида ["число или короткая величина", "подпись до 6 слов"]]}

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
    for item in outline[:MAX_BLOCKS]:
        if item["role"] == "section_divider":
            dividers += 1
            if dividers > MAX_SECTION_DIVIDERS:
                continue
        result.append(item)
    if not result or result[0]["role"] != "title":
        result.insert(0, {"role": "title", "theme": "титульный слайд", "count": None})
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
                trimmed.append(item)
            block[field] = trimmed
    return block


def generate_block(client, role, theme, brief, count=None, models=TEXT_MODELS,
                   style_preamble=""):
    """One block, one call, one requirement (the exact count) — sized for the
    already-chosen slide, so nothing needs resizing after the fact.
    style_preamble: the template design brief (plan 9.4), advisory only."""
    theme_line = f"{theme}\n{style_preamble}" if style_preamble else theme
    prompt = (
        BLOCK_PROMPTS[role]
        .replace("__THEME__", theme_line)
        .replace("__BRIEF__", brief)
        .replace("__COUNT__", str(count or 3))
    )

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=800)
        return _validate_block(extract_json(result["choices"][0]["message"]["content"]), role, count or 3)

    block = call_with_model_fallback(call, models)
    block["type"] = role
    return _enforce_text_budgets(block, role)
