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
from common.phrases import (DANGLING_TAIL_WORDS, cut_at_clause, cut_at_clause_chars,
                            drop_dangling_function_words,
                            unfilled_placeholders)
from common.model_fallback import TEXT_MODELS, call_with_model_fallback
from common.prompt_files import load_prompt
from common.synthesis import SYNTHESIZABLE_TYPES
from template_spec.builder import describe_for_prompt

OUTLINE_ROLES = [
    "title", "section_divider", "bullet_list", "stats_kpi",
    "two_column_comparison", "image_caption", "closing",
]

MAX_SECTION_DIVIDERS = 1
MAX_IMAGE_SLIDES = 1

# Deck size from the VK Tech brief: 10-15 slides, or the number the user asks
# for (a content package's `slides`). The outline prompt used to ask for 5-9
# blocks under a hard cap of 10, so every loop deck came out 8-9 slides — under
# the brief's floor on every run.
DECK_MIN_SLIDES = 10
DECK_MAX_SLIDES = 15
MAX_BLOCKS = DECK_MAX_SLIDES

# Prompts live in prompts/ as versioned files (backlog item 9): a changed
# prompt is a new file, so a run stays reproducible, and switching the
# generator to an open-weights model (item 1) means swapping text files, not
# code. The loaded text is byte-for-byte the text that used to be inline
# (test_prompt_files checks the files and placeholders).
OUTLINE_PROMPT = load_prompt("outline.v1.txt")

BLOCK_PROMPTS = {
    role: load_prompt(f"block_{role}.v1.txt")
    for role in ("title", "section_divider", "image_caption", "closing",
                 "bullet_list", "stats_kpi", "two_column_comparison")
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


def _enforce_outline_rules(outline, max_blocks=MAX_BLOCKS):
    """Hard rules the model has demonstrably ignored when merely asked."""
    result = []
    dividers = 0
    images = 0
    for item in outline[:max_blocks]:
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
    #
    # And it ends ONCE, at the end. The model also returns a closing in the
    # middle or two of them — the loop deck of iter103 ran image slide after
    # «Спасибо», and the image insertion below appends after the last block
    # when that block is not a closing. The last closing the model wrote is
    # kept (it is the one written as the ending) and moved to the end; the
    # blocks it stood before keep their order.
    closings = [item for item in result if item["role"] == "closing"]
    result = [item for item in result if item["role"] != "closing"]
    if closings:
        result.append(closings[-1])
    else:
        result = result[:max_blocks - 1] if len(result) >= max_blocks else result
        result.append({"role": "closing", "theme": "итог и призыв к действию", "count": None})

    # Asking for an image slide "where a visual helps" gets one only about half
    # the time — same lesson as the divider cap: state it as a rule, enforce it
    # in code. Placed just before the closing (a visual right before the CTA),
    # and only when there's room under max_blocks.
    if images == 0 and len(result) < max_blocks:
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


def deck_size_bounds(slides=None):
    """(min, max) slides: exactly `slides` when the user asked for a number,
    the brief's 10-15 otherwise."""
    if slides:
        return int(slides), int(slides)
    return DECK_MIN_SLIDES, DECK_MAX_SLIDES


def generate_outline(client, brief, spec, models=TEXT_MODELS, style_preamble="", slides=None):
    """Outline of `slides` blocks when given (a content package's `slides`),
    DECK_MIN_SLIDES..DECK_MAX_SLIDES otherwise.

    The upper bound is code-enforced like every other hard rule. The lower one
    cannot be — code has no themes to invent — so a short outline gets ONE
    corrective call saying how many blocks came back and how many are needed,
    and the longer of the two is kept. Never a failure: a deck one slide short
    is better than no deck (iter113 lost a whole run to one malformed block)."""
    low, high = deck_size_bounds(slides)
    menu = describe_for_prompt(spec, SYNTHESIZABLE_TYPES)
    if style_preamble:
        # Template design brief (plan 9.4) — advice for tone/length; every
        # hard limit stays code-enforced below regardless of what it says.
        menu = f"{menu}\n\n{style_preamble}"
    prompt = (
        OUTLINE_PROMPT
        .replace("__MENU__", menu)
        .replace("__ROLES__", ", ".join(OUTLINE_ROLES))
        .replace("__RANGE__", f"РОВНО {low}" if low == high else f"{low}-{high}")
        .replace("__BRIEF__", brief)
    )

    def make_call(text):
        def call(model):
            result = client.chat([{"role": "user", "content": text}], model=model, max_tokens=1600)
            return _validate_outline(extract_json(result["choices"][0]["message"]["content"]))
        return call

    outline = _enforce_outline_rules(call_with_model_fallback(make_call(prompt), models),
                                     max_blocks=high)
    if len(outline) < low:
        note = (f"\n\nВ прошлом ответе было {len(outline)} блоков, а нужно "
                f"{'ровно ' + str(low) if low == high else f'от {low} до {high}'}. "
                "Раскрой бриф подробнее и ответь полным списком блоков.")
        try:
            longer = _enforce_outline_rules(
                call_with_model_fallback(make_call(prompt + note), models), max_blocks=high)
            if len(longer) > len(outline):
                outline = longer
        except Exception:  # noqa: BLE001 — the first outline stands
            pass
    return outline


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


def _reject_unfilled_placeholders(block):
    """Raise (-> retry) when the model left a variable for the presenter to
    fill: «обходится компании в X млн рублей» shipped on a real deck (iter147),
    and a deck that asks the reader to imagine the number says nothing. One
    hit in 1715 texts of 48 decks, so the retry is rare and cheap."""
    found = unfilled_placeholders(*_block_texts(block))
    if found:
        raise ValueError(f"unfilled placeholders in the block: {found!r}")


def _block_texts(block):
    texts = [block.get("title", ""), block.get("subtitle", ""), block.get("image", ""),
             block.get("caption", ""), block.get("left_heading", ""), block.get("right_heading", "")]
    for key in ("bullets", "left_points", "right_points"):
        texts += list(block.get(key, []) or [])
    for pair in block.get("stats", []) or []:
        texts += list(pair) if isinstance(pair, (list, tuple)) else [pair]
    return [t for t in texts if isinstance(t, str)]


def _take_exactly(items, count, what):
    """The first `count` items; too FEW is an error (code cannot invent the
    missing ones, so the caller retries), too MANY is repaired here.

    Rejecting a surplus used to cost whole runs: a one-figure stats slide
    (count=1) got three or four pairs back, five retries in a row did the same,
    and the ValueError took the deck down with it — iter113 and twice in
    iter116. The model ranks what it thinks matters first; the head is the
    answer."""
    if not isinstance(items, list) or len(items) < count:
        got = len(items) if isinstance(items, list) else 0
        raise ValueError(f"expected {count} {what}, got {got}")
    return items[:count]


def _validate_block(block, role, count):
    if not isinstance(block, dict):
        raise ValueError("block is not a dict")
    if not block.get("title"):
        raise ValueError("empty title")
    if role == "bullet_list":
        block["bullets"] = _take_exactly(block.get("bullets", []), count, "bullets")
    elif role == "stats_kpi":
        stats = block.get("stats", [])
        pairs = [s for s in stats if isinstance(s, list) and len(s) == 2] if isinstance(stats, list) else []
        block["stats"] = _take_exactly(pairs, count, "[num, label] pairs")
    elif role == "two_column_comparison":
        block["left_points"] = _take_exactly(block.get("left_points", []), count, "left points")
        block["right_points"] = _take_exactly(block.get("right_points", []), count, "right points")
    elif role == "image_caption":
        if not str(block.get("image", "")).strip():
            raise ValueError("image_caption needs an 'image' description")
    return block


# Text budgets (plan 10б), enforced deterministically AFTER generation rather
# than via validation retries: length overruns are frequent and gradual (a
# retry often returns another overrun, burning calls), while trimming is
# lossless enough — a title cut to its first sentence stays a title. Numeric
# COUNTS keep using retries when there are too FEW items (they can't be
# invented locally); a surplus is cut to the first N (_take_exactly).
MAX_TITLE_WORDS = 9
MAX_BULLET_CHARS = 72
SHORTEN_ITEMS_PROMPT = "shorten_items.v1.txt"


def bullet_char_budget(template_budget):
    """The item budget for the ALREADY-CHOSEN slide, learned from the template's
    own sample text (generator.get_item_char_budget), or the flat default.

    A per-slide budget may only TIGHTEN the flat one, never loosen it: the
    survey templates' body paragraphs are long prose, and their sample text asks
    for 250-300 char bullets, which would be a regression on every roomy
    template. Measured budgets on the real corpus: T-Zh mono and universal
    one-line slots 18, T-Zh study slots 28, prose boxes 70+ (i.e. the flat
    default, unchanged)."""
    if not template_budget:
        return MAX_BULLET_CHARS
    return min(MAX_BULLET_CHARS, int(template_budget))


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


# How far past its budget an item may stay whole: as far as the fitter can still
# set it on the slot's line at its 70% font floor (plan 10г), i.e. 1 / 0.7.
# 1.25 was tried first and left half of the iter129 cases cut: the model's
# overruns run 30-40% («План развития на три месяца», 27 characters against 21).
ITEM_KEEP_WHOLE_RATIO = 1.4


def _keep_whole_ceiling(max_chars):
    # Never past the flat cap: MAX_BULLET_CHARS is the limit for roomy boxes,
    # where there is no slot line for the fitter to shrink onto.
    return min(int(max_chars * ITEM_KEEP_WHOLE_RATIO), max(MAX_BULLET_CHARS, max_chars))


_ITEM_FIELDS = ("bullets", "left_points", "right_points")


def _shorten_overlong_items(client, block, max_chars, models):
    """One corrective call for the items no trim can save: longer than the
    keep-whole ceiling, so they would be cut at a word.

    No cut keeps the sense of a 45-character sentence in a 28-character slot:
    «Наставники проходят короткий [курс подготовки]», «Программой охвачены
    ключевые [команды]» (iter134) — and dropping the dangling adjective only
    moves the stop one word back. The model can write the shorter phrase; code
    cannot. Only the offending items are sent, once; whatever comes back is
    used only where it is a real improvement (not longer than the ceiling, not
    shorter than the cut), and a failed call leaves the items to the trim as
    before — a corrective call must never cost the block (iter117)."""
    ceiling = _keep_whole_ceiling(max_chars)
    over = [(field, i, str(item)) for field in _ITEM_FIELDS
            for i, item in enumerate(block.get(field) or []) if len(str(item)) > ceiling]
    if not over:
        return block
    prompt = (load_prompt(SHORTEN_ITEMS_PROMPT)
              .replace("__MAXCHARS__", str(max_chars))
              .replace("__COUNT__", str(len(over)))
              .replace("__ITEMS__", "\n".join(f"- {text}" for _, _, text in over)))

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=600)
        items = extract_json(result["choices"][0]["message"]["content"])
        if not isinstance(items, list) or len(items) != len(over):
            raise ValueError(f"expected {len(over)} shortened items")
        return [str(x).strip() for x in items]

    try:
        rewritten = call_with_model_fallback(call, models, retries_per_model=2)
    except (ValueError, KeyError):
        return block
    for (field, i, original), text in zip(over, rewritten):
        # Only an improvement is taken: a rewrite shorter than the cut it
        # replaces would say less than the cut does.
        if text and len(_trim_to_budget(original, max_chars)) <= len(text) <= ceiling:
            block[field][i] = text
    return block


def _trim_to_budget(item, max_chars):
    """An over-long item: cut at a CLAUSE boundary within the budget; failing
    that, kept whole while it is at most ITEM_KEEP_WHOLE_RATIO over; only past
    that, cut at a clause boundary within that ceiling or — last resort — at a
    word boundary within the budget, as before.

    A word-boundary cut keeps the grammar of neither half: once the budget
    stopped being one word (iter127) the model's 50-60 character items came
    back as «Снижение текучести кадров среди», «Рост удовлетворённости
    молодых», «Обратная связь каждые» — seven on one deck (iter129). Titles
    are cut at clauses for the same reason (common.phrases.cut_at_clause); a
    slightly smaller font is a lesser defect than a phrase that stops
    mid-thought.

    The word-boundary rule below it still guards the last resort:"""
    if len(item) <= max_chars:
        return item
    clause = cut_at_clause_chars(item, max_chars, dash=False)
    if clause:
        return clause
    ceiling = _keep_whole_ceiling(max_chars)
    if len(item) <= ceiling:
        return item
    # Past the ceiling a word cut has to happen, and it goes at the BUDGET, not
    # the ceiling: a longer cut is not a better one — at 35 characters «Ручной
    # сбор показателей из семи независимых систем» came out «…из семи», a
    # preposition and a numeral without their noun.
    return cut_at_clause_chars(item, ceiling, dash=False) or _trim_at_word(item, max_chars)


def _trim_at_word(item, max_chars):
    """Cut an over-long item at a WORD boundary, or leave it alone.

    The old cut was `item[:max].rsplit(" ", 1)[0]`, which silently returns the
    whole slice when the slice holds no space — i.e. it truncates a long single
    word mid-letter and ships «Клиентоориентирова». Shipping a mutilated word is
    worse than shipping a long one: the fitter can shrink text, but nothing
    downstream can put the letters back. So when there is no word boundary to
    cut at, keep the item whole."""
    if len(item) <= max_chars:
        return item
    # max_chars + 1, so a word that ENDS exactly on the budget survives: the
    # space after it has to be inside the slice for rsplit to see the boundary.
    # Without it «Данные расходились» — 18 characters against a budget of 18 —
    # came back as «Данные», because the slice ended mid-«расходились» and the
    # partial word was then dropped. The extra character is never kept: rsplit
    # cuts back to the last boundary either way.
    head = item[:max_chars + 1]
    if " " not in head:
        return item
    trimmed = head.rsplit(" ", 1)[0].rstrip(".,;: ")
    return _drop_dangling_function_words(trimmed) or item


# Both live in common/phrases.py now: the running-header slot needs the same
# rule, and generator cannot import this module (it would close a cycle through
# template_spec.builder).
_DANGLING_TAIL_WORDS = DANGLING_TAIL_WORDS
_drop_dangling_function_words = drop_dangling_function_words


def _enforce_text_budgets(block, role, max_chars=MAX_BULLET_CHARS):
    title = str(block.get("title") or "")
    if len(title.split()) > MAX_TITLE_WORDS:
        title = _first_sentence(title)
        if len(title.split()) > MAX_TITLE_WORDS:
            # At a clause boundary, or not at all — see common.phrases.
            title = cut_at_clause(title, MAX_TITLE_WORDS)
        block["title"] = title.rstrip(".,;: ")
    for field in ("bullets", "left_points", "right_points"):
        if field in block:
            trimmed = []
            for item in block[field]:
                item = _trim_to_budget(str(item), max_chars)
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
                   style_preamble="", used_stats=None, item_chars=None):
    """One block, one call, one requirement (the exact count) — sized for the
    already-chosen slide, so nothing needs resizing after the fact.
    style_preamble: the template design brief (plan 9.4), advisory only.
    used_stats: (numbers, labels) already shown on earlier stat slides — asked
    for in the prompt AND enforced by validation+retry, because asking alone
    doesn't hold (project rule: hard constraints go in code).
    item_chars: how long one list item may be on the ALREADY-CHOSEN slide,
    learned from the template's own sample text (generator.get_item_char_budget)
    — a 1.93in one-line slot and a full-width prose box are both "a list", and a
    flat budget makes the narrow one unreadable. Same asked-and-enforced
    treatment as used_stats."""
    theme_line = f"{theme}\n{style_preamble}" if style_preamble else theme
    max_chars = bullet_char_budget(item_chars)
    used_nums, used_labels = used_stats or (set(), set())
    prompt = (
        BLOCK_PROMPTS[role]
        .replace("__THEME__", theme_line)
        .replace("__BRIEF__", brief)
        .replace("__COUNT__", str(count or 3))
        .replace("__MAXCHARS__", str(max_chars))
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
            _reject_unfilled_placeholders(block)
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
    block = _shorten_overlong_items(client, block, max_chars, models)
    return _enforce_text_budgets(block, role, max_chars)
