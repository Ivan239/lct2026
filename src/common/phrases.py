"""Phrase hygiene shared by the content and generation sides.

Both trim text to fit — bullets to a slot's character budget, the running header
to its slot — and both used to leave the same wreckage behind: a phrase ending
on a preposition. The rule lives here because generator and content_parser
cannot import each other (generator -> two_phase -> template_spec.builder ->
generator is a cycle).
"""

import re
# A phrase must not END on a preposition or conjunction. Cutting at a word
# boundary is not enough: the T-Zh study slot budget is 28 characters, and real
# bullets came out as «Ручной сбор показателей из», «Разные форматы выгрузок у»,
# «Согласование занимало до» — each reads as a sentence chopped mid-thought, and
# each shipped to the user that way.
DANGLING_TAIL_WORDS = {
    "и", "а", "но", "или", "что", "как", "чем", "же", "ли", "бы",
    "в", "во", "на", "за", "по", "из", "изо", "с", "со", "к", "ко", "у", "о",
    "об", "обо", "от", "до", "для", "при", "про", "над", "под", "перед", "без",
    "через", "между", "около", "после", "не", "ни",
    "среди", "против", "кроме", "вместо", "вокруг", "внутри",
    # Words that only ever stand BEFORE their noun: a phrase ending on one has
    # lost that noun — «Обратная связь каждые», «План развития на три» (iter129).
    "каждый", "каждая", "каждое", "каждые", "каждого", "каждую", "каждых",
    "весь", "вся", "все", "всех", "несколько", "много", "многих",
    "два", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", "десять",
}


def drop_dangling_function_words(text):
    """Strip trailing prepositions/conjunctions left behind by a trim."""
    words = text.split()
    while words and words[-1].lower().strip(".,;:—-") in DANGLING_TAIL_WORDS:
        words.pop()
    return " ".join(words)


# Words that open a new clause: cutting BEFORE one leaves a phrase that ends on
# its own terms.
CLAUSE_OPENERS = {"и", "а", "но", "или", "чтобы", "потому", "поэтому", "однако", "либо"}


def cut_at_clause_chars(text, max_chars, dash=True):
    """`text` cut at the LATEST clause boundary within max_chars characters, or
    None when there is none. Same boundaries as cut_at_clause: before a clause
    opener, after a comma or — unless dash=False — at a dash.

    A list item passes dash=False: there the dash stands for the verb, and what
    follows it IS the content — «Без наставников — дольше адаптация» came out
    «Без наставников», «Наставники — опытные сотрудники команд» as «Наставники»
    (iter135)."""
    words = text.split()
    best = None
    for i in range(1, len(words)):
        stripped = words[i].lower().strip(".,;:—-")
        if not dash and (words[i] in ("—", "–") or words[i - 1] in ("—", "–")):
            continue
        ends_clause = words[i - 1].endswith(",") or (dash and words[i - 1].endswith("—"))
        if stripped in CLAUSE_OPENERS or ends_clause:
            cut = drop_dangling_function_words(" ".join(words[:i]).rstrip(".,;:—- "))
            if cut and len(cut) <= max_chars:
                best = cut
    return best


def cut_at_clause(text, max_words):
    """`text` shortened to at most max_words, cut at a clause boundary.

    A flat word-count cut mangles the sense: «Что мешало собирать управленческую
    отчётность вовремя и без ручной сверки» came out as «…и без ручной» — a
    preposition and an adjective with the noun they govern dropped — as the
    title of a slide, the most visible text on it. Cutting before the «и»
    instead gives «Что мешало собирать управленческую отчётность вовремя»,
    which is a title.

    When no boundary lies within the limit the text is returned WHOLE: a long
    title is readable and now shrinks to fit (iter78), while a mangled one says
    something the deck does not mean. Same reasoning as iter28's rule about
    never breaking a word.
    """
    words = text.split()
    if len(words) <= max_words:
        return text
    head = words[:max_words]
    for i in range(len(head) - 1, 0, -1):
        stripped = head[i].lower().strip(".,;:—-")
        ends_clause = head[i - 1].endswith(",") or head[i - 1].endswith("—")
        if stripped in CLAUSE_OPENERS or ends_clause:
            cut = " ".join(head[:i]).rstrip(".,;:—- ")
            if cut:
                return drop_dangling_function_words(cut) or cut
    return text


# A variable the model left for the presenter to fill: «обходится компании в X
# млн рублей», «[название команды]», «____». Measured over the corpus: one hit
# in 1715 texts of our 48 decks — that very sentence — and one in 1824 texts of
# the 13 templates, the designer's own «x% данные показателя», which never goes
# through this path. A bare letter only counts right before a unit, so «Топ X
# продуктов» stays out of it and «в X млн» does not.
PLACEHOLDER_VARIABLE = re.compile(
    r"(?<![\w-])[XxNnХх](?![\w-])\s*(?=млн|млрд|тыс|%|руб|₽|раз|дн|недел|месяц|лет|год|чел|шт|мин|час)"
    r"|\[[^\]\n]{1,30}\]|\{[^}\n]{1,30}\}|_{3,}|<[^>\n]{1,30}>",
    re.IGNORECASE)


def unfilled_placeholders(*texts):
    """The placeholder variables left in these texts, if any.

    Both content paths must ask: two_phase and the legacy parser validate
    independently (CLAUDE.md), and a deck that says «в X млн рублей» is worse
    than a deck one retry slower."""
    found = []
    for text in texts:
        for item in (text if isinstance(text, (list, tuple)) else [text]):
            if isinstance(item, (list, tuple)):
                found += unfilled_placeholders(*item)
            elif isinstance(item, str):
                found += [m.group(0).strip() for m in PLACEHOLDER_VARIABLE.finditer(item)]
    return found
