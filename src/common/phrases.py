"""Phrase hygiene shared by the content and generation sides.

Both trim text to fit — bullets to a slot's character budget, the running header
to its slot — and both used to leave the same wreckage behind: a phrase ending
on a preposition. The rule lives here because generator and content_parser
cannot import each other (generator -> two_phase -> template_spec.builder ->
generator is a cycle).
"""

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
