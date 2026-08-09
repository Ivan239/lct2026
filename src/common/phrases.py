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
