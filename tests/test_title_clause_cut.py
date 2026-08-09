"""A title too long for the word budget is cut at a CLAUSE, or not at all.

The flat cut to nine words mangled the sense on the most visible text of the
slide: «Что мешало собирать управленческую отчётность вовремя и без ручной
сверки» shipped as «…и без ручной» — a preposition and an adjective with the
noun they govern dropped. Cutting before the «и» gives «Что мешало собирать
управленческую отчётность вовремя», which is a title.

When no boundary lies inside the limit the title is kept WHOLE. A long title is
readable and now shrinks to fit its band (iter78); a mangled one says something
the deck does not mean. Same reasoning as iter28's rule about never breaking a
word.
"""

from common.phrases import cut_at_clause
from content_parser.two_phase import MAX_TITLE_WORDS, _enforce_text_budgets

WORDY = "Что мешало собирать управленческую отчётность вовремя и без ручной сверки"


def test_the_cut_lands_before_a_conjunction():
    assert cut_at_clause(WORDY, MAX_TITLE_WORDS) == \
        "Что мешало собирать управленческую отчётность вовремя"


def test_the_latest_boundary_inside_the_limit_wins():
    """Two boundaries are available; taking the later one keeps more of the
    title. «…и аналитики, планирования» rather than «…и аналитики»."""
    text = "Единая платформа отчётности и аналитики, планирования и бюджетирования сети"
    out = cut_at_clause(text, 8)
    assert out == "Единая платформа отчётности и аналитики, планирования", out
    assert text.startswith(out.rstrip(",")), "the cut must be a prefix of the title"


def test_a_title_without_a_boundary_is_kept_whole():
    """Nine words of unbroken noun phrase: cutting anywhere changes the meaning,
    so it is left alone and the layout shrinks it instead."""
    text = "Результаты пилотного внедрения в трёх подразделениях за четвёртый квартал года"
    assert cut_at_clause(text, MAX_TITLE_WORDS) == text


def test_a_short_title_is_untouched():
    assert cut_at_clause("Короткий заголовок", MAX_TITLE_WORDS) == "Короткий заголовок"


def test_the_budget_pass_uses_it():
    """End to end through the pipeline's own entry point, which is what the
    generator actually calls."""
    block = _enforce_text_budgets(
        {"type": "bullet_list", "title": WORDY, "bullets": ["Ручной сбор"]}, "bullet_list")
    assert block["title"] == "Что мешало собирать управленческую отчётность вовремя"
    assert "без ручной" not in block["title"]
