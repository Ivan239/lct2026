"""get_capacity regressions. Two past bugs live here:
1. Title-first shape picking stole the real body on slides whose caption box
   has no explicit run font size -> capacity 1 instead of the real count.
2. Paragraph count was reported without capping by physical icon markers ->
   plans asked for more items than there are icons to point at them."""

from conftest import SURVEY_31, TJ_MONO, TJ_TEMPLATE, requires

from pptx import Presentation

from generator.generator import get_capacity


@requires(SURVEY_31)
def test_bullet_capacity_capped_by_icons():
    prs = Presentation(SURVEY_31)
    # slide 28: 8 paragraphs in the body but only 4 icon markers -> 4.
    assert get_capacity(prs.slides[28], "bullet_list") == 4
    # slide 26: 8 paragraphs, 8 icons -> 8.
    assert get_capacity(prs.slides[26], "bullet_list") == 8
    # slide 24: 7 paragraphs, 7 icons -> 7.
    assert get_capacity(prs.slides[24], "bullet_list") == 7


@requires(SURVEY_31)
def test_stats_single_box_capacity_from_icons():
    prs = Presentation(SURVEY_31)
    # Single-text-block stats slides: capacity = number of icon markers.
    assert get_capacity(prs.slides[1], "stats_kpi") == 4
    assert get_capacity(prs.slides[3], "stats_kpi") == 2


@requires(TJ_MONO)
@requires(TJ_TEMPLATE)
def test_item_char_budget_is_learned_from_the_templates_own_sample():
    """get_capacity answers HOW MANY; this answers HOW LONG. The budget comes
    from the designer's own sample text rather than from chars-per-line, because
    the size those slots inherit is not in the run — it resolves through the
    placeholder/master chain, and the theme routinely lies about the deck's real
    look. Where both can be computed they agree: metrics give ~17 chars for the
    T-Zh mono slot at its rendered size, its sample is 15."""
    from generator.generator import get_item_char_budget

    mono = Presentation(TJ_MONO)
    tight = get_item_char_budget(list(mono.slides)[2], "bullet_list")
    assert tight is not None and 12 <= tight <= 24, f"one-line slot budget is {tight}"

    study = Presentation(TJ_TEMPLATE)
    roomy = get_item_char_budget(list(study.slides)[9], "bullet_list")
    assert roomy is not None and roomy > tight * 2, (
        f"a prose box must not get the same budget as a one-line slot: {roomy} vs {tight}")

    assert get_item_char_budget(list(mono.slides)[2], "stats_kpi") is None
