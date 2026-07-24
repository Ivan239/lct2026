"""get_capacity regressions. Two past bugs live here:
1. Title-first shape picking stole the real body on slides whose caption box
   has no explicit run font size -> capacity 1 instead of the real count.
2. Paragraph count was reported without capping by physical icon markers ->
   plans asked for more items than there are icons to point at them."""

from conftest import SURVEY_31, requires

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
