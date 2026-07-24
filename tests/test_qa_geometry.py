"""qa.geometry.enforce_text_fits: shrinks real overflow, then reaches a fixed
point (idempotent), and never touches shapes whose font metrics can't be
resolved (the char-count estimator's false positives must not "fix" healthy
template text)."""

from conftest import SURVEY_31, requires

from pptx import Presentation
from pptx.util import Pt

from fonts.metrics import FontResolver
from qa.geometry import enforce_text_fits
from template_parser.parser import extract_theme


@requires(SURVEY_31)
def test_shrinks_inflated_text_then_stops():
    prs = Presentation(SURVEY_31)
    resolver = FontResolver(SURVEY_31, extract_theme(SURVEY_31))

    slide_idx = 28
    body = prs.slides[slide_idx].shapes[0]  # the wide 8-paragraph body box
    for p in body.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(60)

    fixes = enforce_text_fits(prs, resolver, slide_indices=[slide_idx])
    assert any(shape_id == body.shape_id for _, shape_id, _, _ in fixes)
    old_new = {sid: (old, new) for _, sid, old, new in fixes}
    old, new = old_new[body.shape_id]
    assert old == 60.0 and new < 60.0

    assert enforce_text_fits(prs, resolver, slide_indices=[slide_idx]) == []


@requires(SURVEY_31)
def test_no_metrics_no_touch():
    prs = Presentation(SURVEY_31)
    fixes = enforce_text_fits(prs, resolver=None)
    assert fixes == []
