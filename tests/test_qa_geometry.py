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


def test_shrinking_stops_at_the_readability_floor(tmp_path):
    """This pass removes overflow by shrinking, and without a floor it shrinks
    into microtext: measured on a real deck, a 49-character bullet in a one-line
    2.86x0.24in slot was fitted at 13pt by the filler and then taken to 8pt
    here, plainly unreadable on the render. The filler already refuses to go
    that far (generator._fit_size_for_shape); this pass was silently undoing it.

    The floor is ABSOLUTE, not a fraction of the current size. A relative floor
    was tried first and broke idempotence — each pass re-floored against its own
    output (60 -> 42 -> 29 on the survey fixture), so a genuinely oversized block
    stopped part-way instead of being fitted."""
    from pptx import Presentation
    from pptx.util import Inches, Pt

    from fonts.metrics import FontResolver
    from qa.geometry import MIN_READABLE_PT, enforce_text_fits
    from template_parser.parser import extract_theme

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(5.63)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(2.86), Inches(0.24))
    box.text_frame.word_wrap = True
    box.text_frame.text = "Ручной сбор показателей из семи независимых систем"
    run = box.text_frame.paragraphs[0].runs[0]
    run.font.size, run.font.name = Pt(13), "Arial"

    path = str(tmp_path / "tiny.pptx")
    prs.save(path)
    prs = Presentation(path)
    resolver = FontResolver(path, extract_theme(path))

    enforce_text_fits(prs, resolver)
    sizes = [r.font.size.pt for s in prs.slides for sh in s.shapes if sh.has_text_frame
             for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
    assert sizes and min(sizes) >= MIN_READABLE_PT, f"shrunk into microtext: {sizes}"

    # And the pass still settles: a second run changes nothing.
    assert enforce_text_fits(prs, resolver) == []
