"""A synthesized heading reserves room for a wrap it cannot predict.

iter64 established this on the cover: the renderer wraps 2-4% earlier than
fontTools advances predict, so a line measured at 95% of its budget is drawn on
the next line anyway. _add_title — the heading of every OTHER synthesized role —
was still measuring against the full width, and on the universal template's
image slide the estimate said 1.02in where the truth was 1.53in: the picture
frame was placed under a two-line title, and the third line «подразделения» was
drawn straight through its top border.

Asserted against the renderer's own behaviour rather than the box: the box is
what was wrong.
"""

from conftest import TJ_UNIVERSAL, requires

from pptx import Presentation
from pptx.util import Emu

from fonts.metrics import FontResolver
from generator.deck_style import apply_observed_style, observe_deck_style
from generator.generator import _template_title_pt
from generator.layout_bounds import infer_content_bounds
from generator.synthesizer import synthesize_image_caption
from generator.text_fit import SINGLE_LINE_SAFETY, estimate_block_height_in
from template_parser.parser import extract_template, extract_theme

TITLE = "Как выглядит витрина данных в интерфейсе руководителя подразделения"


@requires(TJ_UNIVERSAL)
def test_the_picture_frame_clears_the_wrapped_heading():
    prs = Presentation(TJ_UNIVERSAL)
    theme = apply_observed_style(extract_theme(TJ_UNIVERSAL), observe_deck_style(prs))
    theme["title_pt"] = _template_title_pt(prs)
    resolver = FontResolver(TJ_UNIVERSAL, extract_theme(TJ_UNIVERSAL))

    idx = synthesize_image_caption(
        prs, theme, infer_content_bounds(extract_template(TJ_UNIVERSAL)),
        {"title": TITLE, "image": "экран дашборда"}, resolver=resolver)
    slide = prs.slides[idx]

    heading = next(s for s in slide.shapes
                   if s.has_text_frame and s.text_frame.text.strip() == TITLE)
    frame = next(s for s in slide.shapes
                 if s.has_text_frame and "ИЗОБРАЖЕНИЕ" in s.text_frame.text)

    run = heading.text_frame.paragraphs[0].runs[0]
    metrics = resolver.metrics_for(run.font.name)
    # What the renderer will actually take, not what the box claims.
    drawn_in = estimate_block_height_in(
        TITLE, Emu(heading.width).inches * SINGLE_LINE_SAFETY, run.font.size.pt,
        metrics=metrics)
    # A clearance, not a hair: on HEAD the frame sat 0.12in below the drawn text,
    # which is where its border crosses the descenders of «подразделения» — the
    # collision is visible on the render even though the numbers do not overlap.
    # The synthesizer's own gap under a heading is 0.25in.
    clearance = Emu(frame.top).inches - (Emu(heading.top).inches + drawn_in)
    assert clearance >= 0.25, (
        f"frame clears the drawn heading by only {clearance:.2f}in: frame at "
        f"{Emu(frame.top).inches:.2f}in, heading takes {drawn_in:.2f}in from "
        f"{Emu(heading.top).inches:.2f}in")


@requires(TJ_UNIVERSAL)
def test_a_short_heading_does_not_grow_a_taller_box():
    """The margin must not inflate every heading: a title that fits one line
    keeps a one-line box."""
    prs = Presentation(TJ_UNIVERSAL)
    theme = apply_observed_style(extract_theme(TJ_UNIVERSAL), observe_deck_style(prs))
    theme["title_pt"] = _template_title_pt(prs)
    idx = synthesize_image_caption(
        prs, theme, infer_content_bounds(extract_template(TJ_UNIVERSAL)),
        {"title": "Витрина данных", "image": "экран"},
        resolver=FontResolver(TJ_UNIVERSAL, extract_theme(TJ_UNIVERSAL)))
    heading = next(s for s in prs.slides[idx].shapes
                   if s.has_text_frame and s.text_frame.text.strip() == "Витрина данных")
    size = heading.text_frame.paragraphs[0].runs[0].font.size.pt
    assert Emu(heading.height).inches <= 2.2 * size / 72 + 0.3, (
        f"a one-line heading got {Emu(heading.height).inches:.2f}in at {size}pt")
