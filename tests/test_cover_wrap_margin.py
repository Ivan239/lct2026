"""A synthesized cover/closing sizes its title box for a wrap it cannot predict.

Seen on the render of the T-Zh study deck's closing: «Запустим пилот в вашем
подразделении» was drawn on TWO lines and the subtitle, placed just under a
title the estimate called one line, printed straight through the second one.

Measured on that box: the line is 534.2pt against a 561.6pt budget — 95.1%, i.e.
inside the 5% margin docs/LESSONS.md already declares unpredictable ("рендер переносит
строки на ~2-4% раньше, чем предсказывают fontTools-метрики"). The project's
answer elsewhere is the same constant, SINGLE_LINE_SAFETY; the cover path was
still measuring against the full width.

Assuming the wrap costs nothing when it does not happen: the box ends a little
taller and the subtitle sits a little lower.
"""

from conftest import TJ_TEMPLATE, requires

from pptx import Presentation
from pptx.util import Emu

from fonts.metrics import FontResolver
from generator.deck_style import apply_observed_style, observe_deck_style
from generator.generator import _template_cover_pt
from generator.layout_bounds import infer_content_bounds
from generator.synthesizer import synthesize_title
from template_parser.parser import extract_template, extract_theme

CLOSING = {"title": "Запустим пилот в вашем подразделении",
           "subtitle": "Две недели на подключение и обучение команды"}


def _synthesized(template, data, cover_pt=None):
    prs = Presentation(template)
    theme = apply_observed_style(extract_theme(template), observe_deck_style(prs))
    # 28pt is what the pipeline fitted for this closing (the style profile the
    # full run builds is not reproduced here); at that size the title measures
    # 95.1% of its budget, which is the case under test. _template_cover_pt is
    # the default so the short-title check still exercises the real size.
    theme["cover_pt"] = cover_pt or _template_cover_pt(prs)
    idx = synthesize_title(prs, theme, infer_content_bounds(extract_template(template)), data,
                           resolver=FontResolver(template, extract_theme(template)))
    slide = prs.slides[idx]
    boxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    title = next(s for s in boxes if s.text_frame.text.strip() == data["title"])
    subtitle = next(s for s in boxes if s.text_frame.text.strip() == data["subtitle"])
    return title, subtitle


@requires(TJ_TEMPLATE)
def test_subtitle_clears_a_title_that_wraps_inside_the_unpredictable_margin():
    title, subtitle = _synthesized(TJ_TEMPLATE, CLOSING, cover_pt=28)

    # The premise: this title measures just under its budget, which is exactly
    # where a prediction cannot be trusted. If the template's cover size ever
    # changes so that the title fits comfortably, the test proves nothing.
    from generator.text_fit import SINGLE_LINE_SAFETY, _usable_width_in, horizontal_margins_in

    metrics = FontResolver(TJ_TEMPLATE, extract_theme(TJ_TEMPLATE)).metrics_for(
        next((r.font.name for r in title.text_frame.paragraphs[0].runs if r.font.name), None))
    size_pt = title.text_frame.paragraphs[0].runs[0].font.size.pt
    budget = _usable_width_in(Emu(title.width).inches, horizontal_margins_in(title)) * 72
    line = metrics.text_width_pt(CLOSING["title"], size_pt)
    assert SINGLE_LINE_SAFETY * budget < line <= budget, (
        f"fixture: the line is {line / budget:.1%} of the budget, outside the "
        f"margin this test is about")

    # Against the text the RENDERER will draw, not against the box: on HEAD the
    # subtitle sits below a box sized for one line, and the second line — drawn
    # outside that box — is what it collides with. Asserting on the box alone
    # passes on the broken code.
    from generator.text_fit import estimate_block_height_in

    wrapped_in = estimate_block_height_in(
        CLOSING["title"], Emu(title.width).inches * SINGLE_LINE_SAFETY, size_pt,
        metrics=metrics)
    assert Emu(subtitle.top).inches >= Emu(title.top).inches + wrapped_in, (
        f"subtitle starts at {Emu(subtitle.top).inches:.2f}in; the title starts at "
        f"{Emu(title.top).inches:.2f}in and takes {wrapped_in:.2f}in once wrapped")


@requires(TJ_TEMPLATE)
def test_a_title_well_within_its_width_is_not_padded_out():
    """The safety margin must not inflate every cover: a short title still gets
    a one-line box."""
    short = {"title": "Спасибо", "subtitle": "Вопросы"}
    title, _ = _synthesized(TJ_TEMPLATE, short)
    size_pt = title.text_frame.paragraphs[0].runs[0].font.size.pt
    one_line_in = 2.0 * size_pt / 72
    assert Emu(title.height).inches <= one_line_in, (
        f"a one-word title got a {Emu(title.height).inches:.2f}in box at {size_pt}pt")
