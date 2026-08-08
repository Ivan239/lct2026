"""Canvas synthesis (plan 9.3): a role the template lacks is synthesized on a
CLONE of a template slide — its background art and chrome survive — instead of
a blank white page. Verified structurally (renders are eyeballed in sweeps)."""

import os

from conftest import TEMPLATES_DIR, requires

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from common.synthesis import SYNTHESIZE
from generator.generator import generate

TJ = os.path.join(TEMPLATES_DIR, "custom_f496182bb15f42bb.pptx")

PLAN = [
    ({"type": "title", "title": "Т", "subtitle": "П"}, 4),
    ({"type": "closing", "title": "Спасибо!", "subtitle": "Пока"}, SYNTHESIZE),
]


@requires(TJ)
def test_canvas_keeps_background_art(tmp_path):
    out = str(tmp_path / "canvas.pptx")
    generate(TJ, PLAN, out, synth_canvas={1: 10})
    prs = Presentation(out)
    assert len(prs.slides) == 2
    synth = prs.slides[1]
    # The T-Ж folder background is a full-bleed PICTURE — it must survive.
    assert any(s.shape_type == MSO_SHAPE_TYPE.PICTURE for s in synth.shapes)
    texts = " ".join(
        s.text_frame.text for s in synth.shapes if s.has_text_frame
    )
    assert "Спасибо!" in texts
    # The canvas's own content text must NOT leak through.
    assert "что вы есть" not in texts


@requires(TJ)
def test_without_canvas_hint_still_works(tmp_path):
    out = str(tmp_path / "no_canvas.pptx")
    generate(TJ, PLAN, out)  # from-scratch path — package gate must still pass
    assert Presentation(out).slides


@requires(TJ)
def test_image_frame_survives_a_narrow_canvas_and_clears_its_chrome():
    """A canvas is REUSED for a different role, so its own text area is not
    automatically a usable content area. T-Zh study slide 10 is one low line and
    hands back a 1.03in band against the template's 4.04in: the title ate it,
    both candidate bands for the image frame came out NEGATIVE, and the slide
    shipped as a heading alone on a full-page empty card. Every T-Zh template
    has three such canvases, and they are exactly the ones offered for synthesis.

    Falling back to the template-wide band then has to respect the furniture the
    clone still carries — first render after the fallback showed the dashed
    frame crossing the footer rule and boxing in «КОММЕНТАРИЙ»."""
    from pptx.util import Emu

    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.slide_kit import is_chrome_shape
    from generator.synthesizer import IMAGE_PLACEHOLDER_NAME, synthesize_image_caption
    from template_parser.parser import extract_template, extract_theme

    prs = Presentation(TJ)
    theme = apply_observed_style(extract_theme(TJ), observe_deck_style(prs))
    bounds = infer_content_bounds(extract_template(TJ))

    idx = synthesize_image_caption(
        prs, theme, bounds,
        {"title": "Платформа в работе", "image": "экран дашборда"},
        canvas_idx=10,
    )
    slide = prs.slides[idx]
    frames = [s for s in slide.shapes if s.name == IMAGE_PLACEHOLDER_NAME]
    assert frames, "the image slide shipped as a bare heading"

    frame = frames[0]
    footers = [
        s for s in slide.shapes
        if s is not frame and s.top is not None and s.height is not None
        and is_chrome_shape(s, prs.slide_height)
        and s.top > prs.slide_height // 2
    ]
    assert footers, "fixture changed: this canvas is supposed to carry footer chrome"
    lowest_allowed = min(int(s.top) for s in footers)
    assert int(frame.top + frame.height) <= lowest_allowed, (
        f"frame runs into the canvas footer: ends at "
        f"{Emu(frame.top + frame.height).inches:.2f}in, footer starts at "
        f"{Emu(lowest_allowed).inches:.2f}in")


@requires(TJ)
def test_canvas_placeholders_are_blanked_but_managed_chrome_is_kept():
    """A clone keeps the whole footer band, and nothing on this path blanked it:
    «КОММЕНТАРИЙ» shipped verbatim under EVERY synthesized slide of this deck.
    The native fill path has always cleared exactly this class of text.

    The split matters as much as the clearing — the page number and the running
    topic slot are owned by later passes (iter31) and must survive, or this fix
    just re-breaks what that one repaired."""
    from generator.deck_style import apply_observed_style, observe_deck_style
    from generator.layout_bounds import infer_content_bounds
    from generator.slide_kit import is_chrome_shape
    from generator.synthesizer import synthesize_image_caption
    from template_parser.parser import extract_template, extract_theme

    prs = Presentation(TJ)
    theme = apply_observed_style(extract_theme(TJ), observe_deck_style(prs))
    bounds = infer_content_bounds(extract_template(TJ))

    # Slide 8 carries all three: the topic slot, a page number and the
    # ownerless «КОММЕНТАРИЙ» — the split this test is about.
    canvas = prs.slides[8]
    before = [
        s.text_frame.text.strip() for s in canvas.shapes
        if s.has_text_frame and is_chrome_shape(s, prs.slide_height)
        and s.text_frame.text.strip()
    ]
    assert any("КОММЕНТАРИЙ" in t for t in before), (
        "fixture changed: this canvas is supposed to carry the placeholder")
    assert any(t.isdigit() for t in before), (
        "fixture changed: this canvas is supposed to carry a page number")

    idx = synthesize_image_caption(
        prs, theme, bounds, {"title": "Платформа", "image": "дашборд"}, canvas_idx=8)
    after = [
        s.text_frame.text.strip() for s in prs.slides[idx].shapes
        if s.has_text_frame and is_chrome_shape(s, prs.slide_height)
        and s.text_frame.text.strip()
    ]
    assert not any("КОММЕНТАРИЙ" in t for t in after), f"placeholder shipped: {after}"
    assert any(t.isdigit() for t in after), f"the page number was wiped too: {after}"
