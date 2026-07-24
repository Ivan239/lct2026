"""Line-pitch model (text_fit.paragraph_pitch_pt): percent line spacing follows
the renderers' 1.2-em rule, absolute spacing is taken verbatim, and fitting
honors the spacing — measured against real LibreOffice renders (see the
PCT_SPACING_BASE comment in text_fit.py for the calibration data)."""

from pptx.util import Pt

from generator.text_fit import (
    PCT_SPACING_BASE,
    estimate_block_height_in,
    fit_font_size,
    line_height_pt,
    paragraph_pitch_pt,
)


def test_percent_spacing_uses_renderer_rule():
    # 18pt at 150%: LibreOffice renders exactly 18 * 1.5 * 1.2 = 32.4pt/line.
    assert paragraph_pitch_pt(18, line_spacing=1.5) == 18 * 1.5 * PCT_SPACING_BASE


def test_absolute_spacing_taken_verbatim():
    assert paragraph_pitch_pt(18, line_spacing=Pt(30)) == 30.0


def test_no_spacing_falls_back_to_font_line_height():
    assert paragraph_pitch_pt(18, metrics=None, line_spacing=None) == line_height_pt(18)


def test_fit_honors_line_spacing():
    lines = ["Пункт номер один", "Пункт номер два", "Пункт номер три"]
    width, height = 9144000, 914400  # 10in x 1in box
    plain = fit_font_size(lines, width, height, 24)
    spaced = fit_font_size(lines, width, height, 24, line_spacing=1.5)
    # The same text in the same box must fit at a smaller size once 1.5x
    # spacing inflates every line by the renderer's rule.
    assert spaced.pt < plain.pt
    est = estimate_block_height_in(lines, 10.0, spaced.pt, line_spacing=1.5)
    assert est <= 1.0
