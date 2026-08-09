"""When the contrast check stays silent, it says WHY.

The universal stats slide scores 91.9 with a heading drawn across the green
sculpture and «пилотного» unreadable. The harness did not miss it silently — it
reported the box as unmeasured — but a bare count says only "the harness was
blind", not which kind of blindness, and the two kinds call for different
answers: text lying on artwork is a layout problem, glyphs that do not separate
from their background is a colour one.

Geometric detection of the underlying defect was measured and rejected again
this iteration: on the very slide where our heading is unreadable, the
DESIGNER's own «20 227 000» overlaps the same artwork by 0.70 — the fraction
cannot tell a number sitting on the dark part from a heading crossing the bright
one. What separates them is contrast under the glyphs, which is exactly the
measurement that comes back blind here.
"""

from evaluation.contrast import (
    UNMEASURED_NO_GLYPHS,
    UNMEASURED_ON_ARTWORK,
    UNMEASURED_TOO_SMALL,
    _box_contrast,
)


def test_a_tiny_box_says_so():
    reasons = []
    assert _box_contrast([(0, 0, 0)] * 3, reasons=reasons) is None
    assert reasons == [UNMEASURED_TOO_SMALL]


def test_a_box_on_artwork_says_so():
    """No colour holds a majority — the modal shade is just the picture's most
    common one, which is how a photo-backed box used to score a healthy 7.40."""
    pixels = [(r * 7 % 256, r * 13 % 256, r * 29 % 256) for r in range(4000)]
    reasons = []
    assert _box_contrast(pixels, reasons=reasons) is None
    assert reasons == [UNMEASURED_ON_ARTWORK]


def test_a_blank_box_says_no_glyphs():
    reasons = []
    assert _box_contrast([(255, 255, 255)] * 4000, reasons=reasons) is None
    assert reasons == [UNMEASURED_NO_GLYPHS]


def test_a_measurable_box_reports_nothing():
    pixels = [(255, 255, 255)] * 3600 + [(0, 0, 0)] * 400
    reasons = []
    result = _box_contrast(pixels, reasons=reasons)
    assert result is not None and reasons == []
    ratio, ink, bg = result
    assert ratio > 10 and ink == (0, 0, 0)


def test_the_reasons_reach_the_report():
    """The per-slide list the probe prints comes from evaluate_boxed_contrast."""
    import inspect

    from evaluation import contrast

    source = inspect.getsource(contrast.evaluate_boxed_contrast)
    assert "unmeasured_why" in source
