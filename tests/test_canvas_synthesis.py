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
