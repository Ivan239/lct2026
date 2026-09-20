"""The mirror margin comes from the DECK, not from one slide's boxes.

keep_text_inside_slide narrows a box that runs past the right edge to the
margin the deck's own text keeps. Taking that margin from the CURRENT slide
breaks on a layout that lives in the right half: VK Tech's tag slide starts
both its boxes at 5.72in, so the mirror came out at 4.28in — left of the box
itself — and a title running 1.32in off the slide was skipped (iter150, the
harness flagged 9.1). Measured over the corpus (iter151): the deck margin is
0.07-0.68in on all 13 templates, while a single slide's smallest left reaches
3.67in.
"""

import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation

from generator.generator import generate

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
TAGS = 26  # «бирки» 01-04: оба бокса начинаются с 5.72in
BLOCK = {"type": "bullet_list",
         "title": "Утвердить перенос срока отключения старого контура",
         "bullets": ["Первый пункт", "Второй пункт", "Третий пункт", "Четвёртый пункт"]}


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_no_text_box_runs_past_the_right_edge(tmp_path):
    out = str(tmp_path / "deck.pptx")
    generate(VK_TECH, [(BLOCK, TAGS)], out)
    prs = Presentation(out)
    over = [(s.shape_id, s.text_frame.text[:30], int(s.left + s.width))
            for s in prs.slides[0].shapes
            if s.has_text_frame and s.text_frame.text.strip()
            and s.left is not None and s.width and s.left + s.width > prs.slide_width]
    assert not over, over


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_the_box_is_still_wide_enough_to_read(tmp_path):
    """Narrowed to the deck margin, not to nothing: the 40% floor still holds."""
    out = str(tmp_path / "deck.pptx")
    generate(VK_TECH, [(BLOCK, TAGS)], out)
    template = {s.shape_id: int(s.width) for s in Presentation(VK_TECH).slides[TAGS].shapes}
    for shape in Presentation(out).slides[0].shapes:
        if not shape.has_text_frame or not shape.text_frame.text.strip():
            continue
        was = template.get(shape.shape_id)
        if was:
            assert int(shape.width) >= 0.4 * was, (shape.shape_id, was, int(shape.width))
