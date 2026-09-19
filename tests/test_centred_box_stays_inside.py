"""A centred box is slid back inside the slide, not narrowed.

keep_text_inside_slide narrows a LEFT-aligned box that overhangs the right
edge; centred boxes were skipped because narrowing moves the centre the
designer aligned to something. VK Tech's funnel sets its numbers in centred
boxes that overhang by 0.1-0.5in: the designer's own «10» stays inside, our
«Среднее время поиска: 31 секунда» was cut by the slide edge (iter141, the
harness flagged 9.1). The box now slides back inside with its width and its
font intact (iter143).
"""

import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation

from generator.generator import generate

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
FUNNEL = 40
BLOCK = {"type": "two_column_comparison",
         "title": "До и после пилота: эффективность поиска",
         "left_heading": "До пилота", "right_heading": "После пилота",
         "left_points": ["Среднее время поиска: 94 секунды", "Поиски без результата: 38%",
                         "Активных пользователей: 410"],
         "right_points": ["Среднее время поиска: 31 секунда", "Поиски без результата: 12%",
                          "Активных пользователей: 864"]}


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_no_filled_box_hangs_past_the_slide_edge(tmp_path):
    out = str(tmp_path / "deck.pptx")
    generate(VK_TECH, [(BLOCK, FUNNEL)], out)
    prs = Presentation(out)
    slide = prs.slides[0]
    over = [(s.shape_id, s.text_frame.text[:30], s.left + s.width) for s in slide.shapes
            if s.has_text_frame and s.text_frame.text.strip() and s.left is not None and s.width
            and s.left + s.width > prs.slide_width]
    assert not over, over


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_the_centred_box_keeps_its_width(tmp_path):
    """Sliding, not narrowing: a narrowed centred box would re-wrap and shrink
    the text the designer sized for that column."""
    out = str(tmp_path / "deck.pptx")
    generate(VK_TECH, [(BLOCK, FUNNEL)], out)
    template = {s.shape_id: s.width for s in Presentation(VK_TECH).slides[FUNNEL].shapes}
    for shape in Presentation(out).slides[0].shapes:
        if not shape.has_text_frame or not shape.text_frame.text.strip():
            continue
        was = template.get(shape.shape_id)
        if was and shape.text_frame.paragraphs[0].alignment is not None:
            assert shape.width == was, (shape.shape_id, was, shape.width)
