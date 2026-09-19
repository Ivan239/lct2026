"""A placeholder word is not a length sample: the item budget is at least a line.

The per-item budget is learned from the designer's sample text in the slot
(T-Zh mono: «Название пункта» in a one-line 1.93in slot). VK Tech writes
«Пункт» / «Текст» into slots holding 20-28 characters a line; the budget came
out at 12 and every item was cut to one word — «Снижение», «Обязательная»,
«Выход» (iter125, iter126). A sample shorter than one line of its slot says
nothing about length; the line is the limit. Where the sample fills the line,
it still decides.
"""

import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation
from pptx.util import Emu

from fonts.metrics import FontResolver
from generator.generator import _find_slot_boxes, _reference_run
from generator.text_fit import estimate_wrapped_lines, horizontal_margins_in
from template_parser.parser import extract_theme
from template_spec.builder import build_spec

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
MONO = os.path.join(TEMPLATES_DIR, "custom_838830368dac3116.pptx")
ICON_LIST = 45  # four one-line «Пункт» slots, 2.15in wide at 14pt
CONTENTS = 8    # «Содержание»: a four-item list in a 4.22in column
MONO_LIST = 2   # «Название пункта» ×6 in one-line 1.93in slots


def _item_chars(path, idx):
    spec = build_spec(path, {idx: "bullet_list"})
    (entry,) = [s for f in spec["families"] for s in f["slides"] if s["idx"] == idx]
    return entry["item_chars"]


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_placeholder_word_does_not_cut_items_to_one_word():
    budget = _item_chars(VK_TECH, ICON_LIST)
    two_words = "Снижение текучести"
    assert budget >= len(two_words), budget
    assert _item_chars(VK_TECH, CONTENTS) >= len("Ручной сбор показателей")

    # ...and an item that long is still ONE line of its slot, at the slot's size.
    slot = _find_slot_boxes(Presentation(VK_TECH).slides[ICON_LIST], set())[0]
    run = _reference_run(slot)
    metrics = FontResolver(VK_TECH, extract_theme(VK_TECH)).metrics_for(run.font.name)
    item = ("Снижение текучести кадров")[:budget].rstrip()
    assert estimate_wrapped_lines(item, Emu(slot.width).inches, run.font.size.pt,
                                  metrics=metrics, margins_in=horizontal_margins_in(slot)) == 1


@pytest.mark.skipif(not os.path.exists(MONO), reason=f"fixture deck missing: {MONO}")
def test_a_sample_that_fills_its_line_still_decides():
    # 15 chars × 1.2 slack — the calibrating case of get_item_char_budget.
    assert _item_chars(MONO, MONO_LIST) == 18
