"""The render-wrap counter must count what its name says.

It exists to answer one question: how often does OUR text overflow *because* the
renderer wraps a few percent earlier than fontTools predicts (CLAUDE.md). The
first version answered a different one — "estimated text taller than frame" —
and that is a commoner and mostly harmless thing. On the T-Zh universal deck it
printed 15, of which 13 were single-line boxes missing their frame by 0.01in:
the designer's own «2025» at 8pt in a 0.12in box, a «-40%» at 24pt in a 0.36in
one. A designer sizes a box to the cap height, not to the font's full line box,
and no wrap margin can improve a line that never wraps. Same rule as criterion
1.1: "estimate taller than frame" is not a defect measure.

A number that overstates by 15x is read as a defect and chased — iter92 picked
its target from this list. The measure decides what the next iteration works on,
so it has to be honest about which boxes a margin could actually save.
"""

import importlib.util
import os

from conftest import ROOT

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

from fonts.metrics import FontResolver
from generator.text_fit import SINGLE_LINE_SAFETY, _usable_width_in
from template_parser.parser import extract_theme

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe_module)

FONT = "Arial"


def _deck(tmp_path, name, text, width_in, height_in, size_pt):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(width_in), Inches(height_in))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size_pt)
    path = str(tmp_path / name)
    prs.save(path)
    return path, box


def test_a_one_line_box_tighter_than_its_font_is_not_counted(tmp_path):
    """The template's own chrome: 8pt of running text in a 0.12in slot. The
    estimate is 0.13in — taller than the frame, and irrelevant, because there is
    no second line for an earlier wrap to create."""
    path, _ = _deck(tmp_path, "chrome.pptx", "2025", 1.85, 0.12, 8)
    assert probe_module.boxes_over_at_render_wrap(path) == 0


def test_a_line_that_only_the_early_wrap_breaks_is_counted(tmp_path):
    """The other side: text that fits the full width and does not fit the width
    the renderer effectively has. That one IS what a margin would save, so the
    counter must not go quiet everywhere."""
    size_pt = 18
    text = "Ручной сбор показателей из систем"
    # The box is sized FROM the text: hand-picked inches landed 100pt clear of
    # the boundary and the fixture proved nothing about the boundary.
    sized, _ = _deck(tmp_path, "sizing.pptx", text, 3.0, 0.3, size_pt)
    metrics = FontResolver(sized, extract_theme(sized)).metrics_for(FONT)
    drawn = metrics.text_width_pt(text, size_pt)
    insets = 3.0 - _usable_width_in(3.0)  # the frame's default side insets
    width_in = drawn / 72 / ((1 + SINGLE_LINE_SAFETY) / 2) + insets
    probe_deck, box = _deck(tmp_path, "wrap.pptx", text, width_in, 0.3, size_pt)
    budget = _usable_width_in(Emu(box.width).inches) * 72
    assert budget * SINGLE_LINE_SAFETY < drawn <= budget, (
        f"fixture is not on the wrap boundary: {drawn:.0f}pt against {budget:.0f}pt")
    assert probe_module.boxes_over_at_render_wrap(probe_deck) == 1


def test_paragraphs_are_measured_at_their_own_size(tmp_path):
    """A KPI card holds a 24pt number and an 11pt caption. Charging the caption
    the number's size wrapped it into three lines it will never have."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2.2), Inches(0.9))
    frame = box.text_frame
    for text, size in (("-40%", 24), ("времени на подготовку отчётности", 11)):
        para = frame.paragraphs[0] if text.startswith("-") else frame.add_paragraph()
        run = para.add_run()
        run.text = text
        run.font.name = FONT
        run.font.size = Pt(size)
    path = str(tmp_path / "kpi.pptx")
    prs.save(path)
    assert probe_module.boxes_over_at_render_wrap(path) == 0
