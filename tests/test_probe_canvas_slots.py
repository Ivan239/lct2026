"""How much of a reused canvas we actually fill.

A template slide is a grid of text slots. We fill a slide by picking shapes from
it, and when we pick one and blank the rest, the slide keeps the designer's
proportions and loses the design. Measured on the real decks: T-Zh mono's
numbered-list canvas offers six slots and our bullet_list deck fills two, the
universal one offers five and gets two — and both renders show a heading with a
stray paragraph in an otherwise empty slide.

Nothing else in the harness sees this. Geometry is clean (nothing overflows,
nothing collides), sparse slides are exempt from the evenness criteria by
design, and the deck scores 99.5. This is a new measure, so these tests are its
specification rather than a regression guard; the measure itself was checked
against the two real decks above, by eye, on the renders.
"""

import importlib.util
import os

from conftest import ROOT

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)

SYNTHESIZE = probe.SYNTHESIZE


def _deck(path, texts, chrome=None):
    """One slide, one text box per entry of `texts` ("" means an emptied slot)."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(Inches(10))), Emu(int(Inches(5.62)))
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for i, text in enumerate(texts):
        box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5 + i * 0.8),
                                       Inches(4), Inches(0.6))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(18)
    if chrome:  # a page number in the bottom band, the designer's furniture
        box = slide.shapes.add_textbox(Inches(9.4), Inches(5.3), Inches(0.4), Inches(0.16))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = chrome
        run.font.size = Pt(8)
    prs.save(path)
    return path


def test_reports_the_emptiest_reused_canvas(tmp_path):
    source = _deck(str(tmp_path / "src.pptx"), ["Заголовок", "Один", "Два", "Три"])
    deck = _deck(str(tmp_path / "deck.pptx"), ["Что мешало", "Наш пункт", "", ""])
    assert probe.canvas_slots_used(deck, source, [({}, 0)]) == (2, 4, 1)


def test_a_fully_used_canvas_is_not_reported(tmp_path):
    """The number answers "what did we leave empty" — a canvas we fill (or one we
    add our own shape to, as on survey-31) is not an under-filled canvas, and
    printing "2/1" there would read as a defect where there is none."""
    source = _deck(str(tmp_path / "src.pptx"), ["Заголовок", "Один"])
    deck = _deck(str(tmp_path / "deck.pptx"), ["Что мешало", "Наш пункт", "И подзаголовок"])
    assert probe.canvas_slots_used(deck, source, [({}, 0)]) is None


def test_chrome_counts_on_neither_side(tmp_path):
    """A page number is furniture in both decks; counting it would flatter the
    ratio (3/5 instead of 2/4) exactly where the ratio matters most."""
    source = _deck(str(tmp_path / "src.pptx"), ["Заголовок", "Один", "Два", "Три"], chrome="04")
    deck = _deck(str(tmp_path / "deck.pptx"), ["Что мешало", "Наш пункт", "", ""], chrome="02")
    assert probe.canvas_slots_used(deck, source, [({}, 0)]) == (2, 4, 1)


def test_synthesized_slides_have_no_canvas_to_leave_empty(tmp_path):
    source = _deck(str(tmp_path / "src.pptx"), ["Заголовок", "Один", "Два", "Три"])
    deck = _deck(str(tmp_path / "deck.pptx"), ["Синтез", "", "", ""])
    assert probe.canvas_slots_used(deck, source, [({}, SYNTHESIZE)]) is None
