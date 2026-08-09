"""Our title length against the template's own.

No rule about SIZES can make a title look native when it is twice as long as the
headings the template was designed around. Measured over the corpus: the
designers' medians are 13 / 16 / 17 / 29 characters, ours is 33 on every
template. That gap is why iter101's attempt to keep the designer's type size
turned into a choice between microtype and a wall — universal's wide box
improved at 42pt while mono's closing became three lines of 61pt, and the whole
change was reverted.

Printed only when ours is meaningfully longer (40%), because on the study
template the two are close and there is nothing to report.
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


def _deck(path, titles):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(Inches(10))), Emu(int(Inches(5.62)))
    for text in titles:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(8), Inches(0.9))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(40)
    prs.save(path)
    return path


def test_reports_both_medians(tmp_path):
    ours = _deck(str(tmp_path / "deck.pptx"),
                 ["Что мешало собирать отчётность вовремя",   # 37
                  "Результаты пилотного внедрения"])          # 30
    theirs = _deck(str(tmp_path / "src.pptx"), ["Заголовок слайда", "Название пункта"])
    assert probe.title_length_vs_template(ours, theirs) == (34, 16)


def test_a_deck_without_titles_says_nothing(tmp_path):
    """A measure with nothing to measure must say so rather than report a zero."""
    empty = Presentation()
    empty.slide_width, empty.slide_height = Emu(int(Inches(10))), Emu(int(Inches(5.62)))
    empty.slides.add_slide(empty.slide_layouts[6])
    path = str(tmp_path / "empty.pptx")
    empty.save(path)
    assert probe.title_length_vs_template(path, _deck(str(tmp_path / "s.pptx"), ["Заголовок"])) is None
