"""Экспорт в .html (ТЗ: три формата). Без LibreOffice: PNG подставляются готовые.

Что обязано держаться: каждый слайд есть и картинкой, и НАСТОЯЩИМ текстом
(включая текст в группах и таблицах — `slide.shapes` их не видит), текст
экранирован, а подмена шрифтов не прячется."""

import base64

from pptx import Presentation
from pptx.util import Inches

from rendering.export import build_html, slide_texts

# 1×1 прозрачный PNG
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=")


def _deck(tmp_path):
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s1 = prs.slides.add_slide(blank)
    s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "Рост <выручки> & прибыли"
    group = s1.shapes.add_group_shape()
    group.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1)).text_frame.text = "Текст в группе"
    s2 = prs.slides.add_slide(blank)
    table = s2.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2)).table
    table.cell(0, 0).text, table.cell(0, 1).text = "Метрика", "Значение"
    table.cell(1, 0).text, table.cell(1, 1).text = "NPS", "62"
    path = tmp_path / "deck.pptx"
    prs.save(path)
    pngs = []
    for i in (1, 2):
        png = tmp_path / f"deck-{i}.png"
        png.write_bytes(PNG)
        pngs.append(str(png))
    return str(path), pngs


def test_text_is_read_inside_groups_and_tables(tmp_path):
    path, _ = _deck(tmp_path)
    texts = slide_texts(path)
    assert "Текст в группе" in texts[0], texts
    assert "NPS | 62" in texts[1], texts


def test_html_has_every_slide_as_image_and_text(tmp_path):
    path, pngs = _deck(tmp_path)
    out = build_html(path, pngs, str(tmp_path / "deck.html"), title="Дека")
    page = open(out, encoding="utf-8").read()
    assert page.count("data:image/png;base64,") == 2
    assert 'id="slide-1"' in page and 'id="slide-2"' in page
    assert "Текст в группе" in page and "NPS | 62" in page
    assert "Рост &lt;выручки&gt; &amp; прибыли" in page, "текст обязан экранироваться"
    assert "<выручки>" not in page


def test_font_substitution_is_named_not_hidden(tmp_path):
    path, pngs = _deck(tmp_path)
    page = open(build_html(path, pngs, str(tmp_path / "a.html"), substituted=["VK Sans"]),
                encoding="utf-8").read()
    assert "VK Sans" in page and "Arial" in page
    clean = open(build_html(path, pngs, str(tmp_path / "b.html")), encoding="utf-8").read()
    assert "заменены" not in clean
