"""Template placeholder text inside a native TABLE is template placeholder text.

A table is a graphicFrame, not a text frame, and the placeholder check walked
text frames only: on WorkSpace a synthesized slide and the closing were drawn
on a table canvas whose «Текст» / «Заголовок» cells shipped untouched under
our text, and five decks scored «заглушек шаблона нет» (iter125, iter131).
"""

from pptx import Presentation
from pptx.util import Inches

from evaluation.deterministic import evaluate, placeholder_hits


def _deck_with_table(path, cell_text):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(6), Inches(1)).text_frame.text = \
        "Одобрите запуск на всех пользователях"
    table = slide.shapes.add_table(3, 2, Inches(1), Inches(2), Inches(6), Inches(3)).table
    for row in table.rows:
        for cell in row.cells:
            cell.text = cell_text
    prs.save(path)
    return path


def test_placeholder_cells_are_found(tmp_path):
    deck = _deck_with_table(str(tmp_path / "t.pptx"), "Текст")
    assert len(placeholder_hits(Presentation(deck).slides[0])) == 6
    assert evaluate(deck)["dop_no_placeholders"]["score"] == 1


def test_real_cells_are_not(tmp_path):
    deck = _deck_with_table(str(tmp_path / "t.pptx"), "18,4 млн ₽")
    assert placeholder_hits(Presentation(deck).slides[0]) == []
