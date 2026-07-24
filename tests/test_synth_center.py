"""Synthesized slides must not hang their content right under the title with
the lower half of the slide empty (deck f78f14b1, slide 6: a two-column
comparison occupying the top third — the same sparse-slide defect class that
_center_if_underfilled fixed for native fills, but for slides we build from
scratch). Short content is vertically centered between the title and the
bottom bound; content tall enough to fill the area stays where it was."""

from conftest import SURVEY_31, requires

from pptx import Presentation
from pptx.util import Inches

from common.synthesis import SYNTHESIZE
from generator.generator import generate

TWO_COL = {
    "type": "two_column_comparison",
    "title": "Автоматизация против ручной работы",
    "left_heading": "Традиционный подход",
    "left_points": ["Долгая ручная вёрстка", "Неоднородность дизайна", "Зависимость от дизайнеров"],
    "right_heading": "Подход СлайдоГена",
    "right_points": ["Автоматическая генерация", "Сохранение фирменного стиля", "Независимость от человека"],
}

STATS = {
    "type": "stats_kpi",
    "title": "Результаты пилота",
    "stats": [("x10", "ускорение подготовки"), ("0 ч", "ручной вёрстки")],
}

SHORT_BULLETS = {
    "type": "bullet_list",
    "title": "Решение",
    "bullets": ["Разбор шаблона", "Извлечение дизайн-системы", "Сборка презентации"],
}

LONG_BULLETS = {
    "type": "bullet_list",
    "title": "Очень плотный слайд",
    "bullets": [
        f"Пункт номер {i}: развёрнутое описание с достаточно длинным текстом, "
        "чтобы строка переносилась и занимала заметную высоту на слайде"
        for i in range(1, 13)
    ],
}

PLAN = [
    (TWO_COL, SYNTHESIZE),
    (STATS, SYNTHESIZE),
    (SHORT_BULLETS, SYNTHESIZE),
    (LONG_BULLETS, SYNTHESIZE),
]


def _boxes(slide):
    """Shapes in creation order: synthesizers add the title box first."""
    return [s for s in slide.shapes if s.has_text_frame]


@requires(SURVEY_31)
def test_synthesized_content_is_vertically_centered(tmp_path):
    out = str(tmp_path / "synth.pptx")
    generate(SURVEY_31, PLAN, out)
    prs = Presentation(out)

    # Two-column: both columns share a top, and short content moved well below
    # its old position glued to the title.
    title, col_a, col_b = _boxes(prs.slides[0])[:3]
    assert col_a.top == col_b.top
    glued_top = title.top + title.height + Inches(0.3)
    assert col_a.top > glued_top + Inches(0.5)

    # Stats: number boxes share the shifted top.
    stat_boxes = _boxes(prs.slides[1])[1:]
    num_a, num_b = stat_boxes[0], stat_boxes[2]
    assert num_a.top == num_b.top
    stats_glued = _boxes(prs.slides[1])[0].top + _boxes(prs.slides[1])[0].height + Inches(0.35)
    assert num_a.top > stats_glued + Inches(0.5)

    # Short bullet list: centered too.
    b_title, b_body = _boxes(prs.slides[2])[:2]
    assert b_body.top > b_title.top + b_title.height + Inches(0.25) + Inches(0.5)

    # Tall content: estimated height exceeds the available area — must stay
    # exactly where the non-centering layout put it.
    l_title, l_body = _boxes(prs.slides[3])[:2]
    assert abs(int(l_body.top) - int(l_title.top + l_title.height + Inches(0.25))) <= Inches(0.02)
