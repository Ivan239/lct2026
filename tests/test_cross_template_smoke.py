"""One synthetic plan, every real template — the invariants recent iterations
established, checked TOGETHER for the first time.

Each of those fixes was verified on a single template while the network was
down, so nothing so far proves they hold at once, or that a fix made for one
designer's layout does not break another's. No LLM and no rendering: the plan is
hardcoded and every invariant is read straight off the .pptx geometry, so this
stays fast enough to run on every commit."""

import os

from conftest import PRESET_A, SURVEY_31, TJ_MONO, TJ_TEMPLATE, TJ_UNIVERSAL, requires

import pytest
from pptx import Presentation
from pptx.util import Emu

from common.synthesis import SYNTHESIZE
from generator.generator import TOPIC_SLOT_PROMPTS
from generator.slide_kit import is_chrome_shape
from generator.synthesizer import IMAGE_PLACEHOLDER_NAME

PLAN = [
    ({"type": "title", "title": "Платформа «Поток»", "subtitle": "Единая аналитика продаж"}, SYNTHESIZE),
    ({"type": "bullet_list", "title": "Ключевые возможности",
      "bullets": ["Единый дашборд", "Прогноз спроса", "Автоотчёты"]}, SYNTHESIZE),
    ({"type": "stats_kpi", "title": "Результаты",
      "stats": [["+25%", "Рост конверсии"], ["-30 часов", "Экономия времени"]]}, SYNTHESIZE),
    ({"type": "two_column_comparison", "title": "До и после",
      "left_heading": "Разрозненность данных", "left_points": ["Данные в разных системах"],
      "right_heading": "Единая картина", "right_points": ["Всё в одном дашборде"]}, SYNTHESIZE),
    ({"type": "image_caption", "title": "Интерфейс дашборда",
      "image": "Экран ноутбука с графиками продаж"}, SYNTHESIZE),
    ({"type": "closing", "title": "Запустите пилот", "subtitle": "Свяжитесь с менеджером"}, SYNTHESIZE),
]


def _generate(template, tmp_path):
    """Every position gets a CLONED CANVAS on purpose. Without synth_canvas the
    synthesizer draws from scratch, the template's chrome never reaches the deck,
    and the page-number and topic-slot invariants below have nothing to look at —
    the first version of this test passed on all five templates precisely because
    it was checking an empty deck."""
    from generator.generator import generate

    source = Presentation(template)
    canvas_count = len(source.slides._sldIdLst)
    hints = {position: position % canvas_count for position in range(len(PLAN))}
    out = str(tmp_path / (os.path.basename(template).replace(".pptx", "") + "_smoke.pptx"))
    generate(template, PLAN, out, synth_canvas=hints)
    return Presentation(out)


@pytest.mark.parametrize("template", [
    pytest.param(TJ_TEMPLATE, id="tj-teaching", marks=requires(TJ_TEMPLATE)),
    pytest.param(TJ_MONO, id="tj-mono", marks=requires(TJ_MONO)),
    pytest.param(TJ_UNIVERSAL, id="tj-universal", marks=requires(TJ_UNIVERSAL)),
    pytest.param(SURVEY_31, id="survey-31", marks=requires(SURVEY_31)),
    pytest.param(PRESET_A, id="preset-a", marks=requires(PRESET_A)),
])
def test_every_template_holds_the_layout_invariants(template, tmp_path):
    prs = _generate(template, tmp_path)
    width, height = prs.slide_width, prs.slide_height
    slides = list(prs.slides)
    assert slides, "generation produced no slides"

    for position, slide in enumerate(slides, start=1):
        for shape in slide.shapes:
            # iter22: a reserved image area that runs off the page renders as a
            # sliver of dashed border at the slide edge.
            if shape.name == IMAGE_PLACEHOLDER_NAME:
                assert shape.top >= 0 and shape.left >= 0, "placeholder starts off-slide"
                assert shape.top + shape.height <= height + Emu(1000), (
                    f"placeholder runs {Emu(shape.top + shape.height - height).inches:.2f}in "
                    f"past the bottom on slide {position}")
                assert shape.left + shape.width <= width + Emu(1000), "placeholder runs off the right"

            if not shape.has_text_frame or not is_chrome_shape(shape, height):
                continue
            runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
            if len(runs) != 1:
                continue
            text = runs[0].text.strip()
            # iter19: a cloned canvas brings the template's own page number.
            # But not every digit box is a page number — the T-Zh universal
            # template footers a brand year, and the first version of the
            # renumbering rewrote "2025" to "0001" on every slide. Page numbers
            # must follow position; anything year-sized must survive verbatim.
            if text.isdigit() and int(text) <= max(len(slides), 99):
                assert int(text) == position, (
                    f"slide {position} shows stale page number {text!r}")

            # iter23: the designer's topic prompt must not ship to the client.
            assert text.lower() not in TOPIC_SLOT_PROMPTS, (
                f"unfilled topic placeholder {text!r} on slide {position}")


@pytest.mark.parametrize("template", [
    pytest.param(TJ_UNIVERSAL, id="tj-universal", marks=requires(TJ_UNIVERSAL)),
])
def test_year_sized_furniture_survives_renumbering(template, tmp_path):
    """The renumbering must not eat furniture that merely looks numeric. The
    T-Zh universal template footers "2025" on ten slides and the first version
    rewrote it to "0001", "0002", ... — invisible to a post-hoc check, because a
    renumbered year is indistinguishable from a real page number. So compare
    against the SOURCE: year-sized digits present in the template must still be
    present in the generated deck."""
    source = Presentation(template)
    src_height = source.slide_height
    years = {
        shape.text_frame.text.strip()
        for slide in source.slides for shape in slide.shapes
        if shape.has_text_frame and is_chrome_shape(shape, src_height)
        and shape.text_frame.text.strip().isdigit()
        and int(shape.text_frame.text.strip()) > 999
    }
    assert years, "premise: this template carries year-sized furniture"

    prs = _generate(template, tmp_path)
    produced = {
        shape.text_frame.text.strip()
        for slide in prs.slides for shape in slide.shapes
        if shape.has_text_frame and shape.text_frame.text.strip()
    }
    for year in years:
        assert year in produced, f"furniture {year!r} was renumbered away"


# Deliberately long Russian compounds: the short strings in PLAN never stress the
# fitters, so an invariant checked against them guards nothing. These are the
# words that actually broke — "Конкурентоспособность" split mid-letter in a 32pt
# synthesized title on two real templates, and a long deck title split inside the
# running header the topic-fill writes.
HARD_PLAN = [
    ({"type": "title", "title": "Высокопроизводительная инфраструктура",
      "subtitle": "Автоматизированное прогнозирование"}, SYNTHESIZE),
    ({"type": "bullet_list", "title": "Конкурентоспособность",
      "bullets": ["Клиентоориентированность", "Стандартизированность"]}, SYNTHESIZE),
    ({"type": "two_column_comparison", "title": "Сопоставление",
      "left_heading": "Несогласованность", "left_points": ["Труднодоступность информации"],
      "right_heading": "Централизованность", "right_points": ["Взаимозаменяемость данных"]}, SYNTHESIZE),
]


@pytest.mark.parametrize("template", [
    pytest.param(TJ_TEMPLATE, id="tj-teaching", marks=requires(TJ_TEMPLATE)),
    pytest.param(TJ_MONO, id="tj-mono", marks=requires(TJ_MONO)),
    pytest.param(TJ_UNIVERSAL, id="tj-universal", marks=requires(TJ_UNIVERSAL)),
    pytest.param(SURVEY_31, id="survey-31", marks=requires(SURVEY_31)),
])
def test_no_word_is_wider_than_its_box(template, tmp_path):
    """No text may be placed at a size where its longest word cannot fit — that
    is precisely when the renderer breaks a word mid-letter."""
    from fonts.metrics import FontResolver
    from generator.generator import generate
    from generator.text_fit import horizontal_margins_in
    from template_parser.parser import extract_theme

    source = Presentation(template)
    canvas_count = len(source.slides._sldIdLst)
    hints = {position: position % canvas_count for position in range(len(HARD_PLAN))}
    out = str(tmp_path / "hard.pptx")
    generate(template, HARD_PLAN, out, synth_canvas=hints)

    prs = Presentation(out)
    resolver = FontResolver(out, extract_theme(out))
    offenders = []
    for position, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.width:
                continue
            budget_pt = max(1.0, (Emu(shape.width).inches - horizontal_margins_in(shape)) * 72)
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if not run.text.strip() or not run.font.size:
                        continue
                    metrics = resolver.metrics_for(run.font.name) if run.font.name else None
                    if metrics is None:
                        continue
                    for word in run.text.split():
                        if metrics.text_width_pt(word, run.font.size.pt) > budget_pt:
                            offenders.append((position, word, round(run.font.size.pt)))
                            break
    assert not offenders, f"words wider than their box (renderer will split them): {offenders[:4]}"


@requires(TJ_UNIVERSAL)
def test_numbering_badges_are_not_used_as_list_slots(tmp_path):
    """A grid of same-size boxes is only a list pattern if its cells can hold
    list text. The T-Zh universal template numbers its items with a row of
    0.26in badges ("01".."06") — geometrically a perfect grid, and filling it
    stuffed each bullet into a 19pt-wide column that rendered as an 8pt stack of
    three-letter fragments. Real list grids in the corpus are 18-31% of the
    slide width; the badge row is 3%."""
    from generator.generator import _find_slot_boxes, generate

    source = Presentation(TJ_UNIVERSAL)
    slide_width = source.slide_width
    for slide in source.slides:
        for box in _find_slot_boxes(slide, set()):
            assert box.width / slide_width >= 0.08, (
                f"slot {Emu(box.width).inches:.2f}in wide is too narrow to hold list text")

    block = {"type": "bullet_list", "title": "Конкурентоспособность",
             "bullets": ["Клиентоориентированность", "Стандартизированность", "Взаимодействие"]}
    out = str(tmp_path / "badges.pptx")
    generate(TJ_UNIVERSAL, [(block, 2)], out)

    prs = Presentation(out)
    placed = [
        shape for shape in list(prs.slides)[0].shapes
        if shape.has_text_frame and "Клиентоориентированность" in shape.text_frame.text.replace("\xad", "")
    ]
    assert placed, "the bullet text vanished"
    assert placed[0].width / prs.slide_width >= 0.08, "bullet landed in a numbering badge"


@pytest.mark.parametrize("template", [TJ_UNIVERSAL, TJ_MONO], ids=["universal", "mono"])
def test_slots_in_one_row_ship_a_single_font_size(template, tmp_path):
    """Sibling slots are fitted one box at a time, so a short item keeps the
    template size while a long neighbour is shrunk — measured on shipped decks,
    one row came out 8/9/13pt and another 9/11/15pt inside IDENTICAL boxes,
    which reads as sloppy precisely because the boxes sit side by side.

    The bullets below are deliberately of very different lengths: that is what
    makes the per-box fitters disagree, and a test with same-length items would
    pass without the harmonizer existing at all."""
    from generator.generator import generate

    block = {"type": "bullet_list", "title": "Конкурентоспособность",
             "bullets": ["Клиентоориентированность", "Стандартизированность", "Взаимодействие"]}
    out = str(tmp_path / "slots.pptx")
    generate(template, [(block, 2)], out)

    slide = list(Presentation(out).slides)[0]
    sizes = {}
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip().replace("\xad", "")
        if text not in block["bullets"]:
            continue
        found = [run.font.size.pt for para in shape.text_frame.paragraphs
                 for run in para.runs if run.font.size and run.text.strip()]
        if found:
            sizes[text] = min(found)

    assert len(sizes) >= 2, f"expected the bullets to land in sibling boxes, got {sizes}"
    assert len(set(sizes.values())) == 1, f"sibling slots ship different sizes: {sizes}"
