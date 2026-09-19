"""A chart drawn with shapes is someone else's data, like a table.

Its bars carry the designer's numbers in their LENGTHS and nothing resizes
them (backlog item 5). VK Tech's funnel shipped as «до и после пилота» with
our labels beside meaningless bars (iter141, iter143), and its Gantt chart
shipped as a KPI board — «-5 недель» against the designer's own proportions,
one bar with no label at all (iter144). Measured over the corpus (iter146):
exactly three slides match — VK Tech's funnel and Gantt and VK Education's
«Диаграмма Ганта» — and no template loses a slide it was filling well.
"""

import json
import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation

from common.pictures import has_data_object

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
FUNNEL, GANTT, PRODUCT_CARDS = 40, 51, 53


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_the_funnel_and_the_gantt_are_data_objects():
    slides = Presentation(VK_TECH).slides
    assert has_data_object(slides[FUNNEL]), "воронка"
    assert has_data_object(slides[GANTT]), "диаграмма Ганта"


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_a_card_list_is_not():
    """The vertical reading (equal width, varying height) was measured and
    dropped: it found no chart of its own and flagged this slide's four card
    illustrations, 1.69in wide and 1.54-2.26in tall."""
    assert not has_data_object(Presentation(VK_TECH).slides[PRODUCT_CARDS])


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_the_rule_is_narrow():
    """Whole-corpus guard: a chart is rare. If a future tweak starts matching
    card grids or icon rows, this count moves and the measurement is stale."""
    from common.pictures import has_shape_chart

    matches = sum(1 for s in Presentation(VK_TECH).slides if has_shape_chart(s))
    assert matches <= 12, matches
