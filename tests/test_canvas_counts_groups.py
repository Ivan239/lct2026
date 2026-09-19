"""A slide whose content is cards (groups) is not a bare canvas.

Synthesis draws a missing role on a clone of a «furniture-only» template
slide — the one with the fewest content boxes. The count read only top-level
text boxes, so VK Tech's four-card slide (cards are GROUPS; only its title is
a top-level box) counted 1 and won. Synthesis removes the cards but not the
single backing shape with their cut-outs, and the image skeleton and the two
columns were drawn across four empty card backings on every deck that
needed a synthesized slide (iter105, 111, 114, 117, 120).

A minimal profile (no rotation, no breathers) keeps the test independent of a
cached parse: the choice then rests on the content count alone.
"""

import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation

from common.synthesis import SYNTHESIZE
from evaluation.loop import _synth_canvas_hints
from generator.slide_kit import content_groups

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
PROFILE = {"breathers": set(), "backgrounds": {}, "rotation": {"order": []}, "bookends": {}}


@pytest.mark.skipif(not os.path.exists(VK_TECH), reason=f"fixture deck missing: {VK_TECH}")
def test_the_card_slide_is_not_chosen_as_a_synthesis_canvas():
    prs = Presentation(VK_TECH)
    card_slides = {i for i, s in enumerate(prs.slides) if len(content_groups(s)) >= 3}
    assert card_slides, "fixture: VK Tech's card slide"

    plan = [({"type": "title"}, 0), ({"type": "image_caption"}, SYNTHESIZE),
            ({"type": "two_column_comparison"}, SYNTHESIZE)]
    hints = _synth_canvas_hints(VK_TECH, plan, PROFILE)
    assert hints, "no canvas chosen at all"
    chosen = set(hints.values())
    assert not chosen & card_slides, f"card slide(s) {sorted(chosen & card_slides)} chosen as canvas"
