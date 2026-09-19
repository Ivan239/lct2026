"""Small text is measured at the core of its strokes, not at their blend.

At 8pt a stroke is one or two pixels wide and nearly every pixel of it is
blended with the background, so the modal ink colour is lighter than the ink:
VK Tech's team slide declares its «Должность» captions (121,132,146) — about
3.5:1 on the card — and the pixel pass measured 1.88 and flagged the slide
(iter126 took that for a real WCAG failure). The core is the substantial
colour farthest from the background along the same line; only for thin
strokes, and only along that line, so other text in the box is not taken for
it.
"""

import glob
import os

import pytest
from conftest import TEMPLATES_DIR

from pptx import Presentation
from pptx.oxml.ns import qn

from evaluation.contrast import contrast_ratio, evaluate_boxed_contrast

VK_TECH = os.path.join(TEMPLATES_DIR, "vk_tech.pptx")
TEAM_SLIDE = 12
TEAM_RENDER = sorted(glob.glob(os.path.join(
    os.path.dirname(TEMPLATES_DIR), "loop", "*", "vk_tech", "rendered", "vk_tech-13.png")))
DECLARED = (121, 132, 146)


def _one_slide(src, keep, out):
    prs = Presentation(src)
    ids = prs.slides._sldIdLst
    for i, sld_id in reversed(list(enumerate(ids))):
        if i != keep:
            prs.part.drop_rel(sld_id.get(qn("r:id")))
            ids.remove(sld_id)
    prs.save(out)
    return out


@pytest.mark.skipif(not (os.path.exists(VK_TECH) and TEAM_RENDER),
                    reason="fixture deck or its cached render missing")
def test_small_grey_caption_measures_its_declared_colour(tmp_path):
    deck = _one_slide(VK_TECH, TEAM_SLIDE, str(tmp_path / "team.pptx"))
    res = evaluate_boxed_contrast(deck, TEAM_RENDER[:1])
    expected = contrast_ratio(DECLARED, (240, 240, 240))
    assert abs(res["per_slide"][0] - expected) < 0.5, (res["per_slide"], expected)
    assert res["low_contrast_slides"] == [], res["per_slide"]
