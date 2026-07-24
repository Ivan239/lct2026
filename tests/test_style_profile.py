"""Measured style profile (plan 9.0/9.1): the same generic code must recover
the T-Ж folder rotation (hand-verified ground truth) AND report no-rotation on
single-background decks — no per-template branches anywhere."""

import glob
import json
import os
import re

from conftest import ROOT, SURVEY_31, requires

from design_system.style_profile import (
    build_measured_profile,
    rotation_targets,
)

TJ_ID = "custom_f496182bb15f42bb"
TJ = os.path.join(ROOT, "output", "templates", f"{TJ_ID}.pptx")


def _pngs(tid):
    return sorted(
        glob.glob(os.path.join(ROOT, "output", "rendered", f"{tid}-*.png")),
        key=lambda p: int(re.search(r"-(\d+)\.png", p).group(1)),
    )


def _profile(tid):
    arch_path = os.path.join(ROOT, "output", "parsed", f"{tid}_archetypes.json")
    arch = json.load(open(arch_path)) if os.path.exists(arch_path) else None
    return build_measured_profile(_pngs(tid), arch)


@requires(TJ)
def test_tj_rotation_recovered():
    prof = _profile(TJ_ID)
    # Hand-verified: 4-color folder cycle, full confidence, photo breather at 7,
    # white technical slide (seen once) excluded from the cycle.
    assert len(prof["rotation"]["order"]) == 4
    assert prof["rotation"]["confidence"] == 1.0
    assert 7 in prof["breathers"]
    # bookends: deck opens and closes on the same brand color (yellow).
    assert prof["bookends"]["title"] == prof["bookends"]["closing"]
    assert prof["bookends"]["title"] == prof["rotation"]["order"][0]

    targets = rotation_targets(prof, 8)
    assert len(targets) == 8
    # cycle property: adjacent targets never repeat
    assert all(a != b for a, b in zip(targets, targets[1:]))
    # bookend anchors
    assert targets[0] == prof["bookends"]["title"]
    assert targets[-1] == prof["bookends"]["closing"]


@requires(SURVEY_31)
def test_single_background_deck_degrades_to_no_rotation():
    prof = _profile("custom_30e96c06e2d47ec3")
    assert prof["rotation"]["order"] == []
    assert rotation_targets(prof, 8) == []


def test_preset_no_rotation():
    if not _pngs("template_b_startup"):
        import pytest
        pytest.skip("preset renders missing")
    prof = _profile("template_b_startup")
    assert prof["rotation"]["order"] == []
    assert len(prof["palette"]) == 1
