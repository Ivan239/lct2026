"""The probe's plan builder is public, and it can repeat roles.

Every ad-hoc script in this project that assembled a plan by hand got it wrong
the same way: iter42 ignored the plan's `count` and read a capacity fix as a
no-op, iter56 skipped the character budgets, and iter85 skipped them again and
read the renderer's own autofit shrink as a product defect — three times the
same class of mistake, each costing an iteration. build_plan is the one path,
so the correct way is now also the easy way.

repeat_outline covers the axis no content preset reaches: a role used twice
sends the matcher back to the same template slide, i.e. the CLONE path, where
this project's nastiest bugs lived (orphaned sldId entries and a notesSlide
claimed by several slides, both of which made PowerPoint demand repair).
"""

import importlib.util
import os

import pytest

from conftest import ROOT, TJ_TEMPLATE, requires

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)


def test_repeat_multiplies_only_content_roles():
    out = probe.repeat_outline(probe.OUTLINE, 3)
    roles = [item["role"] for item in out]
    assert roles.count("title") == 1, "the cover must not repeat"
    assert roles.count("closing") == 1, "the closing must not repeat"
    assert roles.count("bullet_list") == 3
    assert roles.count("stats_kpi") == 3


def test_repeat_of_one_changes_nothing():
    assert probe.repeat_outline(probe.OUTLINE, 1) == probe.OUTLINE


@requires(TJ_TEMPLATE)
def test_build_plan_applies_count_and_budget():
    """The two things hand-rolled plans keep forgetting, asserted together."""
    from design_system.style_profile import build_measured_profile
    from template_spec.builder import build_spec

    archetypes, pngs = probe._cached_parse("custom_f496182bb15f42bb")
    if archetypes is None:
        pytest.skip("no cached parse for the study template")
    profile = build_measured_profile(pngs, archetypes) if pngs else None
    spec_ = build_spec(TJ_TEMPLATE, archetypes, style_profile=profile)

    plan = probe.build_plan(spec_, probe.OUTLINE, probe.LONG)
    bullets = next(b for b, _ in plan if b.get("type") == "bullet_list")

    original = probe.LONG["bullet_list"]["bullets"]
    assert len(bullets["bullets"]) <= len(original), "count was ignored"
    assert any(len(item) < len(orig) for item, orig in zip(bullets["bullets"], original)), \
        "no bullet was trimmed: the character budget was not applied"


@requires(TJ_TEMPLATE)
def test_a_repeated_role_reuses_a_template_slide():
    """The point of the axis: the same slide index appears more than once, which
    is what makes the generator clone it."""
    from design_system.style_profile import build_measured_profile
    from template_spec.builder import build_spec

    archetypes, pngs = probe._cached_parse("custom_f496182bb15f42bb")
    if archetypes is None:
        pytest.skip("no cached parse for the study template")
    profile = build_measured_profile(pngs, archetypes) if pngs else None
    spec_ = build_spec(TJ_TEMPLATE, archetypes, style_profile=profile)

    plan = probe.build_plan(spec_, probe.repeat_outline(probe.OUTLINE, 3), probe.LONG)
    indices = [idx for _, idx in plan]
    assert len(indices) > len(set(indices)), f"no slide is reused: {indices}"
