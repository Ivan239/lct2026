"""plan_from_outline: capacity-aware slide choice, reuse-instead-of-synthesize
when a family is exhausted, and SYNTHESIZE only for roles the template lacks."""

from common.synthesis import SYNTHESIZE
from matcher.matcher import plan_from_outline

SPEC = {"families": [
    {"id": "f0", "role": "bullet_list", "slides": [
        {"idx": 10, "capacity": 4}, {"idx": 11, "capacity": 8},
    ]},
    {"id": "f1", "role": "title", "slides": [{"idx": 0, "capacity": None}]},
]}


def _assign(outline):
    assignments, skipped = plan_from_outline(outline, SPEC)
    return assignments, skipped


def test_smallest_fitting_capacity_preferred():
    assignments, _ = _assign([{"role": "bullet_list", "theme": "т", "count": 3}])
    (_, idx, count), = assignments
    assert idx == 10  # capacity 4 fits 3; capacity 8 would overpad
    assert count == 4  # generated content must match the slide, not the wish


def test_exhausted_family_reuses_slide_not_synthesize():
    outline = [{"role": "bullet_list", "theme": f"т{i}", "count": 3} for i in range(3)]
    assignments, skipped = _assign(outline)
    indices = [idx for _, idx, _ in assignments]
    assert skipped == []
    assert SYNTHESIZE not in indices
    # Two native slides + one reuse of the best fit.
    assert sorted(indices[:2]) == [10, 11]
    assert indices[2] in (10, 11)


def test_role_absent_from_template_synthesized():
    assignments, skipped = _assign([{"role": "closing", "theme": "финал", "count": None}])
    (_, idx, _), = assignments
    assert idx == SYNTHESIZE
    assert skipped == []


def test_color_rotation_preferred_and_closing_anchored():
    """Plan 9.2: among capacity-equal family members the matcher follows the
    template's discovered color rotation, and the closing position may clone a
    used member to keep the deck's brand bookend."""
    Y, G, B = (250, 240, 190), (170, 230, 200), (190, 230, 250)
    spec = {
        "families": [
            {"id": "f0", "role": "title", "slides": [
                {"idx": 1, "capacity": None, "bg": Y},
                {"idx": 2, "capacity": None, "bg": G},
                {"idx": 3, "capacity": None, "bg": B},
            ]},
        ],
        "rotation": {"order": [Y, G, B], "bookends": {"title": Y, "closing": Y}},
    }
    outline = [{"role": "title", "theme": f"t{i}", "count": None} for i in range(4)]
    assignments, skipped = plan_from_outline(outline, spec)
    assert skipped == []
    indices = [idx for _, idx, _ in assignments]
    # opening on yellow, then rotation, then closing back on yellow via reuse
    assert indices[0] == 1
    assert indices[1] == 2 and indices[2] == 3
    assert indices[3] == 1  # cloned for the closing bookend


def test_no_rotation_spec_unchanged_behavior():
    """A spec without rotation (single-background template) must behave
    exactly as before — capacity first, round-robin reuse."""
    spec = {"families": [
        {"id": "f0", "role": "bullet_list", "slides": [
            {"idx": 10, "capacity": 4, "bg": None}, {"idx": 11, "capacity": 8, "bg": None},
        ]},
    ], "rotation": None}
    outline = [{"role": "bullet_list", "theme": f"т{i}", "count": 3} for i in range(3)]
    assignments, skipped = plan_from_outline(outline, spec)
    assert skipped == []
    assert [idx for _, idx, _ in assignments][:2] == [10, 11]
