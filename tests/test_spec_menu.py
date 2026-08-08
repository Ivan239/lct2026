"""describe_for_prompt must advertise exactly what the outline validator will
accept: native families outside the offered set are omitted, offered roles the
template lacks are listed as synthesizable."""

from conftest import SURVEY_31, requires

from common.synthesis import SYNTHESIZABLE_TYPES
from template_spec.builder import describe_for_prompt

SPEC = {"families": [
    {"id": "f0", "role": "bullet_list", "slides": [{"idx": 1, "capacity": 4}, {"idx": 2, "capacity": 8}]},
    {"id": "f1", "role": "image_caption", "slides": [{"idx": 0, "capacity": None}]},
]}


def test_menu_filtered_to_offered_roles():
    menu = describe_for_prompt(SPEC, SYNTHESIZABLE_TYPES)
    assert "bullet_list: 2 слайд(ов)" in menu
    assert "ёмкость пунктов: 4, 8" in menu
    assert "image_caption: 1 слайд(ов)" in menu  # now an offered role (skeleton frames)
    assert "можно создать с нуля" in menu
    assert "closing" in menu  # offered but not native -> synthesizable


@requires(SURVEY_31)
def test_chart_built_slides_leave_the_template_with_almost_nothing_to_offer():
    """A characterization test, not an endorsement. 58% of survey-31 (and 62%
    of survey-69) are forced to "other" because they are built around a big
    stale-data picture — a chart whose numbers belong to someone else's context.
    The result on this real customer template is that ONE slide out of 31 is
    offered to the matcher, so a generated deck is almost entirely synthesized
    on cloned canvases.

    That is the intended trade (foreign data must not ride along), but it is a
    different product from a deck reusing the designer's layouts, and it was
    invisible until measured. Two obvious relaxations were checked and rejected:
    using those slides natively ships a hole where the chart was (on slide 4 the
    text ends at 4.41in and the chart owned 4.53-7.50in, and generate() does not
    reflow), and reclassifying them as image_caption gains one slide, because
    the outline rules cap image slides at MAX_IMAGE_SLIDES.

    If a future change unlocks them, this test fails and that decision gets made
    deliberately rather than by accident."""
    import json
    import os

    from template_spec.builder import build_spec

    cache = "output/loop/GigaChat-2-Max/custom_30e96c06e2d47ec3/archetypes.json"
    if not os.path.exists(cache):
        return
    archetypes = {int(k): v for k, v in json.load(open(cache)).items()}
    spec = build_spec(SURVEY_31, archetypes, style_profile=None)

    offered = {s["idx"] for fam in spec["families"] for s in fam["slides"]}
    assert len(offered) <= 2, (
        f"survey-31 now offers {len(offered)} of {len(archetypes)} slides — if that "
        "is intended, check on a render that the freed chart area is not left empty")
