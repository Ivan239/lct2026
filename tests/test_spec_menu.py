"""describe_for_prompt must advertise exactly what the outline validator will
accept: native families outside the offered set are omitted, offered roles the
template lacks are listed as synthesizable."""

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
    assert "image_caption" not in menu  # validator would reject it — don't bait the model
    assert "можно создать с нуля" in menu
    assert "closing" in menu  # offered but not native -> synthesizable
