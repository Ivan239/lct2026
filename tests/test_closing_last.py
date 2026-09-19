"""A deck ends once, at the end.

_enforce_outline_rules appended a closing only when there was none. A closing
in the middle stayed there — and the forced image slide was then appended
after the last block, i.e. after «Спасибо» (the iter103 loop deck); two
closings both shipped. The last closing the model wrote is kept and moved to
the end; the blocks around it keep their order.
"""

import json

from content_parser.two_phase import generate_outline

SPEC = {"families": [
    {"id": "f0", "role": "bullet_list", "slides": [{"idx": 1, "capacity": 4}]},
    {"id": "f1", "role": "title", "slides": [{"idx": 0, "capacity": None}]},
]}


class FakeClient:
    def __init__(self, outline):
        self.answer = json.dumps(outline, ensure_ascii=False)

    def chat(self, messages, model=None, **kwargs):
        return {"choices": [{"message": {"content": self.answer}}]}


def _item(role, theme):
    return {"role": role, "theme": theme, "count": 3 if role in ("bullet_list", "stats_kpi") else None}


def test_mid_deck_closing_moves_to_the_end():
    outline = [_item("title", "обложка"), _item("bullet_list", "проблема"),
               _item("closing", "запросить решение"), _item("stats_kpi", "результаты"),
               _item("bullet_list", "план")]
    got = generate_outline(FakeClient(outline), "бриф", SPEC, models=["m"], slides=6)
    roles = [b["role"] for b in got]
    assert roles[-1] == "closing" and roles.count("closing") == 1, roles
    assert roles.index("image_caption") < roles.index("closing"), roles
    assert [b["theme"] for b in got if b["role"] != "image_caption"] == \
        ["обложка", "проблема", "результаты", "план", "запросить решение"]


def test_two_closings_become_one():
    outline = [_item("title", "обложка"), _item("closing", "спасибо"),
               _item("stats_kpi", "результаты"), _item("closing", "утвердить пилот")]
    got = generate_outline(FakeClient(outline), "бриф", SPEC, models=["m"], slides=4)
    closings = [b["theme"] for b in got if b["role"] == "closing"]
    assert closings == ["утвердить пилот"] and got[-1]["role"] == "closing", got
