"""A short outline is completed by asking for the MISSING blocks.

Re-asking for the whole outline gets the same length back — three decks in a
row shipped 9 of 12, 9 of 10 and 10 of 12 after the corrective call
(iter145-148). The follow-up asks a different question: «нужно ещё 3, вот уже
написанные темы», and the new blocks are appended to the ones we have. Never a
failure: whatever comes back, the deck still ships.
"""

import json

import pytest


class _Client:
    """Answers the outline short twice, then returns the missing blocks."""

    def __init__(self, short=9, missing=3):
        self.short, self.missing = short, missing
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        text = messages[-1]["content"]
        self.prompts.append(text)
        if "Придумай ещё" in text:
            blocks = [{"role": "bullet_list", "theme": f"дополнительная тема {i}", "count": 3}
                      for i in range(self.missing)]
        else:
            blocks = ([{"role": "title", "theme": "титул", "count": None}]
                      + [{"role": "bullet_list", "theme": f"тема {i}", "count": 3}
                         for i in range(self.short - 2)]
                      + [{"role": "closing", "theme": "итог", "count": None}])
        return {"choices": [{"message": {"content": json.dumps(blocks, ensure_ascii=False)}}]}


SPEC = {"families": [{"id": "f0", "role": "bullet_list",
                      "slides": [{"idx": 1, "capacity": 3}, {"idx": 2, "capacity": 3}]},
                     {"id": "f1", "role": "title", "slides": [{"idx": 0, "capacity": None}]},
                     {"id": "f2", "role": "closing", "slides": [{"idx": 3, "capacity": None}]}],
        "rotation": None}


def test_outline_reaches_the_requested_size():
    from content_parser.two_phase import generate_outline

    client = _Client(short=9, missing=3)
    outline = generate_outline(client, "бриф", SPEC, models=["m"], slides=12)
    assert len(outline) == 12, [item["role"] for item in outline]
    assert outline[-1]["role"] == "closing", outline[-1]
    themes = [item.get("theme") for item in outline]
    assert len(set(themes)) == len(themes), themes


def test_a_full_outline_asks_nothing_extra():
    from content_parser.two_phase import generate_outline

    client = _Client(short=12, missing=3)
    outline = generate_outline(client, "бриф", SPEC, models=["m"], slides=12)
    assert len(outline) == 12
    assert not any("Придумай ещё" in p for p in client.prompts), client.prompts[-1][:120]


def test_the_deck_ships_even_if_the_follow_up_fails():
    """Best effort: a failing follow-up leaves the short outline alone."""
    from content_parser.two_phase import generate_outline

    class Failing(_Client):
        def chat(self, messages, model=None, **kwargs):
            if "Придумай ещё" in messages[-1]["content"]:
                raise RuntimeError("network")
            return super().chat(messages, model=model, **kwargs)

    outline = generate_outline(Failing(short=9), "бриф", SPEC, models=["m"], slides=12)
    assert 8 <= len(outline) <= 12, len(outline)
