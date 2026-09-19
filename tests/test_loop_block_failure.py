"""One block the model cannot get right costs that slide, not the deck.

Three whole loop runs were lost to one stats block (iter113, iter116 twice):
generate_block raised after its retries and generate_deck let the exception
take the run down. Now the slide is skipped with a reason and the deck is
built from the rest. Network errors are not swallowed — the loop reports those
as «пропущено по сети».
"""

import pytest

import evaluation.loop as loop
from common.synthesis import SYNTHESIZE


@pytest.fixture
def staged(monkeypatch):
    built = {}
    items = [({"role": "title", "theme": "обложка"}, SYNTHESIZE, None),
             ({"role": "stats_kpi", "theme": "итоги"}, SYNTHESIZE, 1),
             ({"role": "closing", "theme": "финал"}, SYNTHESIZE, None)]
    monkeypatch.setattr(loop, "generate_outline", lambda *a, **k: [i for i, _, _ in items])
    monkeypatch.setattr(loop, "plan_from_outline", lambda outline, spec: (items, []))
    monkeypatch.setattr(loop, "_synth_canvas_hints", lambda *a, **k: {})
    monkeypatch.setattr(loop, "generate", lambda src, plan, out, **k: built.setdefault("plan", plan))
    return built


def test_a_block_that_stays_malformed_is_skipped(monkeypatch, staged):
    def block(client, role, theme, brief, **kwargs):
        if role == "stats_kpi":
            raise ValueError("expected 1 [num, label] pairs")
        return {"type": role, "title": theme}

    monkeypatch.setattr(loop, "generate_block", block)
    plan, skipped = loop.generate_deck(None, "GigaChat-2", "t.pptx", {"families": []},
                                       "бриф", "", "out.pptx")
    assert [b["type"] for b, _ in plan] == ["title", "closing"]
    assert skipped == [{"type": "stats_kpi", "title": "итоги",
                        "reason": "блок не собран: expected 1 [num, label] pairs"}]
    assert staged["plan"] == plan


def test_a_network_error_is_not_swallowed(monkeypatch, staged):
    import requests

    def block(client, role, theme, brief, **kwargs):
        raise requests.exceptions.ConnectionError("tunnel down")

    monkeypatch.setattr(loop, "generate_block", block)
    with pytest.raises(requests.exceptions.ConnectionError):
        loop.generate_deck(None, "GigaChat-2", "t.pptx", {"families": []}, "бриф", "", "out.pptx")
