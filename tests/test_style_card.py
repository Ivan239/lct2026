"""Style card (plan 9.4) — mocked: validation rejects malformed cards and
retries, the disk cache is version-stamped (stale card == no card), and the
prompt preamble renders advice without ever carrying hard limits."""

import json

import pytest

from design_system.style_card import (
    STYLE_CARD_VERSION,
    _validate_card,
    build_style_card,
    card_prompt_preamble,
    load_card,
    save_card,
)


class FakeClient:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = 0

    def chat(self, messages, model=None, **kwargs):
        self.calls += 1
        return {"choices": [{"message": {"content": self.responses.pop(0)}}]}


GOOD = json.dumps({
    "tone": "деловой и лаконичный",
    "metaphor": "цветные папки-разделы",
    "length_norms": {"title": "3-6 слов", "bullet": "4-8 слов"},
    "avoid": ["канцелярит"],
})


def test_retry_on_malformed_then_succeed():
    client = FakeClient(["не json {", GOOD])
    card = build_style_card(client, ["Shapes: 3\n- TEXT_BOX ..."], models=["GigaChat-2"])
    assert client.calls == 2
    assert card["tone"] == "деловой и лаконичный"


def test_missing_tone_rejected():
    with pytest.raises(ValueError):
        _validate_card({"length_norms": {}})


def test_cache_roundtrip_and_version_stamp(tmp_path):
    card = json.loads(GOOD)
    save_card(str(tmp_path), "tpl", card)
    assert load_card(str(tmp_path), "tpl") == card

    # a stale version must read as absent
    path = tmp_path / "tpl_style.json"
    stored = json.loads(path.read_text())
    stored["v"] = STYLE_CARD_VERSION - 1
    path.write_text(json.dumps(stored))
    assert load_card(str(tmp_path), "tpl") is None


def test_preamble_renders_and_empty_for_no_card():
    text = card_prompt_preamble(json.loads(GOOD))
    assert "деловой и лаконичный" in text
    assert "3-6 слов" in text
    assert "канцелярит" in text
    assert card_prompt_preamble(None) == ""
