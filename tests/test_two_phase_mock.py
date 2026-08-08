"""The LLM-facing layer (content_parser/two_phase.py) exercised with a scripted
fake client — zero GigaChat tokens. Covers exactly the behaviors that were
originally shipped as bug fixes: retry on malformed JSON, code-enforced title
and divider-cap rules, and count validation per role."""

import json

import pytest

from content_parser.two_phase import (
    MAX_SECTION_DIVIDERS,
    _enforce_outline_rules,
    _validate_block,
    generate_block,
    generate_outline,
)


class FakeClient:
    """Returns queued raw response strings, one per chat() call."""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = 0

    def chat(self, messages, model=None, **kwargs):
        self.calls += 1
        if not self.responses:
            raise AssertionError("FakeClient exhausted")
        return {"choices": [{"message": {"content": self.responses.pop(0)}}]}


SPEC = {"families": [
    {"id": "f0", "role": "bullet_list", "slides": [{"idx": 1, "capacity": 4}]},
    {"id": "f1", "role": "title", "slides": [{"idx": 0, "capacity": None}]},
]}


def test_outline_retries_malformed_json_then_succeeds():
    good = json.dumps([
        {"role": "title", "theme": "продукт", "count": None},
        {"role": "bullet_list", "theme": "проблемы", "count": 4},
    ])
    client = FakeClient(["это не json {", good])
    outline = generate_outline(client, "бриф", SPEC, models=["GigaChat"])
    assert client.calls == 2
    # image_caption and the final closing are both enforced in code
    # (_enforce_outline_rules), not something the model was asked for here
    assert [i["role"] for i in outline] == ["title", "bullet_list", "image_caption", "closing"]


def test_outline_title_forced_and_dividers_capped():
    outline = _enforce_outline_rules([
        {"role": "section_divider", "theme": "а", "count": None},
        {"role": "section_divider", "theme": "б", "count": None},
        {"role": "bullet_list", "theme": "в", "count": 3},
    ])
    assert outline[0]["role"] == "title"  # inserted by code, not trusted to the model
    assert sum(1 for i in outline if i["role"] == "section_divider") == MAX_SECTION_DIVIDERS


def test_outline_null_theme_replaced():
    good = json.dumps([{"role": "title", "theme": "null", "count": None}])
    outline = generate_outline(FakeClient([good]), "бриф", SPEC, models=["GigaChat"])
    assert outline[0]["theme"] == "title"


def test_block_count_enforced_via_retry():
    bad = json.dumps({"title": "Заголовок", "bullets": ["один", "два"]})       # count mismatch
    good = json.dumps({"title": "Заголовок", "bullets": ["один", "два", "три"]})
    client = FakeClient([bad, good])
    block = generate_block(client, "bullet_list", "тема", "бриф", count=3, models=["GigaChat"])
    assert client.calls == 2
    assert block["type"] == "bullet_list" and len(block["bullets"]) == 3


def test_stats_pairs_shape_validated():
    with pytest.raises(ValueError):
        _validate_block({"title": "т", "stats": [["x10", "ускорение"], "не пара"]}, "stats_kpi", 2)
    ok = _validate_block({"title": "т", "stats": [["x10", "ускорение"], ["3", "шаблона"]]}, "stats_kpi", 2)
    assert ok["title"] == "т"


def test_missing_title_rejected():
    with pytest.raises(ValueError):
        _validate_block({"bullets": ["a", "b", "c"]}, "bullet_list", 3)


# --- per-slide item budget ---------------------------------------------------

class CapturingClient(FakeClient):
    """Remembers the prompt text of every call."""

    def __init__(self, responses):
        super().__init__(responses)
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        self.prompts.append(" ".join(m.get("content", "") for m in messages))
        return super().chat(messages, model=model, **kwargs)


def test_block_prompt_carries_the_slides_own_item_budget():
    """A 1.93in one-line slot and a full-width prose box are both "a list", and
    the flat 72-char budget makes the narrow one unreadable: a 24-char single
    word cannot wrap (the box is one line tall) and cannot be broken, so the
    fitter drops to its 9pt floor — measured, against the ~14pt the template
    itself renders that label at."""
    good = json.dumps({"title": "Заголовок", "bullets": ["раз", "два", "три"]})
    client = CapturingClient([good])
    generate_block(client, "bullet_list", "тема", "бриф", count=3,
                   models=["GigaChat"], item_chars=18)
    assert "до 18 символов" in client.prompts[0]
    assert "__MAXCHARS__" not in client.prompts[0]


def test_item_budget_only_tightens_never_loosens():
    """The survey templates' prose boxes hand back 250-300 char samples; letting
    those through would ask for LONGER bullets than today on every roomy
    template, which is a regression, not a fix."""
    from content_parser.two_phase import MAX_BULLET_CHARS, bullet_char_budget

    assert bullet_char_budget(18) == 18
    assert bullet_char_budget(292) == MAX_BULLET_CHARS
    assert bullet_char_budget(None) == MAX_BULLET_CHARS


def test_over_budget_item_is_never_cut_mid_word():
    """`item[:max].rsplit(" ", 1)[0]` silently returns the whole slice when the
    slice holds no space, i.e. it shipped «Клиентоориентирова». The fitter can
    shrink a long word; nothing downstream can put its letters back."""
    from content_parser.two_phase import _trim_to_budget

    assert _trim_to_budget("Клиентоориентированность", 18) == "Клиентоориентированность"
    assert _trim_to_budget("Забота о клиенте и партнёре", 18) == "Забота о клиенте"
    assert _trim_to_budget("Коротко", 18) == "Коротко"
