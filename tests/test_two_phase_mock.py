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
    # slides=4: this test is about the JSON retry, not the deck size
    outline = generate_outline(client, "бриф", SPEC, models=["GigaChat"], slides=4)
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


def test_a_word_ending_exactly_on_the_budget_is_kept():
    """The slice was taken at exactly max_chars, so the space FOLLOWING a word
    that ends on the budget fell outside it: rsplit saw no boundary there, cut
    back to the previous one, and «Данные расходились» — 18 characters against a
    budget of 18 — shipped as «Данные». Seen on the render of the universal
    template's numbered grid, whose cells learn an 18-char budget from the
    designer's own «Название пункта»."""
    from content_parser.two_phase import _trim_to_budget

    assert len("Данные расходились") == 18
    assert _trim_to_budget("Данные расходились между системами", 18) == "Данные расходились"
    # One character less and the boundary genuinely does not fit.
    assert _trim_to_budget("Данные расходились между системами", 17) == "Данные"
    # The extra character is only ever used to SEE the boundary, never kept:
    # a word straddling the budget is still dropped whole.
    assert _trim_to_budget("Данные расходились всюду", 19) == "Данные расходились"


def test_a_trimmed_bullet_does_not_end_on_a_preposition():
    """Cutting at a word boundary is not enough. The T-Zh study slot budget is
    28 characters, and realistic bullets came out as «Ручной сбор показателей
    из», «Разные форматы выгрузок у», «Согласование занимало до» — each reads as
    a sentence chopped mid-thought, and each shipped that way."""
    from content_parser.two_phase import _trim_to_budget

    assert _trim_to_budget("Ручной сбор показателей из семи независимых систем", 28) \
        == "Ручной сбор показателей"
    assert _trim_to_budget("Разные форматы выгрузок у каждого подразделения", 28) \
        == "Разные форматы выгрузок"
    assert _trim_to_budget("Единая витрина данных и отчётности для всех", 28) \
        == "Единая витрина данных"

    # A phrase that already ends on a content word is untouched, and a single
    # long word is still never cut mid-letter (iter28).
    assert _trim_to_budget("Ручной сбор показателей", 28) == "Ручной сбор показателей"
    assert _trim_to_budget("Клиентоориентированность", 18) == "Клиентоориентированность"


def test_the_slot_char_budget_reaches_the_model(monkeypatch, tmp_path):
    """The budget crosses three modules — build_spec puts item_chars in the
    spec, generate_deck looks it up for the slide the matcher chose, and
    generate_block substitutes it into the prompt. iter28 tested only the last
    hop, so a break in either of the first two would have been invisible: the
    model would keep being asked for 72 characters into a 28-character slot."""
    import glob
    import json
    import os
    import re

    from conftest import TJ_TEMPLATE
    from design_system.style_profile import build_measured_profile
    from evaluation.loop import generate_deck
    from template_spec.builder import build_spec

    cache = "output/loop/GigaChat-2/custom_f496182bb15f42bb/archetypes.json"
    if not os.path.exists(TJ_TEMPLATE) or not os.path.exists(cache):
        return
    archetypes = {int(k): v for k, v in json.load(open(cache)).items()}
    pngs = sorted(glob.glob("output/loop/GigaChat-2/custom_f496182bb15f42bb/rendered/*.png"))
    spec = build_spec(TJ_TEMPLATE, archetypes,
                      style_profile=build_measured_profile(pngs, archetypes) if pngs else None)

    seen = []

    class Client:
        def chat(self, messages, model=None, **kwargs):
            prompt = " ".join(m.get("content", "") for m in messages)
            seen.append(prompt)
            count = int(re.search(r"РОВНО (\d+)", prompt).group(1)) if "РОВНО" in prompt else 3
            if "Разбей бриф" in prompt:
                body = json.dumps([{"role": "title", "theme": "t", "count": None},
                                   {"role": "bullet_list", "theme": "b", "count": 3},
                                   {"role": "closing", "theme": "z", "count": None}])
            elif "слайда-списка" in prompt:
                body = json.dumps({"title": "Что мешало",
                                   "bullets": [f"пункт {i}" for i in range(count)]})
            elif "иллюстрацией" in prompt:
                body = json.dumps({"title": "В работе", "image": "экран дашборда"})
            elif "титульного" in prompt:
                body = json.dumps({"title": "Поток", "subtitle": "итоги"})
            else:
                body = json.dumps({"title": "Итог", "subtitle": "давайте"})
            return {"choices": [{"message": {"content": body}}]}

    generate_deck(Client(), "GigaChat-2", TJ_TEMPLATE, spec, "бриф", "",
                  str(tmp_path / "deck.pptx"))

    list_prompts = [p for p in seen if "слайда-списка" in p]
    assert list_prompts, "no bullet-list block was requested"
    assert "до 28 символов" in list_prompts[0], (
        "the slot's own budget did not reach the prompt: "
        + next((l for l in list_prompts[0].splitlines() if "символов" in l), "?"))


class RecordingClient(FakeClient):
    def __init__(self, responses):
        super().__init__(responses)
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        self.prompts.append(messages[-1]["content"])
        return super().chat(messages, model=model, **kwargs)


def _outline(n):
    return json.dumps([{"role": "title", "theme": "т", "count": None}]
                      + [{"role": "bullet_list", "theme": f"б{i}", "count": 3} for i in range(n - 2)]
                      + [{"role": "closing", "theme": "итог", "count": None}])


def test_a_short_outline_gets_one_corrective_call_with_the_count():
    """The brief's floor is 10 slides; the prompt used to ask for 5-9 and every
    loop deck came out 8-9. A short outline is sent back once, saying how many
    blocks came and how many are needed; the longer one is kept."""
    client = RecordingClient([_outline(6), _outline(11)])
    outline = generate_outline(client, "бриф", SPEC, models=["GigaChat"])
    assert client.calls == 2
    assert "10-15" in client.prompts[0]
    assert "было 7 блоков" in client.prompts[1] and "от 10 до 15" in client.prompts[1]
    assert 10 <= len(outline) <= 15


def test_a_failed_correction_keeps_the_first_outline_instead_of_failing():
    """No deck at all is worse than a short one (iter113 lost a whole run to
    one malformed block)."""
    client = RecordingClient([_outline(6)])          # the corrective call finds nothing
    outline = generate_outline(client, "бриф", SPEC, models=["GigaChat"])
    assert len(outline) == 7                          # 6 + the enforced image slide


def test_an_asked_size_is_exact_and_the_cap_follows_it():
    client = RecordingClient([_outline(14)])
    outline = generate_outline(client, "бриф", SPEC, models=["GigaChat"], slides=12)
    assert "РОВНО 12" in client.prompts[0]
    assert len(outline) == 12 and outline[-1]["role"] == "closing"
