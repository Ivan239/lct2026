"""A variable the model left for the presenter is not shipped.

«Потеря 27% новичков обходится компании в X млн рублей ежегодно» went out on a
real deck (iter147): the sentence asks the reader to imagine the number. Both
content paths check it — two_phase and the legacy parser validate
independently (CLAUDE.md) — and both check INSIDE the model call, so the cost
is one retry, never the slide.

Measured over the corpus (iter148): one hit in 1715 texts of our 48 decks —
that very sentence — and one in the 1824 texts of the 13 templates, the
designer's own «x% данные показателя», which never goes through this path.
"""

import json

import pytest


@pytest.mark.parametrize("text", [
    "Потеря 27% новичков обходится компании в X млн рублей ежегодно.",
    "Экономия N% бюджета за квартал",
    "Пилот на [название команды]",
    "Срок: ____",
    "Результат: {метрика}",
])
def test_placeholders_are_found(text):
    from common.phrases import unfilled_placeholders

    assert unfilled_placeholders(text), text


@pytest.mark.parametrize("text", [
    "27% новичков уходят в первые полгода",
    "Топ X продуктов компании",          # буква не перед единицей измерения
    "Сокращение времени релиза на 75%",
    "VK Cloud X series",
    "Наставник тратит 3 часа в неделю",
])
def test_clean_text_is_left_alone(text):
    from common.phrases import unfilled_placeholders

    assert not unfilled_placeholders(text), text


def test_both_content_paths_reject_the_block():
    from content_parser import parser as legacy
    from content_parser import two_phase as tp

    block = {"type": "bullet_list", "title": "Стоимость",
             "bullets": ["обходится в X млн рублей", "второй пункт"]}
    for module in (tp, legacy):
        with pytest.raises(ValueError):
            module._reject_unfilled_placeholders(block)


def test_a_clean_block_passes_both_paths():
    from content_parser import parser as legacy
    from content_parser import two_phase as tp

    block = {"type": "stats_kpi", "title": "Цифры",
             "stats": [["27%", "уходят за полгода"], ["9%", "с наставником"]]}
    for module in (tp, legacy):
        module._reject_unfilled_placeholders(block)


class _Client:
    """First answer carries the variable, second is clean."""

    def __init__(self):
        self.answers = [
            {"title": "Стоимость текучести",
             "bullets": ["Потеря 27% новичков обходится компании в X млн рублей",
                         "Замена разработчика занимает недели",
                         "Наставник тратит 3 часа в неделю"]},
            {"title": "Стоимость текучести",
             "bullets": ["Потеря 27% новичков обходится в 1,1 млн рублей",
                         "Замена разработчика занимает недели",
                         "Наставник тратит 3 часа в неделю"]},
        ]
        self.calls = 0

    def chat(self, messages, model=None, **kwargs):
        block = self.answers[min(self.calls, len(self.answers) - 1)]
        self.calls += 1
        return {"choices": [{"message": {"content": json.dumps(block, ensure_ascii=False)}}]}


def test_the_generated_block_never_keeps_the_variable():
    """The production path, not the helper: on HEAD the first answer ships."""
    from content_parser.two_phase import generate_block

    client = _Client()
    block = generate_block(client, "bullet_list", "текучесть", "бриф", count=3,
                           models=["m1", "m2"], item_chars=90)
    assert client.calls >= 2, client.calls
    assert not any("X млн" in b for b in block["bullets"]), block["bullets"]
