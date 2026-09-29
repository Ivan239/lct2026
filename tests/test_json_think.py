"""Ответ reasoning-модели (Qwen3) начинается с <think>…</think>. Черновик JSON
внутри рассуждения не должен выдаваться за ответ."""

import pytest

from common.json_utils import extract_json


def test_the_reasoning_draft_is_not_the_answer():
    text = ('<think>Сначала набросаю: {"title": "черновик"} — нет, лучше иначе.</think>\n'
            '{"title": "Итог", "bullets": ["один", "два"]}')
    assert extract_json(text) == {"title": "Итог", "bullets": ["один", "два"]}


def test_a_list_answer_after_reasoning():
    text = '<think>роли: [title, closing]</think>```json\n[{"role": "title"}]\n```'
    assert extract_json(text) == [{"role": "title"}]


def test_an_unclosed_reasoning_has_no_answer():
    with pytest.raises(ValueError):
        extract_json('<think>думаю {"title": "черновик"} и обрываюсь')


def test_a_plain_answer_is_untouched():
    assert extract_json('{"a": 1}') == {"a": 1}
