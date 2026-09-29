"""Три варианта вёрстки одной колоды (ТЗ, п. 3.6).

Ось различий: плотность (число слайдов), приоритет ролей и порядок
повествования. Различаются ПЛАНЫ, а вёрстка у всех трёх одна — поэтому ни один
вариант не может соблюдать шаблон хуже другого.
"""

import json

import pytest

from content_parser.variants import (DEFAULT_VARIANT, VARIANTS, variant_names,
                                     variant_preamble, variant_slides)


def test_three_variants_exist():
    assert len(variant_names()) == 3, variant_names()
    for name, spec in VARIANTS.items():
        assert spec["title"] and spec["description"], name


def test_density_differs_and_stays_in_the_brief_range():
    """10–15 слайдов — рамка ТЗ; варианты двигаются внутри неё, а не за неё."""
    for order in (None, 10, 12, 15):
        sizes = {name: variant_slides(name, order) for name in variant_names()}
        assert len(set(sizes.values())) >= 2, (order, sizes)
        for name, n in sizes.items():
            assert 10 <= n <= 15, (order, name, n)


def test_the_low_end_of_the_range_does_not_collapse_the_variants():
    """Множитель от заказа схлопывался на границе: 0.7 × 10 обрезалось до 10, и
    «компактный» выходил ровно той длины, что «визуальный» — девять сданных
    колод отличались только перестановкой слайдов. Заказ — центр, варианты
    расходятся вокруг него."""
    at_low = {name: variant_slides(name, 10) for name in variant_names()}
    assert max(at_low.values()) > min(at_low.values()), at_low
    spread = {name: variant_slides(name, None) for name in variant_names()}
    assert len(set(spread.values())) == 3, spread
    assert spread["compact"] < spread["visual"] < spread["detailed"], spread


def test_each_variant_has_its_own_prompt_file():
    """Новые промпты — отдельными файлами, как требует ТЗ."""
    texts = {name: variant_preamble(name) for name in variant_names()}
    assert texts[DEFAULT_VARIANT] == "", "вариант по умолчанию не меняет промпт"
    others = [t for n, t in texts.items() if n != DEFAULT_VARIANT]
    assert all(len(t) > 50 for t in others), {n: len(t) for n, t in texts.items()}
    assert len(set(others)) == len(others), "добавки вариантов не должны совпадать"


class _Client:
    """Возвращает план фиксированной длины и запоминает, что просили."""

    def __init__(self):
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        self.prompts.append(messages[-1]["content"])
        blocks = ([{"role": "title", "theme": "титул", "count": None}]
                  + [{"role": "bullet_list", "theme": f"тема {i}", "count": 3} for i in range(8)]
                  + [{"role": "closing", "theme": "итог", "count": None}])
        return {"choices": [{"message": {"content": json.dumps(blocks, ensure_ascii=False)}}]}


SPEC = {"families": [{"id": "f0", "role": "bullet_list",
                      "slides": [{"idx": 1, "capacity": 3}, {"idx": 2, "capacity": 3}]},
                     {"id": "f1", "role": "title", "slides": [{"idx": 0, "capacity": None}]},
                     {"id": "f2", "role": "closing", "slides": [{"idx": 3, "capacity": None}]}],
        "rotation": None}


@pytest.mark.parametrize("variant", ["compact", "visual"])
def test_the_variant_instruction_reaches_the_model(variant):
    from content_parser.two_phase import generate_outline

    client = _Client()
    generate_outline(client, "бриф", SPEC, models=["m"], slides=10, variant=variant)
    assert client.prompts, "промпт не отправлен"
    assert variant_preamble(variant).strip().splitlines()[0] in client.prompts[0], \
        client.prompts[0][-200:]


def test_default_variant_sends_the_plain_prompt():
    from content_parser.two_phase import generate_outline

    client = _Client()
    generate_outline(client, "бриф", SPEC, models=["m"], slides=10, variant=DEFAULT_VARIANT)
    assert "ВАРИАНТ ВЁРСТКИ" not in client.prompts[0]
