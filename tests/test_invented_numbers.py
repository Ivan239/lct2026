"""Число на слайде, которого нет в брифе, — выдумка; блок с ней отбраковывается
и перезапрашивается (Приложение 1: «все цифры и факты есть в исходных
материалах?»). Колоды сдачи на Qwen3-32B вышли с «этап 2 — раскатка на 50%»,
«2000 пользователей в неделю», «+175%» — аудит это помечал, но не предотвращал."""

import json

from content_package import sourced_values
from content_parser.two_phase import generate_block, invented_numbers

BRIEF = ("Среднее время поиска 94 секунды, 38% поисков без результата. Пилот на 1200 "
         "сотрудниках: время поиска 31 секунда, без результата 12%.")


class Scripted:
    def __init__(self, *answers):
        self.answers = [json.dumps(a, ensure_ascii=False) for a in answers]
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        self.prompts.append(messages[-1]["content"])
        return {"choices": [{"message": {"content": self.answers.pop(0)}}]}


def test_derived_figures_are_allowed_invented_ones_are_not():
    from content_package import extract_numbers

    allowed = sourced_values(extract_numbers(BRIEF))
    ok = {"title": "Итоги", "bullets": ["Время поиска снизилось на 67%",
                                        "Пустых поисков меньше на 68%", "Этап 2 — отделы"]}
    assert invented_numbers(ok, allowed) == []
    bad = {"title": "План", "bullets": ["Раскатка на 50% пользователей",
                                        "2000 пользователей в неделю"]}
    assert invented_numbers(bad, allowed) == ["50%", "2000"]


def test_a_block_with_an_invented_figure_is_asked_again():
    invented = {"title": "План раскатки", "bullets": ["Этап 1 — 10% пользователей",
                                                     "Этап 2 — 50% пользователей",
                                                     "Этап 3 — все сотрудники"]}
    clean = {"title": "План раскатки", "bullets": ["Этап 1 — пилотные команды",
                                                  "Этап 2 — все отделы",
                                                  "Этап 3 — вся компания"]}
    client = Scripted(invented, clean)
    block = generate_block(client, "bullet_list", "план раскатки", BRIEF, count=3, models=["m"])
    assert len(client.prompts) == 2
    assert [b.replace("\xa0", " ") for b in block["bullets"]] == clean["bullets"]


def test_the_model_is_told_where_numbers_come_from():
    client = Scripted({"title": "Итоги", "bullets": ["Поиск за 31 секунду", "один", "два"]})
    generate_block(client, "bullet_list", "итоги", BRIEF, count=3, models=["m"])
    assert "ТОЛЬКО из брифа" in client.prompts[0]


def test_when_every_retry_invents_the_deck_still_builds():
    """Сеть не падает на всю колоду: после исчерпанных повторов блок берётся,
    и число остаётся видимым в аудите (как было до правила)."""
    invented = {"title": "План", "bullets": ["Раскатка на 50%", "два", "три"]}
    client = Scripted(*([invented] * 6))
    block = generate_block(client, "bullet_list", "план", BRIEF, count=3, models=["m"])
    assert block["bullets"][0].startswith("Раскатка")


def test_a_rich_package_does_not_blind_the_check():
    """Пересчёт всех пар (включая голые ячейки таблицы) на реальном пакете
    разрешал 159 целых из 1..200 — и выдуманные «50%», «+134%» проходили как
    «выводы». Выводы — только из пары чисел одной единицы."""
    from content_package import extract_numbers, load_package, to_brief_text

    package = load_package("samples/content_packages/feature_smart_search")
    allowed = sourced_values(extract_numbers(to_brief_text(package)))
    assert all(v in allowed for v in (67, 68, 26, 3)), "честные выводы 94→31 с и 38→12%"
    assert not any(v in allowed for v in (50, 134, 175, 2000))
    assert sum(1 for i in range(1, 201) if i in allowed) < 60
