"""An item no trim can save is rewritten by the model, not cut mid-thought.

The two answers below are a real exchange (iter135, 28-character slot of T-Zh
universal). Cut at a word, the first answer's items ship as «Наставник
закрепляется», «Срок программы», «Формат взаимодействия» — the iter134 deck
exactly. One corrective call for the items past the keep-whole ceiling returns
whole phrases; and a dash is not a clause boundary for an item — «Наставники —
опытные сотрудники команд» was being cut back to «Наставники» by our own trim.
"""

import json

from content_parser.two_phase import generate_block

FIRST = {"title": "Устройство программы", "bullets": [
    "Наставник закрепляется за каждым новым сотрудником",
    "Срок программы — первые три месяца работы",
    "Формат взаимодействия: встречи 3 часа в неделю",
    "Наставники выбираются из опытных сотрудников команд",
    "Обязательное обучение наставников перед началом"]}
REWRITE = ["Наставник у каждого нового сотрудника", "Программа действует 3 месяца",
           "Встречи с наставником — 3 часа в неделю", "Наставники — опытные сотрудники команд",
           "Обязательное обучение наставников"]


class Scripted:
    def __init__(self, *answers):
        self.answers = [json.dumps(a, ensure_ascii=False) for a in answers]
        self.prompts = []

    def chat(self, messages, model=None, **kwargs):
        self.prompts.append(messages[-1]["content"])
        return {"choices": [{"message": {"content": self.answers.pop(0)}}]}


def _bullets(client):
    block = generate_block(client, "bullet_list", "устройство программы",
                           "бриф: программа 3 месяца, встречи 3 часа в неделю", count=5,
                           models=["m"], item_chars=28)
    return [b.replace("\xa0", " ") for b in block["bullets"]]


def test_overlong_items_are_rewritten_whole():
    client = Scripted(FIRST, REWRITE)
    assert _bullets(client) == REWRITE
    assert len(client.prompts) == 2 and "28" in client.prompts[1]


def test_a_failed_rewrite_leaves_the_trim():
    client = Scripted(FIRST, {"not": "a list"}, {"still": "not"})
    got = _bullets(client)
    assert got[0] == "Наставник закрепляется", got


def test_an_item_within_the_ceiling_keeps_what_follows_its_dash():
    first = dict(FIRST, bullets=["Без наставников — дольше адаптация"] + FIRST["bullets"][1:])
    got = _bullets(Scripted(first, REWRITE[1:]))
    assert got[0] == "Без наставников — дольше адаптация", got
