"""An over-budget list item is not cut in the middle of a phrase.

Once the item budget stopped being one word (iter127), the model's items came
back 30-40% over it, and the word-boundary cut shipped seven phrases that stop
mid-thought on one VK Tech deck: «Снижение текучести кадров среди», «Рост
удовлетворённости молодых», «План развития на три», «Обратная связь каждые»
(iter129). The items below are what the model wrote there, reconstructed from
the cuts; the budgets are the two slides' own (31 and 21).

Rule: a clause boundary within the budget, else the item whole while the
fitter can still set it at its font floor (1.4x), else a clause boundary within
that, else a word cut at the budget that drops tails needing a noun.
"""

import json

from content_parser.two_phase import generate_block


class FakeClient:
    def __init__(self, response):
        self.response = response

    def chat(self, messages, model=None, **kwargs):
        return {"choices": [{"message": {"content": self.response}}]}


def _bullets(items, budget):
    client = FakeClient(json.dumps({"title": "Преимущества программы", "bullets": items},
                                   ensure_ascii=False))
    block = generate_block(client, "bullet_list", "преимущества", "бриф", count=len(items),
                           models=["fake"], item_chars=budget)
    return [b.replace("\xa0", " ") for b in block["bullets"]]


def test_cards_keep_whole_phrases():
    items = ["Снижение текучести кадров среди новичков", "Быстрое освоение рабочих процессов",
             "Экономия затрат на замену сотрудников", "Повышение эффективности адаптации",
             "Передача опыта и укрепление команды", "Рост удовлетворённости молодых специалистов"]
    got = _bullets(items, 31)
    # within 1.4x of the budget and no clause boundary inside it: whole
    assert got[0] == items[0] and got[5] == items[5], got
    # a clause boundary within the budget is taken
    assert got[4] == "Передача опыта", got


def test_a_forced_cut_does_not_end_on_a_word_that_needs_its_noun():
    got = _bullets(["План развития на три месяца", "Обратная связь каждые две недели"], 21)
    assert got[0] == "План развития на три месяца", got   # 27 <= 21 * 1.4
    assert got[1] == "Обратная связь", got                # 32 > 29: cut, «каждые» dropped
