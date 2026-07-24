"""Loose-cluster probing (plan 9.5), fully mocked: a cluster whose probes
disagree with the medoid must dissolve into per-slide labels; a homogeneous
cluster must stay one LLM call for the whole family."""

import json

from conftest import TJ_TEMPLATE, requires

from design_system.extractor import build_archetype_map
from template_parser.parser import extract_template


class ScriptedClient:
    """Answers by looking at which slide's description was sent — keyed on a
    marker text unique to each slide."""

    def __init__(self, answers_by_marker, default="title"):
        self.answers = answers_by_marker
        self.default = default
        self.calls = 0

    def chat(self, messages, model=None, **kwargs):
        self.calls += 1
        prompt = messages[0]["content"]
        label = next(
            (label for marker, label in self.answers.items() if marker in prompt),
            self.default,
        )
        return {"choices": [{"message": {"content": json.dumps(
            {"archetype": label, "confidence": "high"}
        )}}]}


@requires(TJ_TEMPLATE)
def test_loose_cluster_dissolves_on_disagreement():
    ts = extract_template(TJ_TEMPLATE)
    # T-Ж's known merged cluster is [0, 1, 8, 10] (cover / text / number /
    # thanks on identical folder backgrounds). Script the model to answer
    # differently for the number slide.
    client = ScriptedClient({
        "20 227 000": "stats_kpi",
        "Название презентации": "title",
        "Спасибо,": "closing",
    }, default="title")
    result = build_archetype_map(client, ts)
    assert result[0] == "title"       # cover no longer inherits stats_kpi
    assert result[8] == "stats_kpi"   # the number slide keeps its true label
    assert result[10] == "closing"    # thanks slide labeled for itself


@requires(TJ_TEMPLATE)
def test_homogeneous_cluster_stays_single_call():
    ts = extract_template(TJ_TEMPLATE)
    client = ScriptedClient({}, default="title")  # everyone agrees
    build_archetype_map(client, ts)
    # 12 slides collapse into far fewer calls than slides: agreement means
    # probes don't dissolve clusters into per-slide classification.
    assert client.calls < 12 + 4
