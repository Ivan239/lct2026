"""The offline probe must never leave an artifact that could be read as fresh.

Two ways it used to:

1. A SKIPPED template kept whatever an older run had produced. survey-69 has
   been skipped since iter53 (its cached parses are poisoned), and its deck and
   eval.json from a much older run sat in the output directory for dozens of
   iterations — long enough that a later measurement read them as current.

2. A run that DIED half-way kept the previous deck and renders. That is iter69
   verbatim: a SyntaxError killed the probe, the stale renders were inspected,
   and the conclusion drawn from them ("the change did nothing") was right only
   by accident.

Same rule in both cases: an artifact that may not have been rewritten is not
evidence.
"""

import importlib.util
import os

import pytest

from conftest import ROOT

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe_module)


def _seed(out_root, name):
    stale_dir = os.path.join(out_root, name)
    os.makedirs(stale_dir, exist_ok=True)
    stale = os.path.join(stale_dir, "deck.pptx")
    with open(stale, "w", encoding="utf-8") as f:
        f.write("an older run")
    return stale


def test_a_skipped_template_leaves_nothing_behind(tmp_path):
    out_root = str(tmp_path)
    stale = _seed(out_root, "survey-69")

    result = probe_module.probe("survey-69", "no_such_template_id",
                                probe_module.LONG, out_root)

    assert result.get("skipped"), "fixture: this template is supposed to be skipped"
    assert not os.path.exists(stale), "a skipped template kept an older deck"


def test_a_crashing_run_leaves_nothing_behind(tmp_path, monkeypatch):
    out_root = str(tmp_path)
    name, template_id = probe_module.REAL_TEMPLATES[0]
    if probe_module._cached_parse(template_id)[0] is None:
        pytest.skip(f"no cached parse for {name}")
    stale = _seed(out_root, name)

    def boom(*args, **kwargs):
        raise RuntimeError("simulated crash")

    monkeypatch.setattr(probe_module, "generate", boom)
    with pytest.raises(RuntimeError):
        probe_module.probe(name, template_id, probe_module.LONG, out_root)

    assert not os.path.exists(stale), "a crashed run kept the previous deck"
