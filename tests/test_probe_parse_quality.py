"""The probe reports how much of the template its cached parse classified.

"из шаблона 1/6" was read as "this template offers one usable slide" for twenty
iterations. It does not: survey-31's best cached parse labels ONE of its 31
slides and calls the other 30 "other", so the matcher has nothing to match and
every content slide gets synthesized. The template itself is fine — its icon
bullet lists and stat slides are visible on its own renders.

The number was never printed, and a measurement nobody prints is a measurement
nobody checks — the same reason `unmeasured_boxes` exists in the contrast
report.
"""

import importlib.util
import os

from conftest import ROOT

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe_module)


def test_classified_counts_only_real_labels():
    assert probe_module._classified({0: "title", 1: "other", 2: "bullet_list"}) == 2
    assert probe_module._classified({0: "other", 1: "other"}) == 0


def test_cached_parse_prefers_the_richer_cache(tmp_path, monkeypatch):
    """Several models may have cached the same template. Taking the first
    acceptable one leaves a 1-of-31 parse in place when a better one exists;
    the richest wins instead."""
    import json

    root = tmp_path / "loop"
    poor = {0: "other", 1: "other", 2: "closing"}
    rich = {0: "title", 1: "bullet_list", 2: "closing"}
    for model, parse in (("model-a", poor), ("model-b", rich)):
        d = root / model / "tpl"
        d.mkdir(parents=True)
        (d / "archetypes.json").write_text(json.dumps(parse), encoding="utf-8")

    monkeypatch.setattr(probe_module, "LOOP_ROOT", str(root))
    archetypes, _ = probe_module._cached_parse("tpl")
    assert archetypes == {int(k): v for k, v in rich.items()}
