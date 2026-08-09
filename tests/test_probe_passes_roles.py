"""The probe evaluates a deck the way the production loop does.

evaluation/loop.py hands slide_roles to evaluate_deck; the offline probe did
not, for thirty-five iterations. Without roles the evenness criteria treat the
cover and the closing as ordinary content slides and mark the deck uneven for
having a cover — the exact false positive SPARSE_EXEMPT_ROLES exists to
prevent. Every dop_distribution and dop_pacing number in those reports was
pessimistic, and the probe and the loop could not be compared.
"""

import importlib.util
import os

import pytest

from conftest import ROOT, TJ_TEMPLATE, requires

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe)


@requires(TJ_TEMPLATE)
def test_the_probe_hands_over_the_roles(tmp_path, monkeypatch):
    if probe._cached_parse("custom_f496182bb15f42bb")[0] is None:
        pytest.skip("no cached parse for the study template")

    seen = {}
    real = probe.evaluate_deck

    def spy(deck, **kwargs):
        seen.update(kwargs)
        return real(deck, **kwargs)

    monkeypatch.setattr(probe, "evaluate_deck", spy)
    monkeypatch.setattr(probe, "render_pptx_to_pngs", lambda *a, **k: [])
    probe.probe("study", "custom_f496182bb15f42bb", probe.LONG, str(tmp_path))

    roles = seen.get("slide_roles")
    assert roles, f"slide_roles was not passed: {sorted(seen)}"
    assert roles[0] == "title", roles
    assert roles[max(roles)] == "closing", roles


def test_roles_change_what_evenness_measures():
    """Why it matters, on the harness side: the same deck is measured over
    different slides once the cover and closing are exempt."""
    from evaluation.deterministic import SPARSE_EXEMPT_ROLES

    assert "title" in SPARSE_EXEMPT_ROLES and "closing" in SPARSE_EXEMPT_ROLES
    assert "stats_kpi" not in SPARSE_EXEMPT_ROLES, (
        "a KPI board is dense by design — exempting the whole role would hide "
        "a genuinely uneven deck")
