"""Every deck the probe writes says what produced it.

Twice a measurement in this project read a deck from an older run and drew a
conclusion from it. iter72 fixed one half — a template's output directory is
cleared when it is skipped and before it is regenerated — but clearing only
helps the directories a run actually touches, and an --out directory from an
earlier session sits on disk with a preset that no longer matches. The stamp
travels inside the file, so a reader can always ask what it is looking at:
a deck from before this change answers «(no stamp)».
"""

import importlib.util
import os

import pytest

from conftest import ROOT

from pptx import Presentation

spec = importlib.util.spec_from_file_location(
    "offline_probe", os.path.join(ROOT, "scripts", "offline_probe.py"))
probe_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(probe_module)


def _deck(tmp_path):
    path = str(tmp_path / "deck.pptx")
    Presentation().save(path)
    return path


def test_the_stamp_names_the_preset_and_the_commit(tmp_path):
    path = _deck(tmp_path)
    probe_module._stamp(path, "wordy", 5)

    stamp = probe_module.describe(path)
    assert "content=wordy" in stamp, stamp
    assert "items=5" in stamp, stamp
    assert "commit=" in stamp, stamp


def test_a_plan_driven_run_says_so(tmp_path):
    path = _deck(tmp_path)
    probe_module._stamp(path, "long", None)
    assert "items=plan" in probe_module.describe(path)


def test_an_unstamped_deck_is_recognisable(tmp_path):
    """The whole point: a deck from an older run must not look current."""
    assert probe_module.describe(_deck(tmp_path)) == "(no stamp)"


def test_the_stamp_survives_a_reopen(tmp_path):
    path = _deck(tmp_path)
    probe_module._stamp(path, "short", None)
    reopened = Presentation(path)
    assert "content=short" in (reopened.core_properties.comments or "")


@pytest.mark.skipif(not os.path.isdir(os.path.join(ROOT, ".git")), reason="not a git checkout")
def test_the_commit_is_real(tmp_path):
    path = _deck(tmp_path)
    probe_module._stamp(path, "long", None)
    sha = probe_module.describe(path).split("commit=")[1].strip()
    assert sha and sha != "?" and len(sha) >= 7, sha
