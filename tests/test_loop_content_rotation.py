"""The improvement loop generates from content packages, not one number-free
brief.

The canonical brief has no numbers, so every number on a loop deck was
invented and dop_numbers_sourced stood at 1 on every deck (iter108: 20 of 20)
— a criterion that measures nothing. The loop now rotates the sample packages
(samples/content_packages/: feature, project, initiative) and hands each
package's numbers to the audit.

The trap pinned here: the template schedule has six slots and there are three
packages. Counters moving in lockstep would give VK Tech (slots 0 and 3) the
same package forever.
"""

import argparse
import importlib.util
import os

from conftest import ROOT

spec = importlib.util.spec_from_file_location(
    "improve_loop", os.path.join(ROOT, "scripts", "improve_loop.py"))
loop = importlib.util.module_from_spec(spec)
spec.loader.exec_module(loop)


def test_every_template_slot_sees_every_package(tmp_path, monkeypatch):
    monkeypatch.setattr(loop, "STATE_FILE", str(tmp_path / "state.json"))
    packages = loop._packages()
    assert len(packages) >= 3
    seen = {}
    cycles = len(packages)
    for n in range(len(loop.TEMPLATE_SCHEDULE) * cycles):
        _, template, package = loop._next_from_rotation()
        seen.setdefault((n % len(loop.TEMPLATE_SCHEDULE), template), set()).add(package)
    for (slot, template), got in seen.items():
        if template == loop.BLIND_TEMPLATE or template in loop.VK_TEMPLATES:
            assert got == set(packages), f"slot {slot} ({template}) saw only {sorted(got)}"


def test_a_package_brings_its_facts_and_its_numbers(tmp_path):
    args = argparse.Namespace(content="auto", brief_file=None)
    package = loop._packages()[0]
    brief, numbers, label = loop._resolve_content(args, package)
    assert "Факты (используй только их" in brief
    assert numbers and package in label


def test_the_canonical_brief_is_still_available_and_has_no_numbers():
    args = argparse.Namespace(content="canonical", brief_file=None)
    brief, numbers, label = loop._resolve_content(args, "anything")
    assert brief == loop.CANONICAL_BRIEF and numbers == [] and "канонический" in label
