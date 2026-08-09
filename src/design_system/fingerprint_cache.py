"""Cross-template memory of classified slide geometries.

A cluster's fingerprint (see clustering.slide_fingerprint) is pure geometry —
shape roles, size buckets, grid cells — with no client text, so classifications
learned on one uploaded template can safely be reused on every later one.
Corporate decks overwhelmingly reuse the same handful of recipes (title +
bullets, KPI grid, two-column comparison...), which means the marginal LLM cost
of processing a template falls as the library grows.

Entries vote: a repeated confirmation strengthens an entry, a conflicting
classification weakens it, and an entry voted down to zero gets replaced. Only
confident, non-"other" classifications are ever written (the caller enforces
that) — a hesitant guess must not poison every future template.
"""

import json
import os

# Bump on ANY change to classification behavior: the prompt wording, the slide
# description format (design_system/extractor.py), or the model line. Entries
# written by an older classifier are silently ignored (and overwritten on the
# next put) — without this, a poisoned/outdated label replays forever: a T-Ж
# template was re-classified twice with zero LLM calls, both times faithfully
# reproducing the WRONG labels cached by the very first buggy run.
CLASSIFIER_VERSION = 4


class FingerprintCache:
    def __init__(self, path):
        self.path = path
        try:
            with open(path, encoding="utf-8") as f:
                self._data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self._data = {}

    def get(self, fingerprint):
        entry = self._data.get(fingerprint)
        if not entry or entry.get("v") != CLASSIFIER_VERSION:
            return None
        return entry["archetype"]

    def put(self, fingerprint, archetype):
        entry = self._data.get(fingerprint)
        if entry is None or entry.get("v") != CLASSIFIER_VERSION:
            self._data[fingerprint] = {"archetype": archetype, "votes": 1, "v": CLASSIFIER_VERSION}
        elif entry["archetype"] == archetype:
            entry["votes"] += 1
        else:
            entry["votes"] -= 1
            if entry["votes"] <= 0:
                self._data[fingerprint] = {"archetype": archetype, "votes": 1, "v": CLASSIFIER_VERSION}
        self._save()

    def _save(self):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)
