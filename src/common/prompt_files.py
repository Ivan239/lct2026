"""Prompts kept as versioned files under prompts/ (backlog item 9).

A prompt in a file can be read, diffed and swapped for another model without
touching code — the switch to an open-weights generator (backlog item 1) needs
exactly that. The version is part of the file name (`shorten_items.v1.txt`), so
a changed prompt is a new file and old runs stay reproducible.
"""

import os

PROMPTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), "prompts")


def load_prompt(name):
    with open(os.path.join(PROMPTS_DIR, name), encoding="utf-8") as f:
        return f.read()
