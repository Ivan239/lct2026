import os
import sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

TEMPLATES_DIR = os.path.join(ROOT, "output", "templates")

# Real uploaded decks — the regression corpus every past bug was found on.
# Skipped (not failed) when absent so the suite still runs on a fresh clone.
SURVEY_31 = os.path.join(TEMPLATES_DIR, "custom_30e96c06e2d47ec3.pptx")
SURVEY_69 = os.path.join(TEMPLATES_DIR, "custom_78dc579e05d11399.pptx")
TJ_TEMPLATE = os.path.join(TEMPLATES_DIR, "custom_f496182bb15f42bb.pptx")
TJ_MONO = os.path.join(TEMPLATES_DIR, "custom_838830368dac3116.pptx")
TJ_UNIVERSAL = os.path.join(TEMPLATES_DIR, "custom_47dfd8952eb47583.pptx")
PRESET_A = os.path.join(TEMPLATES_DIR, "template_a_corporate.pptx")
PRESET_B = os.path.join(TEMPLATES_DIR, "template_b_startup.pptx")


def requires(path):
    return pytest.mark.skipif(not os.path.exists(path), reason=f"fixture deck missing: {path}")
