"""Prompts are versioned files under prompts/ (backlog item 9).

The outline and block prompts moved out of content_parser/two_phase.py
unchanged (iter138: loaded text compared byte-for-byte with the inline text of
HEAD). What must hold from now on: every prompt the code loads exists, and
every placeholder the code substitutes is in the file — a renamed placeholder
would silently ship «__BRIEF__» to the model instead of the brief.
"""

import os

from common.prompt_files import PROMPTS_DIR, load_prompt
from content_parser import parser as legacy
from content_parser import two_phase as tp
from design_system import extractor, style_card

SUBSTITUTED = {
    "outline": ("__MENU__", "__ROLES__", "__RANGE__", "__BRIEF__"),
    "title": ("__THEME__", "__BRIEF__"),
    "section_divider": ("__THEME__",),
    "image_caption": ("__THEME__", "__BRIEF__"),
    "closing": ("__THEME__", "__BRIEF__"),
    "bullet_list": ("__THEME__", "__BRIEF__", "__COUNT__", "__MAXCHARS__"),
    "stats_kpi": ("__THEME__", "__BRIEF__", "__COUNT__"),
    "two_column_comparison": ("__THEME__", "__BRIEF__", "__COUNT__"),
}


def test_outline_and_block_prompts_come_from_files():
    assert tp.OUTLINE_PROMPT == load_prompt("outline.v1.txt")
    assert set(tp.BLOCK_PROMPTS) == set(tp.OUTLINE_ROLES)
    for role, text in tp.BLOCK_PROMPTS.items():
        assert text == load_prompt(f"block_{role}.v1.txt"), role


def test_every_substituted_placeholder_is_in_its_file():
    for name, placeholders in SUBSTITUTED.items():
        text = tp.OUTLINE_PROMPT if name == "outline" else tp.BLOCK_PROMPTS[name]
        missing = [p for p in placeholders if p not in text]
        assert not missing, (name, missing)
    shorten = load_prompt(tp.SHORTEN_ITEMS_PROMPT)
    assert all(p in shorten for p in ("__MAXCHARS__", "__ITEMS__", "__COUNT__"))


def test_no_placeholder_survives_a_real_block_prompt():
    seen = []

    class Client:
        def chat(self, messages, model=None, **kwargs):
            seen.append(messages[-1]["content"])
            return {"choices": [{"message": {"content":
                '{"title": "Заголовок", "bullets": ["один", "два", "три"]}'}}]}

    tp.generate_block(Client(), "bullet_list", "тема", "бриф", count=3, models=["m"], item_chars=28)
    assert seen and "__" not in seen[0], seen[0]


def test_the_remaining_prompts_come_from_files_too():
    """iter140: the legacy parser, the style card and the classifier. Their
    text moved byte-for-byte, which is what lets CLASSIFIER_VERSION and
    STYLE_CARD_VERSION stay put — a changed prompt would replay stale cached
    labels and cards forever."""
    assert legacy.SCHEMA_PROMPT == load_prompt("schema_blocks.v1.txt")
    assert legacy.TITLE_PROMPT_TEMPLATE == load_prompt("title_for_block.v1.txt")
    assert legacy.RESIZE_PROMPT_TEMPLATE == load_prompt("resize_block.v1.txt")
    assert style_card.CARD_PROMPT == load_prompt("style_card.v1.txt")
    assert extractor.TEXT_PROMPT_TEMPLATE == load_prompt("classify_slide_text.v1.txt")


def test_the_vision_prompt_lists_the_archetypes():
    """It is the one prompt built by substitution at import: the file carries
    __ARCHETYPES__, and a renamed placeholder would ship the literal to the
    model."""
    assert "__" not in extractor.VISION_PROMPT, extractor.VISION_PROMPT
    for archetype in extractor.ARCHETYPES:
        assert archetype in extractor.VISION_PROMPT, archetype


def test_the_legacy_prompts_keep_their_placeholders():
    assert "__BLOCK__" in legacy.TITLE_PROMPT_TEMPLATE
    for placeholder in ("__BLOCK__", "__COUNT__", "__FIELD__"):
        assert placeholder in legacy.RESIZE_PROMPT_TEMPLATE, placeholder
    for placeholder in ("__DESCRIPTIONS__", "__MEASURED__"):
        assert placeholder in style_card.CARD_PROMPT, placeholder
    for placeholder in ("__ARCHETYPES__", "__DESCRIPTION__"):
        assert placeholder in extractor.TEXT_PROMPT_TEMPLATE, placeholder


def test_no_prompt_is_left_inline():
    """The point of item 9: swapping a prompt for another model must not need a
    code edit. A new triple-quoted prompt in these modules would undo it."""
    import inspect

    for module in (tp, legacy, extractor, style_card):
        source = inspect.getsource(module)
        for line in source.splitlines():
            stripped = line.strip()
            if stripped.startswith("#"):
                continue
            assert not (("PROMPT" in stripped or "PROMPTS" in stripped)
                        and '"""' in stripped and "=" in stripped), (module.__name__, line)


def test_prompt_files_are_versioned():
    names = [n for n in os.listdir(PROMPTS_DIR) if n.endswith(".txt")]
    assert names and all(".v" in n for n in names), names
