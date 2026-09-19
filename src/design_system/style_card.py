"""Template style card — the one-LLM-call half of the Template Style Profile
(plan 9.4). Measured signals (style_profile.py) capture the color/photo
rhythm; what they cannot see is the SEMANTIC design language: tone of voice,
how long texts are meant to be per role, the visual metaphor, what to avoid.
One call per template at upload time, cached forever with a version stamp —
the cost amortizes over every future generation against that template.

The card is advice for the generation prompts, never a source of hard
constraints: numeric limits (exact item counts) stay code-enforced, because
the model does not reliably follow worded rules (project-wide lesson)."""

import json
import os

from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, call_with_model_fallback
from common.prompt_files import load_prompt

# Bump on ANY change to the prompt or card schema — same poisoning story as
# fingerprint_cache.CLASSIFIER_VERSION: stale cached cards would otherwise
# replay an outdated design brief forever.
STYLE_CARD_VERSION = 1

# The move to a file kept the text byte-for-byte, so the version above
# stays; editing the file is a prompt change and must bump it.
CARD_PROMPT = load_prompt("style_card.v1.txt")


def _measured_summary(profile, archetype_map):
    if not profile:
        return "нет данных"
    parts = [f"палитра фонов: {len(profile.get('palette', []))} цвет(ов)"]
    if profile.get("rotation", {}).get("order"):
        parts.append("фоны чередуются по циклу")
    roles = sorted(set((archetype_map or {}).values()))
    if roles:
        parts.append(f"роли слайдов: {', '.join(roles)}")
    return "; ".join(parts)


def _validate_card(card):
    if not isinstance(card, dict) or not card.get("tone"):
        raise ValueError("style card missing tone")
    norms = card.get("length_norms")
    if norms is not None and not isinstance(norms, dict):
        raise ValueError("length_norms must be a dict")
    avoid = card.get("avoid")
    if avoid is not None and not isinstance(avoid, list):
        raise ValueError("avoid must be a list")
    return card


def build_style_card(client, slide_descriptions, profile=None, archetype_map=None,
                     models=TEXT_MODELS):
    """slide_descriptions: list of extractor._describe_slide strings for the
    cluster medoids (compact, layout-noise already stripped). One LLM call
    with validation+retries; raises if every attempt fails."""
    prompt = (
        CARD_PROMPT
        .replace("__DESCRIPTIONS__", "\n\n".join(slide_descriptions[:12]))
        .replace("__MEASURED__", _measured_summary(profile, archetype_map))
    )

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=500)
        return _validate_card(extract_json(result["choices"][0]["message"]["content"]))

    return call_with_model_fallback(call, models)


def card_path(parsed_dir, template_id):
    return os.path.join(parsed_dir, f"{template_id}_style.json")


def load_card(parsed_dir, template_id):
    """Version-checked read; a stale or unreadable card counts as absent."""
    try:
        with open(card_path(parsed_dir, template_id), encoding="utf-8") as f:
            stored = json.load(f)
        if stored.get("v") != STYLE_CARD_VERSION:
            return None
        return stored["card"]
    except Exception:
        return None


def save_card(parsed_dir, template_id, card):
    with open(card_path(parsed_dir, template_id), "w", encoding="utf-8") as f:
        json.dump({"v": STYLE_CARD_VERSION, "card": card}, f, ensure_ascii=False, indent=2)


def card_prompt_preamble(card):
    """The injectable design-brief paragraph for outline/block prompts.
    Advice only — hard limits stay in code."""
    if not card:
        return ""
    lines = [f"Дизайн-бриф шаблона: тон — {card['tone']}."]
    if card.get("metaphor"):
        lines.append(f"Характер шаблона: {card['metaphor']}.")
    norms = card.get("length_norms") or {}
    if norms:
        norm_str = ", ".join(f"{k}: {v}" for k, v in norms.items())
        lines.append(f"Ориентиры длины текстов: {norm_str}.")
    if card.get("avoid"):
        lines.append("Избегай: " + "; ".join(str(a) for a in card["avoid"]) + ".")
    return "\n".join(lines)
