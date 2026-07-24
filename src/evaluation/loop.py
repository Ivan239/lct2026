"""One turn of the improvement loop, headless and per-LLM.

    pick a model -> ensure a template parsed BY that model -> generate a deck
    from a canonical brief -> score it against the rubric.

Deliberately self-contained: it drives the same lower-level pipeline blocks the
API does, but does NOT import api.main (whose import runs preset bootstrap and
reaches GigaChat). Each model gets its own workdir under output/loop/<model>/ so
"a template from LLM X" is a real, cached, separately-owned artifact — the
provenance the loop needs to answer "do we already have one from this LLM?".
"""

import json
import os
import re
import time

from pptx import Presentation

from common.pictures import has_oversized_picture
from design_system.extractor import _describe_slide, build_archetype_map
from design_system.fingerprint_cache import FingerprintCache
from design_system.style_card import (
    build_style_card,
    card_path,
    card_prompt_preamble,
    load_card,
    save_card,
)
from design_system.style_profile import build_measured_profile
from content_parser.two_phase import generate_block, generate_outline
from evaluation.evaluate import evaluate_deck
from generator.generator import generate
from matcher.matcher import plan_from_outline
from rendering.render import render_pptx_to_pngs
from template_parser.parser import extract_template
from template_spec.builder import build_spec

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
LOOP_ROOT = os.path.join(BASE, "output", "loop")
# Shared geometry-only cache — same file the API uses; safe to share across every
# template and model (no client text in it), and it saves re-classification.
FINGERPRINT_CACHE_PATH = os.path.join(BASE, "output", "parsed", "fingerprint_cache.json")


def safe_model(model):
    return re.sub(r"[^A-Za-z0-9._-]", "_", model)


def _workdir(model, out_root=LOOP_ROOT):
    d = os.path.join(out_root, safe_model(model))
    os.makedirs(os.path.join(d, "rendered"), exist_ok=True)
    os.makedirs(os.path.join(d, "decks"), exist_ok=True)
    return d


def ensure_template(client, model, source_pptx, source_name, out_root=LOOP_ROOT, force=False):
    """Parse `source_pptx` under `model` if not already cached for it. Returns a
    dict with the archetype map, spec, style preamble, and meta (incl. which LLM
    parsed it and when). Reuses the cached parse on subsequent iterations."""
    workdir = _workdir(model, out_root)
    arch_path = os.path.join(workdir, "archetypes.json")
    meta_path = os.path.join(workdir, "meta.json")

    render_dir = os.path.join(workdir, "rendered")
    png_paths = render_pptx_to_pngs(source_pptx, render_dir)

    if os.path.exists(arch_path) and not force:
        with open(arch_path, encoding="utf-8") as f:
            archetype_map = {int(k): v for k, v in json.load(f).items()}
        meta = json.load(open(meta_path, encoding="utf-8")) if os.path.exists(meta_path) else {}
    else:
        struct = extract_template(source_pptx)
        archetype_map = build_archetype_map(
            client, struct, rendered_png_paths=png_paths, text_models=[model],
            fingerprint_cache=FingerprintCache(FINGERPRINT_CACHE_PATH),
        )
        # Same exclusion the API applies: a slide built around a big source
        # picture can't host our text-only archetypes.
        prs = Presentation(source_pptx)
        for idx in list(archetype_map):
            if has_oversized_picture(prs.slides[idx], prs.slide_width, prs.slide_height):
                archetype_map[idx] = "other"
        with open(arch_path, "w", encoding="utf-8") as f:
            json.dump(archetype_map, f, ensure_ascii=False, indent=2)

        if load_card(workdir, "template") is None:
            try:
                descriptions = [_describe_slide(s) for s in struct["slides"]]
                card = build_style_card(client, descriptions, profile=None,
                                        archetype_map=archetype_map, models=[model])
                save_card(workdir, "template", card)
            except Exception as e:  # noqa: BLE001 — style card is advisory
                print(f"[loop] style_card failed for {model}: {e}", flush=True)

        meta = {
            "model": model,
            "source_pptx": os.path.abspath(source_pptx),
            "source_name": source_name,
            "parsed_at": int(time.time()),
        }
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

    profile = build_measured_profile(png_paths, archetype_map) if png_paths else None
    spec = build_spec(source_pptx, archetype_map, style_profile=profile)
    style_preamble = card_prompt_preamble(load_card(workdir, "template"))
    return {
        "workdir": workdir,
        "archetype_map": archetype_map,
        "spec": spec,
        "style_preamble": style_preamble,
        "png_paths": png_paths,
        "meta": meta,
        "has_card": load_card(workdir, "template") is not None,
    }


def generate_deck(client, model, source_pptx, spec, brief, style_preamble, out_pptx):
    """Two-phase generation under a single model. Returns (plan, skipped) where
    plan is [(block, slide_idx)] — its order is the final slide order."""
    outline = generate_outline(client, brief, spec, models=[model], style_preamble=style_preamble)
    assignments, skipped_items = plan_from_outline(outline, spec)
    plan = []
    for item, slide_idx, final_count in assignments:
        block = generate_block(client, item["role"], item["theme"], brief,
                               count=final_count, models=[model], style_preamble=style_preamble)
        plan.append((block, slide_idx))
    generate(source_pptx, plan, out_pptx)
    skipped = [{"type": it["role"], "title": it.get("theme")} for it in skipped_items]
    return plan, skipped


def run_iteration(client, model, source_pptx, brief, source_name=None,
                  out_root=LOOP_ROOT, judge_client=None):
    """Full turn for one model. `client` parses the template and generates the
    deck (the model under test); `judge_client` scores it — kept separate so the
    judge stays independent of the generator (defaults to `client` for the
    single-backend case). Returns a bundle: template info, deck path, the rubric
    evaluation, and what was skipped. Raises only on a hard failure that leaves
    no deck to score (the caller records that and moves on)."""
    source_name = source_name or os.path.splitext(os.path.basename(source_pptx))[0]
    t = ensure_template(client, model, source_pptx, source_name, out_root=out_root)

    ts = int(time.time())
    out_pptx = os.path.join(t["workdir"], "decks", f"deck_{ts}.pptx")
    plan, skipped = generate_deck(client, model, source_pptx, t["spec"], brief,
                                  t["style_preamble"], out_pptx)
    slide_roles = {pos: block["type"] for pos, (block, _) in enumerate(plan)}

    result = evaluate_deck(out_pptx, brief, client=judge_client or client,
                           slide_roles=slide_roles, label=f"{safe_model(model)}_{ts}")
    return {
        "model": model,
        "source_name": source_name,
        "deck": out_pptx,
        "n_slides": len(plan),
        "skipped": skipped,
        "evaluation": result,
        "template_meta": t["meta"],
    }
