"""Orchestrator: render a deck, score every rubric criterion (deterministic +
LLM), compute the weighted 100-point total, and persist the result.

This is the "оцениваешь по нашему шаблону" step of the improvement loop. It's the
one place that knows the deck as a whole — text for the content judge, images for
the visual judge, geometry for the deterministic checks — so it owns rendering
and text extraction and hands the pieces to each scorer.
"""

import json
import os
import time

from pptx import Presentation

from evaluation import deterministic, judge, rubric
from rendering.render import render_pptx_to_pngs

EVAL_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "output", "evaluations")


def _slide_texts(pptx_path):
    prs = Presentation(pptx_path)
    texts = []
    for slide in prs.slides:
        parts = [s.text_frame.text.strip() for s in slide.shapes
                 if s.has_text_frame and s.text_frame.text.strip()]
        texts.append("\n".join(parts))
    return texts


def evaluate_deck(pptx_path, brief, client=None, slide_roles=None,
                  render_dir=None, out_json=None, label=None,
                  vision_models=None, text_models=None):
    """Score one deck. Without a `client` only the deterministic half runs (still
    a real partial score, zero tokens) — useful for CI and for validating the
    harness. Returns the full result dict and writes it to output/evaluations/."""
    render_dir = render_dir or os.path.join(EVAL_DIR, "_render")
    os.makedirs(render_dir, exist_ok=True)
    os.makedirs(EVAL_DIR, exist_ok=True)

    scores = {}
    scores.update(deterministic.evaluate(pptx_path, slide_roles=slide_roles))

    judged = {}
    if client is not None:
        png_paths = render_pptx_to_pngs(pptx_path, render_dir)
        slide_texts = _slide_texts(pptx_path)
        # Force image/infographics criteria to N/A when the deck has no such
        # media — deterministic, so it can't drift with the judge's mood.
        media = deterministic.deck_media(pptx_path)
        skip = set()
        if media["substantive_pictures"] == 0:
            skip |= {"4.1", "4.2", "4.3", "4.4", "4.5", "dop_image_style"}
        if media["charts"] == 0:
            skip |= {"7.1", "7.2", "7.3"}
        kw = {"skip_ids": skip}
        if vision_models:
            kw["vision_models"] = vision_models
        if text_models:
            kw["text_models"] = text_models
        judged = judge.judge(client, png_paths, brief, slide_texts, **kw)
        scores.update(judged)

    # criteria never scored (no client, or judge fully failed) -> N/A, excluded
    # from the weighted total rather than dragging it to zero.
    for cid in rubric.CRITERIA:
        scores.setdefault(cid, {"score": None, "detail": "не оценивалось"})

    total = rubric.weighted_total(scores)
    result = {
        "pptx": os.path.abspath(pptx_path),
        "label": label or os.path.basename(pptx_path),
        "brief": brief,
        "timestamp": int(time.time()),
        "total_100": total,
        "buckets": rubric.bucket_breakdown(scores),
        "scores": {
            cid: {"title": rubric.CRITERIA[cid][0], "mode": rubric.CRITERIA[cid][1], **scores[cid]}
            for cid in rubric.CRITERIA
        },
        "llm_evaluated": bool(judged and any(v["score"] is not None for v in judged.values())),
    }

    out_json = out_json or os.path.join(EVAL_DIR, f"{result['label']}_{result['timestamp']}.json")
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    result["_json_path"] = os.path.abspath(out_json)
    return result


def format_report(result):
    """Human-readable one-screen summary of an evaluate_deck result."""
    lines = [
        f"# Оценка: {result['label']}",
        f"Итог: {result['total_100']}/100" + ("" if result["llm_evaluated"] else "  (только детерминированные критерии)"),
        "",
        "## По категориям",
    ]
    for b in result["buckets"].values():
        s = b["score_1_5"]
        lines.append(f"- {b['title']} (вес {b['weight']:.0%}): "
                     + (f"{s}/5" if s is not None else "N/A"))
    lines.append("")
    lines.append("## Слабые места (score ≤ 3)")
    weak = sorted(
        ((cid, v) for cid, v in result["scores"].items()
         if v["score"] is not None and v["score"] <= 3),
        key=lambda kv: kv[1]["score"],
    )
    if weak:
        for cid, v in weak:
            lines.append(f"- [{v['score']}] {cid} {v['title']} — {v['detail']}")
    else:
        lines.append("- нет (все оценённые критерии ≥ 4)")
    return "\n".join(lines)
