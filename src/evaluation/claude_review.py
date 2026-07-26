"""Claude's own visual/content judgment — the LLM-mode half of the rubric.

Per explicit decision: no LLM API is ever called to judge a deck. GigaChat vision
was the original judge and was measured to be a bad one — it scored a deck with
two duplicate stat slides and half-empty layouts on nearly every slide as
90+/100, and stayed wrong across several rounds of prompt tuning. The user
looked at the same renders and called it 3-4/10 immediately. Claude reviewing
the rendered PNGs directly (Read tool) IS the judge now; this module only merges
that judgment into evaluate_deck()'s deterministic-only result and recomputes
the weighted total — pure arithmetic, no network call.
"""

from evaluation import rubric

LLM_CRITERIA = {cid for cid, (_, mode) in rubric.CRITERIA.items() if mode == "llm"}


def apply_claude_scores(result, scores):
    """result: an evaluate_deck() dict. scores: {criterion_id: (score_1_to_5_or_None,
    detail_str)} — Claude's own per-criterion judgment after looking at the
    renders. Only touches llm-mode criteria (deterministic ones are the harness's
    own job, not Claude's to overwrite); raises on anything else so a typo'd id
    fails loudly instead of silently no-op'ing.

    Missing llm-mode criteria are left as whatever evaluate_deck put there
    ("ожидает оценки Клода" -> N/A) rather than forced — Claude can score a
    subset (e.g. only the ones relevant to the deck's actual media) and leave
    the rest N/A same as the old skip_ids gating did."""
    unknown = set(scores) - LLM_CRITERIA
    if unknown:
        raise ValueError(f"не LLM-критерии (это забота детерминированных проверок): {sorted(unknown)}")

    for cid, (score, detail) in scores.items():
        title, mode = rubric.CRITERIA[cid]
        result["scores"][cid] = {"title": title, "mode": mode, "score": score, "detail": detail}

    result["total_100"] = rubric.weighted_total(result["scores"])
    result["buckets"] = rubric.bucket_breakdown(result["scores"])
    result["llm_evaluated"] = True
    result["judge"] = "claude"
    return result
