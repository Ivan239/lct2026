"""Apply the reviewer's own visual/content judgment to a deterministic-only evaluation.

No LLM API involved — the reviewer looks at the rendered slide PNGs directly (Read
tool), then calls this with its own scores for the llm-mode rubric criteria.
This just merges the judgment into the eval JSON in place and recomputes the
weighted total.

    .venv/bin/python3 scripts/review_score.py --eval output/evaluations/xyz.json \
        --scores '{"1.3": [3, "верх слайда плотный, низ пустой"], "5.6": [5, "источников нет, но и не выдуманы"]}'

--scores is a JSON object: {criterion_id: [score_1_to_5_or_null, "detail"]}.
Criteria omitted are left as-is (N/A unless already scored).
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from evaluation.manual_review import apply_review_scores
from evaluation.evaluate import format_report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--eval", required=True, help="path to an evaluate_deck() JSON")
    ap.add_argument("--scores", required=True,
                    help='JSON object: {"crit_id": [score_or_null, "detail"], ...}')
    args = ap.parse_args()

    with open(args.eval, encoding="utf-8") as f:
        result = json.load(f)

    raw = json.loads(args.scores)
    scores = {cid: (v[0], v[1]) for cid, v in raw.items()}
    result = apply_review_scores(result, scores)

    result.pop("_json_path", None)
    with open(args.eval, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(format_report(result))


if __name__ == "__main__":
    main()
