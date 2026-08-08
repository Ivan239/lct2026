"""One improvement-loop iteration, headless.

    .venv/bin/python3 scripts/improve_loop.py [--model GigaChat-2-Max]
                                              [--source output/templates/<file>.pptx]
                                              [--brief-file path.txt]

Picks a model, ensures a template parsed by that model exists, generates a deck
from a canonical brief, scores it against docs/evaluation_rubric.md, and prints
the report + the paths. The detailed improvement PLAN and any code changes are
authored by whoever reads this output (a human, or the agent on a scheduled
fire) — this script's job is to produce the reproducible evaluation to plan from.
"""

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dotenv import load_dotenv

load_dotenv()

from evaluation.evaluate import format_report
from evaluation.loop import run_iteration
from llm_clients import backends

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEMPLATES_DIR = os.path.join(BASE, "output", "templates")

# Rotated across iterations for diversification. Max first — it's the verified
# flagship. No judge model here anymore — Claude scores the LLM-mode criteria
# by looking at the renders (see evaluation.claude_review), not an API call.
ROTATION = ["GigaChat-2-Max", "GigaChat-2-Pro", "GigaChat-2", "GigaChat-3-Ultra"]

# REAL customer templates, not the 4-slide toy presets. The presets were useful
# to get the loop running, but their cases are trivial — one bullet family, one
# stats family, generous boxes — so the deck always "fits" and the hard problems
# (dense designer layouts, icon lists, capacity mismatches, mixed backgrounds)
# never show up. These are the files actually uploaded to the product.
TEMPLATE_ROTATION = [
    "custom_f496182bb15f42bb",   # Т—Ж Учебный шаблон (12 слайдов, лёгкий по весу)
    "custom_838830368dac3116",   # Т—Ж Монохромный — деловые презентации с цифрами (12)
    "custom_47dfd8952eb47583",   # Т—Ж Универсальный — для любых задач (12)
    "custom_30e96c06e2d47ec3",   # 31-слайдовая дека («example»)
    "custom_78dc579e05d11399",   # November survey results 2024 — 69 слайдов, самый тяжёлый
]
STATE_FILE = os.path.join(BASE, ".loop_state.json")


def _load_state():
    if os.path.exists(STATE_FILE):
        try:
            return json.load(open(STATE_FILE))
        except Exception:
            pass
    return {}


def _next_from_rotation():
    """Advance model and template together, on co-prime-ish counters so the loop
    walks through combinations instead of pinning one template to one model.
    State file is gitignored so it never pollutes the safety-net commits."""
    st = _load_state()
    i, j = st.get("i", 0), st.get("t", 0)
    model = ROTATION[i % len(ROTATION)]
    template = TEMPLATE_ROTATION[j % len(TEMPLATE_ROTATION)]
    with open(STATE_FILE, "w") as f:
        json.dump({"i": (i + 1) % len(ROTATION),
                   "t": (j + 1) % len(TEMPLATE_ROTATION)}, f)
    return model, template


def _resolve_source(arg):
    """`--source auto` (default) rotates real templates; a bare id like
    custom_838830368dac3116 is resolved inside output/templates/; anything else
    is treated as a path."""
    if arg and arg != "auto":
        if os.path.exists(arg):
            return arg
        candidate = os.path.join(TEMPLATES_DIR, f"{arg}.pptx")
        if os.path.exists(candidate):
            return candidate
        raise SystemExit(f"шаблон не найден: {arg}")
    return None

# A canonical brief that exercises the whole rubric: a clear task, a named
# audience and tone (for prompt-adherence/adaptation criteria), and a
# problem->solution->metrics->outcome arc (for structure/logic/completeness).
CANONICAL_BRIEF = (
    "Продукт: облачная платформа «Поток» для аналитики продаж среднего бизнеса. "
    "Задача презентации: убедить коммерческого директора внедрить платформу. "
    "Аудитория: коммерческие директора и руководители отделов продаж. "
    "Тон: деловой, без хайпа, с опорой на цифры. "
    "Нужно раскрыть: проблему разрозненных данных о продажах; как «Поток» их "
    "объединяет; ключевые возможности (единый дашборд, прогноз спроса, "
    "автоотчёты); измеримые результаты внедрения (рост конверсии, экономия "
    "времени, точность прогноза); и завершающий слайд с призывом к пилоту."
)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default="GigaChat-2-Max",
                    help="model name, or 'auto' to rotate through ROTATION")
    ap.add_argument("--source", default="auto",
                    help="'auto' rotates real templates, or a template id / path")
    ap.add_argument("--brief-file", default=None)
    ap.add_argument("--out-json", default=None, help="also write the result bundle here")
    args = ap.parse_args()

    brief = CANONICAL_BRIEF
    if args.brief_file:
        with open(args.brief_file, encoding="utf-8") as f:
            brief = f.read().strip()

    rotated_model, rotated_template = _next_from_rotation()
    requested = rotated_model if args.model == "auto" else args.model
    source = _resolve_source(args.source) or os.path.join(TEMPLATES_DIR, f"{rotated_template}.pptx")
    try:
        gen_client, model_name = backends.resolve(requested)
    except backends.BackendUnavailable as e:
        print(f"[loop] модель {requested} в очереди: {e}")
        return
    source_name = os.path.splitext(os.path.basename(source))[0]
    print(f"[loop] model={model_name} source={source_name}", flush=True)
    result = run_iteration(gen_client, model_name, source, brief, source_name=source_name)

    ev = result["evaluation"]
    print("\n" + format_report(ev))
    print(f"\nSlides: {result['n_slides']}  Skipped: {len(result['skipped'])}")
    print(f"Из шаблона: {result['native_slides']} нативных / "
          f"{result['synth_slides']} синтезированных;  "
          f"шаблон предлагает {result['template_offered']} слайдов из {result['template_total']}")
    print(f"Deck:   {result['deck']}")
    print(f"Eval:   {ev['_json_path']}")

    # No LLM judge is called, ever — GigaChat vision was measured to be a bad
    # one (scored a duplicate-slides, half-empty deck 90+/100; the user looked
    # at the same renders and called it 3-4/10). Claude IS the judge: look at
    # every PNG below, then finalize with scripts/claude_score.py.
    print("\n>>> ПОСМОТРИ ГЛАЗАМИ на каждый слайд (Read) — ты судья, не число выше:")
    for p in ev["png_paths"]:
        print("    " + p)
    print(f"\nПосле просмотра примени свою оценку:")
    print(f'    .venv/bin/python3 scripts/claude_score.py --eval {ev["_json_path"]} '
          f'--scores \'{{"1.3": [4, "..."], ...}}\'')

    if args.out_json:
        slim = {k: v for k, v in result.items() if k != "evaluation"}
        slim["total_100"] = ev["total_100"]
        slim["eval_json"] = ev["_json_path"]
        with open(args.out_json, "w", encoding="utf-8") as f:
            json.dump(slim, f, ensure_ascii=False, indent=2)
        print(f"Bundle: {args.out_json}")


if __name__ == "__main__":
    main()
