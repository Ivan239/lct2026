"""One improvement-loop iteration, headless.

    .venv/bin/python3 scripts/improve_loop.py [--model GigaChat-2-Max]
                                              [--source output/templates/<file>.pptx]
                                              [--content auto|canonical|<package>]
                                              [--brief-file path.txt]

Picks a model, ensures a template parsed by that model exists, generates a deck
from a content package (samples/content_packages/, rotated) or the canonical
brief, scores it against docs/evaluation_rubric.md, and prints
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

from content_package import extract_numbers, load_package, to_brief_text
from evaluation.evaluate import format_report
from evaluation.loop import run_iteration
from llm_clients import backends

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEMPLATES_DIR = os.path.join(BASE, "output", "templates")

# Rotated across iterations for diversification. Max first — it's the verified
# flagship. No judge model here anymore — the reviewer scores the LLM-mode criteria
# by looking at the renders (see evaluation.manual_review), not an API call.
ROTATION = ["GigaChat-2-Max", "GigaChat-2-Pro", "GigaChat-2", "GigaChat-3-Ultra"]

# REAL customer templates, not the 4-slide toy presets. The presets were useful
# to get the loop running, but their cases are trivial — one bullet family, one
# stats family, generous boxes — so the deck always "fits" and the hard problems
# (dense designer layouts, icon lists, capacity mismatches, mixed backgrounds)
# never show up. These are the files actually uploaded to the product.
# The three templates of the VK Tech case (docs/TZ_VK.md). Two are worked on;
# the third is BLIND: the final defence runs on a template no team has seen, so
# one of ours stands in for it — generated and reviewed, never tuned against.
# If a change lifts the two and drops the blind one, it was overfitting.
VK_TEMPLATES = ["vk_tech", "vk_education"]
BLIND_TEMPLATE = "vk_workspace"

# The earlier corpus stays as a regression check, one turn in six.
REGRESSION_TEMPLATES = [
    "custom_f496182bb15f42bb",   # Т—Ж Учебный шаблон (12 слайдов, лёгкий по весу)
    "custom_838830368dac3116",   # Т—Ж Монохромный — деловые презентации с цифрами (12)
    "custom_47dfd8952eb47583",   # Т—Ж Универсальный — для любых задач (12)
    "custom_30e96c06e2d47ec3",   # 31-слайдовая дека («example»)
    "custom_78dc579e05d11399",   # November survey results 2024 — 69 слайдов, самый тяжёлый
]

# One cycle of six turns: VK×4, regression×1, blind×1.
TEMPLATE_SCHEDULE = ["vk:0", "vk:1", "regression", "vk:0", "vk:1", "blind"]
STATE_FILE = os.path.join(BASE, ".loop_state.json")

# The loop's content: the sample content packages, one per purpose from the VK
# Tech brief (feature, project, initiative). The canonical brief below has no
# numbers at all, so every number on a deck made from it is invented — the
# check «все цифры есть в исходных материалах» stood at 1 on every such deck
# (iter108: 20 of 20) and measured nothing. A package carries facts, a table and
# the reference numbers.
PACKAGES_DIR = os.path.join(BASE, "samples", "content_packages")


def _packages():
    return sorted(d for d in os.listdir(PACKAGES_DIR)
                  if os.path.isdir(os.path.join(PACKAGES_DIR, d)))


def _package_step(p, next_slot):
    """The package counter moves one per iteration and one more whenever the
    template schedule wraps. Moving in lockstep would pin content to template:
    six schedule slots against three packages means VK Tech (slots 0 and 3)
    would only ever see the same package."""
    return p + 1 + (1 if next_slot == 0 else 0)


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
    i, j, r, p = st.get("i", 0), st.get("t", 0), st.get("r", 0), st.get("p", 0)
    model = ROTATION[i % len(ROTATION)]
    slot = TEMPLATE_SCHEDULE[j % len(TEMPLATE_SCHEDULE)]
    if slot == "blind":
        template = BLIND_TEMPLATE
    elif slot == "regression":
        template = REGRESSION_TEMPLATES[r % len(REGRESSION_TEMPLATES)]
        r += 1
    else:
        template = VK_TEMPLATES[int(slot.split(":")[1])]
    packages = _packages()
    package = packages[p % len(packages)] if packages else None
    next_slot = (j + 1) % len(TEMPLATE_SCHEDULE)
    with open(STATE_FILE, "w") as f:
        json.dump({"i": (i + 1) % len(ROTATION),
                   "t": next_slot,
                   "r": r % len(REGRESSION_TEMPLATES),
                   "p": _package_step(p, next_slot) % max(1, len(packages))}, f)
    return model, template, package


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


def _resolve_content(args, rotated_package):
    """(brief, source_numbers, label, slides). A package gives the labelled
    brief, its own numbers as the reference and the deck size it asks for; a
    plain brief is its own reference and gets the brief's 10-15 slides."""
    if args.brief_file:
        with open(args.brief_file, encoding="utf-8") as f:
            brief = f.read().strip()
        return brief, extract_numbers(brief), f"бриф {os.path.basename(args.brief_file)}", None
    if args.content == "canonical" or (args.content == "auto" and not rotated_package):
        return CANONICAL_BRIEF, extract_numbers(CANONICAL_BRIEF), "канонический бриф", None
    path = (os.path.join(PACKAGES_DIR, rotated_package) if args.content == "auto"
            else args.content)
    package = load_package(path)
    label = f"пакет {os.path.basename(os.path.normpath(path))} ({package['purpose_label'] or '—'})"
    return to_brief_text(package), package["numbers"], label, package["slides"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default="GigaChat-2-Max",
                    help="model name, or 'auto' to rotate through ROTATION")
    ap.add_argument("--source", default="auto",
                    help="'auto' rotates real templates, or a template id / path")
    ap.add_argument("--content", default="auto",
                    help="'auto' rotates samples/content_packages, 'canonical' is the "
                         "old brief without numbers, or a package folder / .zip")
    ap.add_argument("--brief-file", default=None,
                    help="a plain brief instead of a package (numbers checked against it)")
    ap.add_argument("--out-json", default=None, help="also write the result bundle here")
    args = ap.parse_args()

    rotated_model, rotated_template, rotated_package = _next_from_rotation()
    brief, source_numbers, content_label, slides = _resolve_content(args, rotated_package)
    requested = rotated_model if args.model == "auto" else args.model
    source = _resolve_source(args.source) or os.path.join(TEMPLATES_DIR, f"{rotated_template}.pptx")
    try:
        gen_client, model_name = backends.resolve(requested)
    except backends.BackendUnavailable as e:
        # "пропущено по сети" is the phrase the loop's own contract uses for
        # this case, so it is greppable in the logs.
        note = "пропущено по сети" if "недоступен по сети" in str(e) else "в очереди"
        print(f"[loop] {note}: модель {requested} — {e}")
        return
    source_name = os.path.splitext(os.path.basename(source))[0]
    print(f"[loop] model={model_name} source={source_name} content={content_label}", flush=True)
    result = run_iteration(gen_client, model_name, source, brief, source_name=source_name,
                           source_numbers=source_numbers, slides=slides)
    result["content"] = content_label

    ev = result["evaluation"]
    print("\n" + format_report(ev))
    asked = f"заказано {slides}" if slides else "по ТЗ 10–15"
    print(f"\nSlides: {result['n_slides']} ({asked})  Skipped: {len(result['skipped'])}")
    secs = result.get("deck_seconds")
    if secs is not None:
        # The VK Tech brief: one deck in at most 5 minutes.
        verdict = "в норме ТЗ" if secs <= 300 else "ДОЛЬШЕ 5 МИНУТ — дефект по ТЗ"
        print(f"Время сборки деки: {secs:.0f} с ({verdict})")
    if source_name == BLIND_TEMPLATE:
        print("[loop] СЛЕПОЙ шаблон: только контроль, правки по нему не делать")
    print(f"Из шаблона: {result['native_slides']} нативных / "
          f"{result['synth_slides']} синтезированных;  "
          f"шаблон предлагает {result['template_offered']} слайдов из {result['template_total']}")
    print(f"Deck:   {result['deck']}")
    print(f"Eval:   {ev['_json_path']}")

    # No LLM judge is called, ever — GigaChat vision was measured to be a bad
    # one (scored a duplicate-slides, half-empty deck 90+/100; the user looked
    # at the same renders and called it 3-4/10). the reviewer IS the judge: look at
    # every PNG below, then finalize with scripts/review_score.py.
    print("\n>>> ПОСМОТРИ ГЛАЗАМИ на каждый слайд (Read) — ты судья, не число выше:")
    for p in ev["png_paths"]:
        print("    " + p)
    print(f"\nПосле просмотра примени свою оценку:")
    print(f'    .venv/bin/python3 scripts/review_score.py --eval {ev["_json_path"]} '
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
