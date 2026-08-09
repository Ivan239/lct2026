"""Generate + render + score a deck on EVERY real template, without the network.

Why this exists: the loop's own runner needs GigaChat, and the improvement work
does not stop when the network does. This drives the same pipeline blocks —
build_spec -> plan_from_outline -> _synth_canvas_hints -> generate — with content
written here instead of by a model, so every layout decision is exercised and the
renders can be looked at.

Two things it deliberately does that the earlier throwaway probes got wrong:

1. It honours the plan's `count`. A probe that passes a fixed three-item block
   ignores the capacity the matcher just computed, so a capacity fix looks like a
   no-op — that happened for real (iter42: a one-big-number stats slide kept
   showing three lines after the fix, because the probe never asked for one).

2. It defaults to LONG content. Short, convenient bullets stop finding defects:
   the same five decks scored 99.4-100 and looked clean on short content, and the
   moment the text became realistic the checks produced three flags and the
   renders showed 8pt microtext (iter51) and decor sitting on a KPI figure
   (iter52). `--content short` is kept for comparing the two.

3. It can vary the ITEM COUNT. The plan's capacity match decides which template
   slide a role lands on, so a fixed brief only ever exercises one branch: with
   three bullets the universal template hands the list to its prose slide
   (capacity 3) and its designed 6-cell numbered grid — a whole layout — had
   never been rendered by this probe. `--items 5` reaches it.

Usage:
    .venv/bin/python3 scripts/offline_probe.py [--content long|short] [--only NAME]
                                               [--items N]
"""

import argparse
import datetime
import glob
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from pptx import Presentation  # noqa: E402

from common.synthesis import SYNTHESIZE  # noqa: E402
from content_parser.two_phase import _enforce_text_budgets, bullet_char_budget  # noqa: E402
from design_system.style_profile import build_measured_profile  # noqa: E402
from evaluation.evaluate import evaluate_deck  # noqa: E402
from evaluation.loop import _synth_canvas_hints  # noqa: E402
from generator.generator import generate  # noqa: E402
from matcher.matcher import plan_from_outline  # noqa: E402
from rendering.render import render_pptx_to_pngs  # noqa: E402
from template_spec.builder import build_spec  # noqa: E402

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
TEMPLATES = os.path.join(BASE, "output", "templates")
LOOP_ROOT = os.path.join(BASE, "output", "loop")
OUT_ROOT = os.path.join(BASE, "output", "offline_probe")

REAL_TEMPLATES = [
    ("universal", "custom_47dfd8952eb47583"),
    ("mono", "custom_838830368dac3116"),
    ("study", "custom_f496182bb15f42bb"),
    ("survey-31", "custom_30e96c06e2d47ec3"),
    ("survey-69", "custom_78dc579e05d11399"),
]

OUTLINE = [
    {"role": "title", "theme": "платформа", "count": None},
    {"role": "bullet_list", "theme": "что мешало", "count": 3},
    {"role": "stats_kpi", "theme": "результаты", "count": 3},
    {"role": "two_column_comparison", "theme": "было и стало", "count": 2},
    {"role": "image_caption", "theme": "платформа в работе", "count": None},
    {"role": "closing", "theme": "пилот", "count": None},
]

LONG = {
    "title": {"type": "title", "title": "Платформа управленческой отчётности «Поток»",
              "subtitle": "Итоги пилотного внедрения в трёх подразделениях за квартал"},
    "bullet_list": {"type": "bullet_list", "title": "Что мешало собирать отчётность вовремя",
                    "bullets": ["Ручной сбор показателей из семи независимых систем",
                                "Разные форматы выгрузок у каждого подразделения",
                                "Согласование занимало до двух недель"]},
    "stats_kpi": {"type": "stats_kpi", "title": "Результаты пилотного внедрения",
                  "stats": [["-40%", "времени на подготовку регулярной отчётности"],
                            ["+18", "подключённых клиентов ежемесячно"],
                            ["4 дня", "на полное внедрение в подразделении"]]},
    "two_column_comparison": {"type": "two_column_comparison", "title": "Как изменился процесс",
                              "left_heading": "Было",
                              "left_points": ["Ручной сбор показателей", "Разрозненные файлы в почте"],
                              "right_heading": "Стало",
                              "right_points": ["Автоматический сбор по расписанию", "Единая витрина данных"]},
    "image_caption": {"type": "image_caption", "title": "Как выглядит витрина данных",
                      "image": "экран дашборда с графиком отгрузок по неделям"},
    "closing": {"type": "closing", "title": "Запустим пилот в вашем подразделении",
                "subtitle": "Две недели на подключение и обучение команды"},
}

# A third brief, with the TITLES long. The last two defects — a cover title
# running through the balloon artwork and a running header cut to «ЕДИНАЯ
# ПЛАТФОРМА УПРАВЛЕНЧЕСКОЙ» — were both found on decks assembled by hand,
# because LONG and SHORT differ in the length of their BODY text and keep the
# titles short. Anything that depends on a title's length was invisible here.
WORDY = {
    "title": {"type": "title",
              "title": "Единая платформа управленческой отчётности и аналитики для розничной сети",
              "subtitle": "Итоги пилотного внедрения в трёх подразделениях за четвёртый квартал"},
    "bullet_list": {"type": "bullet_list",
                    "title": "Что мешало собирать управленческую отчётность вовремя и без ручной сверки",
                    "bullets": ["Ручной сбор показателей из семи независимых систем",
                                "Разные форматы выгрузок у каждого подразделения",
                                "Согласование занимало до двух недель"]},
    "stats_kpi": {"type": "stats_kpi",
                  "title": "Результаты пилотного внедрения в трёх подразделениях за квартал",
                  "stats": [["-40%", "времени на подготовку регулярной отчётности"],
                            ["+18", "подключённых клиентов ежемесячно"],
                            ["4 дня", "на полное внедрение в подразделении"]]},
    "two_column_comparison": {"type": "two_column_comparison",
                              "title": "Как изменился процесс подготовки управленческой отчётности",
                              "left_heading": "Было",
                              "left_points": ["Ручной сбор показателей", "Разрозненные файлы в почте"],
                              "right_heading": "Стало",
                              "right_points": ["Автоматический сбор по расписанию", "Единая витрина данных"]},
    "image_caption": {"type": "image_caption",
                      "title": "Как выглядит витрина данных в интерфейсе руководителя подразделения",
                      "image": "экран дашборда с графиком отгрузок по неделям"},
    "closing": {"type": "closing",
                "title": "Запустим пилот в вашем подразделении уже в этом квартале",
                "subtitle": "Две недели на подключение и обучение команды"},
}


SHORT = {
    "title": {"type": "title", "title": "Платформа Поток", "subtitle": "Итоги квартала"},
    "bullet_list": {"type": "bullet_list", "title": "Что мешало",
                    "bullets": ["Ручные отчёты", "Разные форматы", "Долгие согласования"]},
    "stats_kpi": {"type": "stats_kpi", "title": "Результаты",
                  "stats": [["-40%", "на отчёт"], ["+18", "клиентов"], ["4 дня", "внедрение"]]},
    "two_column_comparison": {"type": "two_column_comparison", "title": "Было и стало",
                              "left_heading": "Было", "left_points": ["Ручной сбор", "Файлы"],
                              "right_heading": "Стало", "right_points": ["Автосбор", "Витрина"]},
    "image_caption": {"type": "image_caption", "title": "Платформа в работе",
                      "image": "экран дашборда"},
    "closing": {"type": "closing", "title": "Запустим пилот", "subtitle": "Две недели"},
}


def _classified(archetypes):
    return sum(1 for role in archetypes.values() if role != "other")


def _cached_parse(template_id):
    """The archetype map + renders any earlier loop run left behind. Classifying
    needs the network; reusing a cached parse is what makes this offline."""
    best = None
    for model in sorted(os.listdir(LOOP_ROOT)) if os.path.isdir(LOOP_ROOT) else []:
        path = os.path.join(LOOP_ROOT, model, template_id, "archetypes.json")
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8") as f:
            archetypes = {int(k): v for k, v in json.load(f).items()}
        # Skip a parse that is nothing but "other": that is a classification
        # that failed (loop._is_failed_parse), and probing against it measures
        # a deck with no native slides for reasons that have nothing to do with
        # the code under test. Measured: 9 of 20 cached parses were like that.
        if set(archetypes.values()) <= {"other"}:
            continue
        pngs = sorted(glob.glob(os.path.join(LOOP_ROOT, model, template_id, "rendered", "*.png")))
        best = (archetypes, pngs) if best is None or _classified(archetypes) > _classified(best[0]) else best
    return best if best is not None else (None, None)


def _stamp(deck_path, content, items):
    """Write what produced this deck into the file's own properties.

    Twice now a measurement has read a deck from an older run and drawn a
    conclusion from it: once because a skipped template kept its output (fixed
    in iter72 by clearing the directory), and once because an --out directory
    from an earlier session was still on disk with a preset that no longer
    matched. Clearing helps only the directories a run actually touches; a stamp
    travels with the file, so a reader can always ask what it is looking at.
    """
    try:
        sha = subprocess.run(["git", "-C", BASE, "rev-parse", "--short", "HEAD"],
                             capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception:
        sha = "?"
    prs = Presentation(deck_path)
    prs.core_properties.comments = (
        f"{STAMP_MARKER} content={content} items={items or 'plan'} "
        f"at {datetime.datetime.now().isoformat(timespec='seconds')} commit={sha}"
    )
    prs.save(deck_path)


STAMP_MARKER = "offline_probe"


def describe(deck_path):
    """The stamp, for a measurement that wants to be sure what it is reading.

    Matched on the marker rather than on emptiness: python-pptx fills comments
    with «generated using python-pptx» of its own accord, so "not empty" would
    have called a foreign deck stamped."""
    comments = Presentation(deck_path).core_properties.comments or ""
    return comments if comments.startswith(STAMP_MARKER) else "(no stamp)"


def build_plan(spec, outline, blocks):
    """(block, slide_idx) pairs the way the real pipeline builds them.

    Public because every ad-hoc script that assembles a plan by hand has got it
    wrong the same way: iter42 ignored the plan's `count`, iter56 skipped the
    character budgets, and an experiment in iter85 skipped them again and read
    the renderer's own autofit shrink as a product defect. The correct path has
    to be the easy one.
    """
    item_chars = {s["idx"]: s.get("item_chars")
                  for fam in spec["families"] for s in fam["slides"]}
    plan = []
    for item, slide_idx, count in plan_from_outline(outline, spec)[0]:
        block = dict(blocks[item["role"]])
        if count:  # honour the capacity the matcher computed
            for field in ("bullets", "stats", "left_points", "right_points"):
                if field in block:
                    block[field] = list(block[field])[:count]
        block = _enforce_text_budgets(
            block, item["role"], bullet_char_budget(item_chars.get(slide_idx)))
        plan.append((block, slide_idx))
    return plan


def repeat_outline(outline, times):
    """The outline with its CONTENT roles repeated — the axis no preset covers.

    A role used twice sends the matcher back to the same template slide, which
    is the clone path, where the nastiest bugs of this project lived (orphaned
    sldId entries and a notesSlide claimed by several slides, both of which made
    PowerPoint demand repair). One deck of each role never exercises it."""
    if times <= 1:
        return outline
    out = []
    for item in outline:
        out.append(item)
        if item["role"] in ("bullet_list", "stats_kpi", "two_column_comparison"):
            out.extend(dict(item) for _ in range(times - 1))
    return out


def probe(name, template_id, blocks, out_root, outline=OUTLINE, stamp=None):
    src = os.path.join(TEMPLATES, f"{template_id}.pptx")
    archetypes, pngs = _cached_parse(template_id)
    if archetypes is None:
        # A skipped template must leave NOTHING behind. survey-69 has been
        # skipped since iter53 (its cached parses are poisoned) while its deck
        # and eval.json from an older run sat in the output directory — and were
        # read as current during a later measurement. Same trap as iter69, where
        # a crashed run left the previous deck in place and it looked like the
        # fix had done nothing: an artifact that may not have been rewritten is
        # not evidence.
        stale = os.path.join(out_root, name)
        if os.path.isdir(stale):
            shutil.rmtree(stale)
        return {"template": name, "skipped": "нет разобранного шаблона в output/loop"}

    profile = build_measured_profile(pngs, archetypes) if pngs else None
    spec = build_spec(src, archetypes, style_profile=profile)
    plan = build_plan(spec, outline, blocks)

    # Cleared BEFORE generating for the same reason it is cleared on a skip: a
    # run that dies half-way leaves the previous deck and renders in place, and
    # they read as the result of the run that just failed. That happened for
    # real in iter69 — a SyntaxError killed the probe, the old renders were
    # inspected, and the conclusion "the change did nothing" was right only by
    # accident.
    out_dir = os.path.join(out_root, name)
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    deck = os.path.join(out_dir, "deck.pptx")
    generate(src, plan, deck, synth_canvas=_synth_canvas_hints(src, plan, profile),
             canvas_backgrounds=(profile or {}).get("backgrounds"))
    _stamp(deck, *(stamp or ("?", None)))
    renders = render_pptx_to_pngs(deck, os.path.join(out_dir, "render"))
    result = evaluate_deck(deck, brief="offline probe", label=f"probe_{name}",
                           render_dir=os.path.join(out_dir, "render"),
                           out_json=os.path.join(out_dir, "eval.json"))
    native = sum(1 for _, idx in plan if idx != SYNTHESIZE)
    # How much of the TEMPLATE the parse actually classified. Without this the
    # summary's "из шаблона 1/6" reads as "this template offers one usable
    # slide", and that is how it was read for twenty iterations — survey-31's
    # best cached parse labels 1 of its 31 slides and calls the other 30
    # "other", i.e. the classification all but failed. loop._is_failed_parse
    # only rejects a parse where EVERY slide is "other", so this one slips
    # through and every content slide gets synthesized.
    classified = _classified(archetypes)
    weak = {cid: v["score"] for cid, v in result["scores"].items()
            if v.get("score") is not None and v["score"] <= 3}
    return {
        "template": name, "deck": deck, "renders": renders,
        "total_100": result["total_100"],
        "native": f"{native}/{len(plan)}",
        "parse": f"{classified}/{len(archetypes)}",
        "weak": weak,
        "unmeasured": [i + 1 for i, n in
                       enumerate(result["contrast"].get("unmeasured_boxes") or []) if n],
        "unmeasured_why": sorted({w for ws in (result["contrast"].get("unmeasured_why") or [])
                                  for w in ws}),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", choices=("long", "short", "wordy"), default="long")
    ap.add_argument("--only", default=None, help="probe one template by name")
    ap.add_argument("--repeat", type=int, default=1,
                    help="repeat each content role N times — a longer deck, and the only "
                         "way this probe reaches the clone path")
    ap.add_argument("--items", type=int, default=None,
                    help="how many bullets the brief asks for (default: the outline's 3) "
                         "— changes which template slide the matcher picks")
    ap.add_argument("--out", default=OUT_ROOT)
    args = ap.parse_args()

    blocks = {"long": LONG, "short": SHORT, "wordy": WORDY}[args.content]
    outline = OUTLINE
    if args.items:
        blocks = {k: dict(v) for k, v in blocks.items()}
        extra = ["Отчёты собирались вручную в конце месяца",
                 "Данные расходились между системами"]
        pool = list(blocks["bullet_list"]["bullets"]) + extra
        blocks["bullet_list"]["bullets"] = (pool * 3)[:args.items]
        outline = [dict(item, count=args.items) if item["role"] == "bullet_list" else item
                   for item in OUTLINE]
    outline = repeat_outline(outline, args.repeat)
    rows = []
    for name, template_id in REAL_TEMPLATES:
        if args.only and args.only != name:
            continue
        rows.append(probe(name, template_id, blocks, args.out, outline,
                          stamp=(args.content, args.items)))

    print(f"\n# Офлайн-прогон, контент: {args.content}\n")
    for r in rows:
        if r.get("skipped"):
            print(f"- {r['template']:10} пропущен: {r['skipped']}")
            continue
        weak_parse = ""
        if r.get("parse"):
            got, total = (int(x) for x in r["parse"].split("/"))
            weak_parse = f"  разбор {r['parse']:>6}" + ("  ← почти всё «other»" if got * 2 < total else "")
        print(f"- {r['template']:10} итог {r['total_100']:6}  из шаблона {r['native']:>4}{weak_parse}"
              + (f"  слабые: {r['weak']}" if r["weak"] else "  слабых нет")
              + (f"  НЕ ПРОВЕРЕН контраст: {r['unmeasured']}"
                 + (f" ({'; '.join(r['unmeasured_why'])})" if r.get("unmeasured_why") else "")
                 if r["unmeasured"] else ""))
    print("\nПОСМОТРИ ГЛАЗАМИ — числа выше не заменяют взгляд на слайд:")
    for r in rows:
        for path in r.get("renders", []):
            print("    " + path)


if __name__ == "__main__":
    main()
