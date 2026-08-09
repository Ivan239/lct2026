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
import glob
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

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


def _cached_parse(template_id):
    """The archetype map + renders any earlier loop run left behind. Classifying
    needs the network; reusing a cached parse is what makes this offline."""
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
        return archetypes, pngs
    return None, None


def probe(name, template_id, blocks, out_root, outline=OUTLINE):
    src = os.path.join(TEMPLATES, f"{template_id}.pptx")
    archetypes, pngs = _cached_parse(template_id)
    if archetypes is None:
        return {"template": name, "skipped": "нет разобранного шаблона в output/loop"}

    profile = build_measured_profile(pngs, archetypes) if pngs else None
    spec = build_spec(src, archetypes, style_profile=profile)
    assignments, _ = plan_from_outline(outline, spec)

    item_chars = {s["idx"]: s.get("item_chars")
                  for fam in spec["families"] for s in fam["slides"]}
    plan = []
    for item, slide_idx, count in assignments:
        block = dict(blocks[item["role"]])
        if count:  # honour the capacity the matcher computed — see module docstring
            for field in ("bullets", "stats", "left_points", "right_points"):
                if field in block:
                    block[field] = list(block[field])[:count]
        # …and the character budget, which the real pipeline applies to every
        # block (two_phase.generate_block). Skipping it made the probe ship a
        # 49-character bullet into a slot whose budget is 28 and then report the
        # 9pt microtext that followed as if it were a product defect.
        block = _enforce_text_budgets(
            block, item["role"], bullet_char_budget(item_chars.get(slide_idx)))
        plan.append((block, slide_idx))

    out_dir = os.path.join(out_root, name)
    os.makedirs(out_dir, exist_ok=True)
    deck = os.path.join(out_dir, "deck.pptx")
    generate(src, plan, deck, synth_canvas=_synth_canvas_hints(src, plan, profile),
             canvas_backgrounds=(profile or {}).get("backgrounds"))
    renders = render_pptx_to_pngs(deck, os.path.join(out_dir, "render"))
    result = evaluate_deck(deck, brief="offline probe", label=f"probe_{name}",
                           render_dir=os.path.join(out_dir, "render"),
                           out_json=os.path.join(out_dir, "eval.json"))
    native = sum(1 for _, idx in plan if idx != SYNTHESIZE)
    weak = {cid: v["score"] for cid, v in result["scores"].items()
            if v.get("score") is not None and v["score"] <= 3}
    return {
        "template": name, "deck": deck, "renders": renders,
        "total_100": result["total_100"],
        "native": f"{native}/{len(plan)}",
        "weak": weak,
        "unmeasured": [i + 1 for i, n in
                       enumerate(result["contrast"].get("unmeasured_boxes") or []) if n],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--content", choices=("long", "short"), default="long")
    ap.add_argument("--only", default=None, help="probe one template by name")
    ap.add_argument("--items", type=int, default=None,
                    help="how many bullets the brief asks for (default: the outline's 3) "
                         "— changes which template slide the matcher picks")
    ap.add_argument("--out", default=OUT_ROOT)
    args = ap.parse_args()

    blocks = LONG if args.content == "long" else SHORT
    outline = OUTLINE
    if args.items:
        blocks = {k: dict(v) for k, v in blocks.items()}
        extra = ["Отчёты собирались вручную в конце месяца",
                 "Данные расходились между системами"]
        pool = list(blocks["bullet_list"]["bullets"]) + extra
        blocks["bullet_list"]["bullets"] = (pool * 3)[:args.items]
        outline = [dict(item, count=args.items) if item["role"] == "bullet_list" else item
                   for item in OUTLINE]
    rows = []
    for name, template_id in REAL_TEMPLATES:
        if args.only and args.only != name:
            continue
        rows.append(probe(name, template_id, blocks, args.out, outline))

    print(f"\n# Офлайн-прогон, контент: {args.content}\n")
    for r in rows:
        if r.get("skipped"):
            print(f"- {r['template']:10} пропущен: {r['skipped']}")
            continue
        print(f"- {r['template']:10} итог {r['total_100']:6}  из шаблона {r['native']:>4}"
              + (f"  слабые: {r['weak']}" if r["weak"] else "  слабых нет")
              + (f"  НЕ ПРОВЕРЕН контраст: {r['unmeasured']}" if r["unmeasured"] else ""))
    print("\nПОСМОТРИ ГЛАЗАМИ — числа выше не заменяют взгляд на слайд:")
    for r in rows:
        for path in r.get("renders", []):
            print("    " + path)


if __name__ == "__main__":
    main()
