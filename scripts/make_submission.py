#!/usr/bin/env python3
"""Девять презентаций для сдачи: 3 варианта вёрстки × 3 шаблона, один контент.

Требование ТЗ (раздел 6): «результат демонстрации сервиса — три варианта
презентации, сгенерированные на одном контенте на трёх разных шаблонах (всего 9
вариантов презентаций)».

Запуск:
    set -a; . ./.env; set +a
    .venv/bin/python3 scripts/make_submission.py                # все девять
    .venv/bin/python3 scripts/make_submission.py --template vk_tech --variant compact

Результат — output/submission/<шаблон>/<вариант>.pptx, рядом README.md с
описанием оси различий и таблицей времени генерации каждой колоды.
"""

import argparse
import json
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))

from content_package import load_package, to_brief_text  # noqa: E402
from content_parser.variants import VARIANTS, variant_slides  # noqa: E402
from evaluation.loop import ensure_template, generate_deck  # noqa: E402
from llm_clients.backends import open_weights_client  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES = os.path.join(ROOT, "output", "templates")
OUT = os.path.join(ROOT, "output", "submission")
PACKAGE = os.path.join(ROOT, "samples", "content_packages", "feature_smart_search")

TEMPLATE_TITLES = {
    "vk_tech": "VK Tech",
    "vk_workspace": "VK WorkSpace",
    "vk_education": "VK Education",
}


def _client(model_arg):
    """Открытая модель из LLM_BASE_URL; без неё — запасной провайдер разработки."""
    configured = open_weights_client()
    if configured:
        return configured
    from llm_clients.gigachat import GigaChatClient
    print("! LLM_BASE_URL не задан — колоды собираются на запасном провайдере "
          "разработки; для сдачи перегенерировать на открытой модели", flush=True)
    return GigaChatClient(verify_ssl=False), (model_arg or "GigaChat-2-Max")


def build(template, variant, client, model, brief, slides):
    src = os.path.join(TEMPLATES, f"{template}.pptx")
    t = ensure_template(client, model, src, template)
    out_dir = os.path.join(OUT, template)
    os.makedirs(out_dir, exist_ok=True)
    out_pptx = os.path.join(out_dir, f"{variant}.pptx")

    started = time.monotonic()
    plan, skipped = generate_deck(client, model, src, t["spec"], brief,
                                  t["style_preamble"], out_pptx, profile=t["profile"],
                                  slides=variant_slides(variant, slides),
                                  variant=variant)
    seconds = round(time.monotonic() - started, 1)
    print(f"  {template}/{variant}: {len(plan)} слайдов за {seconds} с"
          + (f", пропущено {len(skipped)}" if skipped else ""), flush=True)
    return {"template": template, "variant": variant, "slides": len(plan),
            "seconds": seconds, "path": os.path.relpath(out_pptx, ROOT)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", action="append", choices=list(TEMPLATE_TITLES))
    ap.add_argument("--variant", action="append", choices=list(VARIANTS))
    ap.add_argument("--model", default=None, help="имя модели у провайдера")
    ap.add_argument("--slides", type=int, default=None,
                    help="заказать длину явно; по умолчанию её выбирает вариант")
    args = ap.parse_args()

    templates = args.template or list(TEMPLATE_TITLES)
    variants = args.variant or list(VARIANTS)

    package = load_package(PACKAGE)
    brief = to_brief_text(package)
    # По умолчанию длину выбирает ВАРИАНТ внутри диапазона ТЗ (10–15):
    # иначе фиксированный заказ пакета прижимает все три к нижней границе.
    slides = args.slides
    numbers = package["numbers"]

    client, model = _client(args.model)
    print(f"модель: {model}; контент: {os.path.basename(PACKAGE)}; "
          f"{len(templates)}×{len(variants)} = {len(templates) * len(variants)} колод", flush=True)

    os.makedirs(OUT, exist_ok=True)
    rows, failed = [], []
    for template in templates:
        for variant in variants:
            try:
                rows.append(build(template, variant, client, model, brief, slides))
            except Exception as e:  # noqa: BLE001 — одна колода не должна валить прогон
                print(f"  ! {template}/{variant}: {e}", flush=True)
                failed.append({"template": template, "variant": variant, "error": str(e)})

    with open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump({"model": model, "package": os.path.basename(PACKAGE),
                   "decks": rows, "failed": failed}, f, ensure_ascii=False, indent=2)

    lines = ["# Девять презентаций: 3 варианта × 3 шаблона", "",
             f"Один и тот же контент — пакет «{os.path.basename(PACKAGE)}», "
             f"модель — {model}.", "",
             "## Ось различий", ""]
    for name, spec in VARIANTS.items():
        lines.append(f"- **{spec['title']}** (`{name}`) — {spec['description']}.")
    lines += ["", "Различается только ПЛАН колоды: состав слайдов, их число и порядок.",
              "Вёрстка у всех трёх одна и та же, поэтому ни один вариант не соблюдает",
              "правила шаблона хуже другого.", "", "## Собранные колоды", "",
              "| Шаблон | Вариант | Слайдов | Время сборки |", "|---|---|---|---|"]
    for r in rows:
        lines.append(f"| {TEMPLATE_TITLES[r['template']]} | {VARIANTS[r['variant']]['title']} "
                     f"| {r['slides']} | {r['seconds']} с |")
    if failed:
        lines += ["", "Не собрались: " + ", ".join(f"{f['template']}/{f['variant']}" for f in failed)]
    with open(os.path.join(OUT, "README.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    print(f"\nготово: {len(rows)} колод в {os.path.relpath(OUT, ROOT)}"
          + (f", не собралось {len(failed)}" if failed else ""))
    return 0 if rows else 1


if __name__ == "__main__":
    sys.exit(main())
