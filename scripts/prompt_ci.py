"""Prompt-CI (docs/IMPROVEMENT_PLAN.md item 5): measures how reliably each
GigaChat model follows the two-phase prompts, using the same validators the
production pipeline enforces.

SPENDS REAL GIGACHAT TOKENS — run deliberately, not from tests/CI hooks:

    .venv/bin/python3 scripts/prompt_ci.py --runs 5 --models GigaChat-2 GigaChat-2-Pro

Per run: 1 outline call + 1 bullet_list block + 1 stats_kpi block (so
--runs 5 on two models ≈ 30 small completions). Prints a pass-rate table per
model and stage; exits non-zero if any cell is below --threshold.
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import urllib3

from dotenv import load_dotenv

load_dotenv()
# verify_ssl=False is deliberate on this network path (see CLAUDE.md) — the
# per-request warning would drown the actual pass/fail output.
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from content_parser.two_phase import generate_block, generate_outline
from llm_clients.gigachat import GigaChatClient
from template_spec.builder import build_spec

BRIEF = """Продукт: СлайдоГен — сервис автоматической генерации презентаций в фирменном
стиле компании. Проблема: компании тратят часы на ручную адаптацию контента под
корпоративный шаблон, дизайнеры перегружены вёрсткой. Решение: сервис разбирает
PPTX-шаблон, извлекает дизайн-систему и собирает новую презентацию из контента
автоматически. Метрики пилота: x10 ускорение подготовки, 3 компании, 0 часов
ручной работы."""

TEMPLATE = os.path.join(
    os.path.dirname(__file__), "..", "output", "templates", "template_a_corporate.pptx"
)
ARCHETYPES = {0: "title", 1: "bullet_list", 2: "two_column_comparison", 3: "stats_kpi"}

STAGES = ("outline", "bullet_list", "stats_kpi")


def run_stage(client, stage, spec, model):
    if stage == "outline":
        outline = generate_outline(client, BRIEF, spec, models=[model])
        assert outline[0]["role"] == "title"
    elif stage == "bullet_list":
        generate_block(client, "bullet_list", "проблемы компаний", BRIEF, count=4, models=[model])
    else:
        generate_block(client, "stats_kpi", "метрики пилота", BRIEF, count=3, models=[model])


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--runs", type=int, default=3)
    parser.add_argument("--models", nargs="+", default=["GigaChat-2", "GigaChat-2-Pro"])
    parser.add_argument("--threshold", type=float, default=0.8)
    args = parser.parse_args()

    client = GigaChatClient(verify_ssl=False)
    spec = build_spec(TEMPLATE, ARCHETYPES)

    results = {}  # (model, stage) -> [ok, ...]
    for model in args.models:
        for stage in STAGES:
            for i in range(args.runs):
                try:
                    run_stage(client, stage, spec, model)
                    ok = True
                except Exception as e:
                    ok = False
                    print(f"  FAIL {model}/{stage} run {i + 1}: {e}", flush=True)
                results.setdefault((model, stage), []).append(ok)
                print(f"  {model:16s} {stage:12s} run {i + 1}: {'ok' if ok else 'FAIL'}", flush=True)

    print(f"\n{'model':16s} {'stage':12s} pass-rate  (runs={args.runs}, "
          "валидация включает ретраи пайплайна — это доля УСПЕШНЫХ итогов, не первых ответов)")
    worst = 1.0
    for (model, stage), oks in sorted(results.items()):
        rate = sum(oks) / len(oks)
        worst = min(worst, rate)
        print(f"{model:16s} {stage:12s} {rate:8.0%}")

    if worst < args.threshold:
        print(f"\nworst pass-rate {worst:.0%} < threshold {args.threshold:.0%}")
        sys.exit(1)


if __name__ == "__main__":
    main()
