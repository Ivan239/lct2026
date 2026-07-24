import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from matcher.matcher import match_content_to_slides
from generator.generator import generate

BASE = os.path.join(os.path.dirname(__file__), "..")
PARSED_DIR = os.path.join(BASE, "output", "parsed")
TEMPLATES_DIR = os.path.join(BASE, "output", "templates")
OUT_DIR = os.path.join(BASE, "output", "generated")
os.makedirs(OUT_DIR, exist_ok=True)

TEMPLATE_NAME = sys.argv[1] if len(sys.argv) > 1 else "template_a_corporate"


def main():
    with open(os.path.join(PARSED_DIR, "content_blocks.json"), encoding="utf-8") as f:
        content_blocks = json.load(f)
    with open(os.path.join(PARSED_DIR, f"{TEMPLATE_NAME}_archetypes.json"), encoding="utf-8") as f:
        archetype_map = {int(k): v for k, v in json.load(f).items()}

    plan, skipped = match_content_to_slides(content_blocks, archetype_map)

    print("Plan:")
    for block, idx in plan:
        print(f"  slide {idx} [{archetype_map[idx]}] <- {block['type']}: {block.get('title')}")
    if skipped:
        print("Skipped (no matching template slide available):")
        for b in skipped:
            print(f"  {b['type']}: {b.get('title')}")

    template_path = os.path.join(TEMPLATES_DIR, f"{TEMPLATE_NAME}.pptx")
    out_path = os.path.join(OUT_DIR, f"{TEMPLATE_NAME}_generated.pptx")
    generate(template_path, plan, out_path)
    print("\nSaved:", out_path)


if __name__ == "__main__":
    main()
