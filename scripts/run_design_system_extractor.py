import glob
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dotenv import load_dotenv

load_dotenv()

from llm_clients.gigachat import GigaChatClient
from design_system.extractor import build_archetype_map

RENDER_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "rendered")
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "parsed")
os.makedirs(OUT_DIR, exist_ok=True)

TEMPLATE_NAMES = ["template_a_corporate", "template_b_startup"]


def main():
    client = GigaChatClient(verify_ssl=False)
    for name in TEMPLATE_NAMES:
        pages = sorted(glob.glob(os.path.join(RENDER_DIR, f"{name}-*.png")))
        archetype_map = build_archetype_map(client, pages)
        out_path = os.path.join(OUT_DIR, f"{name}_archetypes.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(archetype_map, f, ensure_ascii=False, indent=2)
        print(name, "->", archetype_map)


if __name__ == "__main__":
    main()
