import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from dotenv import load_dotenv

load_dotenv()

from llm_clients.gigachat import GigaChatClient
from content_parser.parser import parse_brief

BRIEF_PATH = os.path.join(os.path.dirname(__file__), "..", "examples", "product_brief.txt")
OUT_PATH = os.path.join(os.path.dirname(__file__), "..", "output", "parsed", "content_blocks.json")
os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)


def main():
    client = GigaChatClient(verify_ssl=False)
    with open(BRIEF_PATH, encoding="utf-8") as f:
        brief = f.read()

    blocks = parse_brief(client, brief)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(blocks, f, ensure_ascii=False, indent=2)

    print(f"{len(blocks)} content blocks -> {OUT_PATH}")
    for b in blocks:
        print(" -", b["type"], ":", b.get("title"))


if __name__ == "__main__":
    main()
