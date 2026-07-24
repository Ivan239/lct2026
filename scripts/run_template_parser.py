import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from template_parser.parser import extract_template

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "templates")
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "parsed")
os.makedirs(OUT_DIR, exist_ok=True)


def main():
    for fname in sorted(os.listdir(TEMPLATES_DIR)):
        if not fname.endswith(".pptx"):
            continue
        path = os.path.join(TEMPLATES_DIR, fname)
        spec = extract_template(path)
        out_path = os.path.join(OUT_DIR, fname.replace(".pptx", ".json"))
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(spec, f, ensure_ascii=False, indent=2)

        print(f"=== {fname} ===")
        print("Palette:", spec["theme"]["palette"])
        print("Fonts:", spec["theme"]["fonts"])
        print(f"{len(spec['slides'])} slides parsed, saved to {out_path}")
        for slide in spec["slides"]:
            texts = [
                run["text"]
                for shape in slide["shapes"]
                if shape["text"]
                for para in shape["text"]
                for run in para
                if run["text"]
            ]
            print(f"  slide {slide['index']} [{slide['layout_name']}]: {len(slide['shapes'])} shapes, text: {texts[:3]}")
        print()


if __name__ == "__main__":
    main()
