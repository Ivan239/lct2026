import glob
import os
import subprocess
import sys

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "templates")
RENDER_DIR = os.path.join(os.path.dirname(__file__), "..", "output", "rendered")

SOFFICE_CANDIDATES = [
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice():
    for candidate in SOFFICE_CANDIDATES:
        try:
            subprocess.run([candidate, "--version"], capture_output=True, check=True)
            return candidate
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
    raise RuntimeError("soffice not found; is LibreOffice installed?")


def render_pptx_to_pdf(soffice, pptx_path, out_dir):
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
        check=True,
        capture_output=True,
    )
    return os.path.join(out_dir, os.path.basename(pptx_path).replace(".pptx", ".pdf"))


def pdf_to_pngs(pdf_path, out_dir, name):
    # pdftoppm ships with poppler; fall back to soffice PNG export (single-page only) if missing.
    try:
        subprocess.run(
            ["pdftoppm", "-png", "-r", "150", pdf_path, os.path.join(out_dir, name)],
            check=True,
            capture_output=True,
        )
        return sorted(glob.glob(os.path.join(out_dir, f"{name}-*.png")))
    except FileNotFoundError:
        return None


def main():
    soffice = find_soffice()
    os.makedirs(RENDER_DIR, exist_ok=True)

    for pptx_path in sorted(glob.glob(os.path.join(TEMPLATES_DIR, "*.pptx"))):
        name = os.path.basename(pptx_path).replace(".pptx", "")
        print(f"Rendering {name} ...")
        pdf_path = render_pptx_to_pdf(soffice, pptx_path, RENDER_DIR)
        pages = pdf_to_pngs(pdf_path, RENDER_DIR, name)
        if pages:
            print(f"  {len(pages)} PNG pages -> {RENDER_DIR}")
        else:
            print(f"  PDF only (install poppler for per-slide PNGs): {pdf_path}")


if __name__ == "__main__":
    main()
