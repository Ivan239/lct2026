import glob
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from functools import lru_cache

SOFFICE_CANDIDATES = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]

# Where to look for fonts the render host actually has. Families referenced by
# a deck but absent here get substituted in the PREVIEW ONLY (see
# _prepare_render_copy) — LibreOffice's own fallback for a missing family
# produces broken glyph spacing (overlapping letters), which reads as a
# generation bug to anyone looking at the preview even though the .pptx the
# user downloads is untouched and correct.
FONT_DIRS = [
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts",
]
FALLBACK_TYPEFACE = "Arial"

_TYPEFACE_RE = re.compile(rb'typeface="([^"]+)"')


def _find_soffice():
    for candidate in SOFFICE_CANDIDATES:
        try:
            subprocess.run([candidate, "--version"], capture_output=True, check=True)
            return candidate
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
    raise RuntimeError("soffice not found; is LibreOffice installed?")


def _normalize(family):
    return re.sub(r"\s+", " ", family).strip().lower()


@lru_cache(maxsize=1)
def _installed_families():
    """Real font family names (from the fonts' own name tables, not filenames)
    of everything installed on this machine, normalized for comparison."""
    from fontTools.ttLib import TTCollection, TTFont

    families = set()
    for font_dir in FONT_DIRS:
        for path in glob.glob(os.path.join(font_dir, "*")):
            try:
                if path.lower().endswith(".ttc"):
                    fonts = TTCollection(path, lazy=True).fonts
                else:
                    fonts = [TTFont(path, lazy=True)]
                for font in fonts:
                    for name_id in (16, 1):  # typographic family, then legacy family
                        value = font["name"].getDebugName(name_id)
                        if value:
                            families.add(_normalize(value))
            except Exception:
                continue  # not a parseable font — irrelevant
    return families


def _referenced_typefaces(pptx_path):
    faces = set()
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            if not name.endswith(".xml"):
                continue
            for match in _TYPEFACE_RE.finditer(z.read(name)):
                face = match.group(1).decode("utf-8", errors="ignore")
                if not face.startswith("+"):  # theme tokens resolve via theme1.xml, listed separately
                    faces.add(face)
    return faces


def _prepare_render_copy(pptx_path, tmp_dir):
    """Returns a path to render from: the original file when every referenced
    font is installed, otherwise a copy (same basename, so downstream PDF/PNG
    naming is unchanged) with missing families rewritten to FALLBACK_TYPEFACE.
    Preview-only — the original file is never modified."""
    missing = {f for f in _referenced_typefaces(pptx_path) if _normalize(f) not in _installed_families()}
    if not missing:
        return pptx_path

    copy_path = os.path.join(tmp_dir, os.path.basename(pptx_path))
    with zipfile.ZipFile(pptx_path) as zin, zipfile.ZipFile(copy_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                for face in missing:
                    data = data.replace(
                        b'typeface="%s"' % face.encode("utf-8"),
                        b'typeface="%s"' % FALLBACK_TYPEFACE.encode("utf-8"),
                    )
            zout.writestr(item, data)
    return copy_path


def convert_to_pdf(pptx_path, out_dir):
    """pptx → {out_dir}/{name}.pdf через LibreOffice. Недостающие на хосте
    шрифты подменяются в КОПИИ (см. _prepare_render_copy) — исходный .pptx не
    трогается. Возвращает путь к PDF."""
    os.makedirs(out_dir, exist_ok=True)
    soffice = _find_soffice()
    name = os.path.splitext(os.path.basename(pptx_path))[0]

    tmp_dir = tempfile.mkdtemp(prefix="render_")
    try:
        render_source = _prepare_render_copy(pptx_path, tmp_dir)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, render_source],
            check=True,
            capture_output=True,
        )
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return os.path.join(out_dir, f"{name}.pdf")


# Тема перечисляет шрифт на КАЖДУЮ письменность (<a:font script="Thai"
# typeface="Angsana New"/>): на VK Tech это 28 семейств, которыми наш текст не
# набран. Подменять их безвредно, а называть пользователю — шум, в котором
# теряются два настоящих (Calibri, Play).
_SCRIPT_FONT_RE = re.compile(rb'<a:font script="[^"]*" typeface="[^"]*"\s*/>')


def substituted_typefaces(pptx_path):
    """Семейства, которых нет на этом хосте и которые в PDF/PNG заменены на
    FALLBACK_TYPEFACE. Пусто — рендер честный, как в PowerPoint."""
    faces = set()
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            if name.endswith(".xml"):
                data = _SCRIPT_FONT_RE.sub(b"", z.read(name))
                for match in _TYPEFACE_RE.finditer(data):
                    face = match.group(1).decode("utf-8", errors="ignore")
                    if face and not face.startswith("+"):
                        faces.add(face)
    return sorted(f for f in faces if _normalize(f) not in _installed_families())


def render_pptx_to_pngs(pptx_path, out_dir):
    """Converts a pptx to one PNG per slide in out_dir, returns sorted list of PNG paths.
    Side effect used by export: {out_dir}/{name}.pdf stays next to the PNGs."""
    name = os.path.splitext(os.path.basename(pptx_path))[0]
    pdf_path = convert_to_pdf(pptx_path, out_dir)
    subprocess.run(
        ["pdftoppm", "-png", "-r", "150", pdf_path, os.path.join(out_dir, name)],
        check=True,
        capture_output=True,
    )
    return sorted(glob.glob(os.path.join(out_dir, f"{name}-*.png")))
