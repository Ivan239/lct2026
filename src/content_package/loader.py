"""Content package: import and decomposition (VK Tech brief, «Импорт и
декомпозиция … контент-пакетов»).

The organizers do not ship a content package (confirmed 19.09.2026), so the
format is ours — documented in docs/CONTENT_PACKAGE.md. In short, a folder or a
.zip of one:

    package.json   optional: title, purpose, slides, audience, tone, language,
                   image captions
    brief.md       REQUIRED: the author's text — what the deck is for
    facts.md       optional: one fact per line; the numbers in it (and in the
                   brief) are the source of truth for the audit question «все
                   цифры и факты со слайда есть в исходных материалах?»
    data/*.csv     optional: tables — the source for charts and native tables
    images/*       optional: png / jpg / jpeg / webp to place on slides

Decomposition keeps each kind of material apart (brief, facts, numbers,
tables, images) instead of one blob of text: charts need the tables as
tables, the audit needs the numbers as a set, and the layout step needs to know
which pictures exist. `to_brief_text` flattens it back into one brief so the
current text pipeline (content_parser.two_phase, which takes a string) can run
on a package today, before the downstream steps learn to use the parts.
"""

import csv
import io
import json
import os
import re
import tempfile
import zipfile

# The four purposes the VK Tech brief names («по краткому брифу и назначению
# (фича, продукт, проект, инициатива)»). Russian labels are what the prompt
# sees; the keys are what a package declares.
PURPOSES = {
    "feature": "фича",
    "product": "продукт",
    "project": "проект",
    "initiative": "инициатива",
}

IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp")

# Target deck size from the brief: 10-15 slides, or what the user asks for.
# The bounds here only reject nonsense; the generator enforces the real limits.
MIN_SLIDES, MAX_SLIDES = 3, 30

# A number as it appears in business prose: signed, with thousands separated
# by spaces, a decimal comma or point, and an optional unit that belongs to it.
# Numbers are compared as normalised strings, so «+25%» and «25 %» match.
_NUMBER = re.compile(
    r"(?<![\w.,])[+\-−–]?\d{1,3}(?:[  ]\d{3})+(?:[.,]\d+)?"
    r"|(?<![\w.,])[+\-−–]?\d+(?:[.,]\d+)?")
_UNIT = re.compile(r"\s?(%|‰|x|×|раз[а]?|млн|млрд|тыс\.?|₽|руб\.?|\$|€|"
                   r"ч(?:ас(?:а|ов)?)?|дн(?:я|ей)?|мин(?:ут)?|сек|мес(?:\.|яц(?:а|ев)?)?|лет|год(?:а)?)",
                   re.IGNORECASE)


class ContentPackageError(ValueError):
    """The package cannot be used; the message says what is wrong and where."""


def load_package(path):
    """Load a content package from a folder or a .zip and decompose it.

    Returns a dict:
        title, purpose, purpose_label, slides, audience, tone, language,
        brief (str), facts (list[str]), numbers (sorted list[str]),
        tables (list of {name, columns, rows, numeric_columns}),
        images (list of {path, caption}),
        root (folder the package was read from)

    A .zip is extracted into a temporary folder that lives as long as the
    process; image paths point into it.
    """
    if not os.path.exists(path):
        raise ContentPackageError(f"контент-пакет не найден: {path}")
    root = _extract_zip(path) if zipfile.is_zipfile(path) else path
    if not os.path.isdir(root):
        raise ContentPackageError(f"контент-пакет — это папка или .zip, а не файл: {path}")
    root = _single_top_folder(root)

    meta = _read_manifest(root)
    brief_path = os.path.join(root, "brief.md")
    if not os.path.isfile(brief_path):
        raise ContentPackageError(
            f"в контент-пакете нет brief.md (обязательный файл): {root}")
    brief = _read_text(brief_path).strip()
    if not brief:
        raise ContentPackageError(f"brief.md пустой: {brief_path}")

    facts = _read_facts(os.path.join(root, "facts.md"))
    tables = _read_tables(os.path.join(root, "data"))
    images = _read_images(os.path.join(root, "images"), meta.get("images") or {})

    purpose = meta.get("purpose")
    if purpose is not None and purpose not in PURPOSES:
        raise ContentPackageError(
            f"package.json: purpose={purpose!r}, допустимо: {', '.join(PURPOSES)}")
    slides = meta.get("slides")
    if slides is not None:
        if not isinstance(slides, int) or not MIN_SLIDES <= slides <= MAX_SLIDES:
            raise ContentPackageError(
                f"package.json: slides={slides!r}, нужно целое от {MIN_SLIDES} до {MAX_SLIDES}")

    numbers = set(extract_numbers(brief))
    for fact in facts:
        numbers.update(extract_numbers(fact))
    for table in tables:
        for row in table["rows"]:
            for cell in row:
                numbers.update(extract_numbers(cell))

    return {
        "title": meta.get("title") or _first_heading(brief),
        "purpose": purpose,
        "purpose_label": PURPOSES.get(purpose),
        "slides": slides,
        "audience": meta.get("audience"),
        "tone": meta.get("tone"),
        "language": meta.get("language") or "ru",
        "brief": brief,
        "facts": facts,
        "numbers": sorted(numbers),
        "tables": tables,
        "images": images,
        "root": root,
    }


def extract_numbers(text):
    """Normalised numbers in `text`, units kept: «+25 %» → «+25%»,
    «20 227 000» → «20227000», «3,5 млн» → «3.5млн»."""
    out = []
    for match in _NUMBER.finditer(text or ""):
        raw = match.group(0)
        unit = _UNIT.match(text, match.end())
        number = raw.replace(" ", "").replace(" ", "").replace(",", ".")
        number = number.replace("−", "-").replace("–", "-")
        if unit:
            number += unit.group(1).lower().rstrip(".")
        out.append(number)
    return out


def to_brief_text(package):
    """One brief string for the current text pipeline (two_phase takes a
    string). Keeps the parts labelled so the model can tell the author's
    request from the facts it must stay within."""
    lines = []
    if package.get("title"):
        lines.append(f"Тема: {package['title']}")
    if package.get("purpose_label"):
        lines.append(f"Назначение: {package['purpose_label']}")
    if package.get("audience"):
        lines.append(f"Аудитория: {package['audience']}")
    if package.get("tone"):
        lines.append(f"Тон: {package['tone']}")
    if package.get("slides"):
        lines.append(f"Объём: {package['slides']} слайдов")
    lines.append("")
    lines.append(package["brief"])
    if package.get("facts"):
        lines.append("")
        lines.append("Факты (используй только их, цифр сверх них не придумывай):")
        lines.extend(f"- {fact}" for fact in package["facts"])
    for table in package.get("tables") or []:
        lines.append("")
        lines.append(f"Таблица «{table['name']}»: {', '.join(table['columns'])}")
        for row in table["rows"][:12]:
            lines.append("  " + " | ".join(row))
    if package.get("images"):
        lines.append("")
        lines.append("Изображения в пакете:")
        lines.extend(f"- {img['caption']}" for img in package["images"])
    return "\n".join(lines).strip()


# ---------------------------------------------------------------------------

def _extract_zip(path):
    target = tempfile.mkdtemp(prefix="content_package_")
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            name = info.filename
            # A zip made on Windows without the UTF-8 flag carries cp866 names.
            if not info.flag_bits & 0x800:
                try:
                    name = name.encode("cp437").decode("cp866")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    pass
            if name.startswith("/") or ".." in name.split("/"):
                raise ContentPackageError(f"небезопасный путь в архиве: {name!r}")
            if name.startswith("__MACOSX/") or os.path.basename(name).startswith("._"):
                continue
            dest = os.path.join(target, name)
            if info.is_dir():
                os.makedirs(dest, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            with zf.open(info) as src, open(dest, "wb") as out:
                out.write(src.read())
    return target


def _single_top_folder(root):
    """Zipping a folder puts everything one level down; look through it."""
    if os.path.isfile(os.path.join(root, "brief.md")):
        return root
    entries = [e for e in os.listdir(root) if not e.startswith(".")]
    if len(entries) == 1 and os.path.isdir(os.path.join(root, entries[0])):
        return os.path.join(root, entries[0])
    return root


def _read_text(path):
    with open(path, encoding="utf-8-sig") as f:
        return f.read()


def _read_manifest(root):
    path = os.path.join(root, "package.json")
    if not os.path.isfile(path):
        return {}
    try:
        data = json.loads(_read_text(path))
    except json.JSONDecodeError as e:
        raise ContentPackageError(f"package.json не читается: {e}") from e
    if not isinstance(data, dict):
        raise ContentPackageError("package.json должен быть объектом {…}")
    return data


def _read_facts(path):
    if not os.path.isfile(path):
        return []
    facts = []
    for line in _read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        facts.append(re.sub(r"^([-*•]|\d+[.)])\s+", "", line))
    return facts


def _read_tables(folder):
    if not os.path.isdir(folder):
        return []
    tables = []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith(".csv"):
            continue
        text = _read_text(os.path.join(folder, name))
        delimiter = _detect_delimiter(text)
        rows = [[c.strip() for c in r] for r in csv.reader(io.StringIO(text), delimiter=delimiter)
                if any(c.strip() for c in r)]
        if not rows:
            continue
        columns, body = rows[0], rows[1:]
        numeric = [col for i, col in enumerate(columns)
                   if body and all(_is_number(r[i]) for r in body if i < len(r) and r[i])]
        tables.append({"name": os.path.splitext(name)[0], "columns": columns,
                       "rows": body, "numeric_columns": numeric})
    return tables


def _detect_delimiter(text):
    """The delimiter that gives every row the same number of columns (>1).

    csv.Sniffer guesses wrong on the most common Russian CSV: «;» between
    fields and a decimal COMMA inside them («6,2»). It picked the comma, and a
    three-column budget table came back as «Месяц;План | млн ₽;Факт | млн ₽».
    Tab and semicolon are tried before the comma for the same reason: when both
    are consistent, the comma is the one that also serves as a decimal mark."""
    lines = [line for line in text.splitlines() if line.strip()][:20]
    for delimiter in ("\t", ";", ","):
        counts = {len(next(csv.reader([line], delimiter=delimiter))) for line in lines}
        if len(counts) == 1 and counts.pop() > 1:
            return delimiter
    return ","


def _is_number(cell):
    return bool(re.fullmatch(r"[+\-−–]?[\d  ]+(?:[.,]\d+)?\s?%?", cell.strip()))


def _read_images(folder, captions):
    if not os.path.isdir(folder):
        return []
    images = []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith(IMAGE_EXTENSIONS):
            continue
        stem = os.path.splitext(name)[0]
        caption = captions.get(name) or re.sub(r"[_\-]+", " ", stem).strip()
        images.append({"path": os.path.join(folder, name), "caption": caption})
    return images


def _first_heading(brief):
    for line in brief.splitlines():
        line = line.strip().lstrip("#").strip()
        if line:
            return line[:120]
    return None
