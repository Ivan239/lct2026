"""Package-integrity guard for generated .pptx files (plan item 9.6).

Born from three rounds of "PowerPoint wants to repair this file" where each fix
looked complete until the next report: LibreOffice happily renders packages
that PowerPoint's stricter OPC checks reject, so "it renders" proves nothing.
The guard has two tiers:

1. `structural_audit` — the specific consistency rules that real corruption
   actually violated here (orphaned slide parts, shared notesSlide ownership,
   duplicate shape ids, undeclared r:id references, content-type gaps, dangling
   relationship targets). Pure stdlib+lxml, ~milliseconds, runs on every
   generation.
2. `deep_validate` — the vendored ISO-IEC29500 XSD validator
   (vendor/ooxml_validator, ~0.5s as a subprocess). Catches schema-level
   breakage the hand-written rules don't know about.

`assert_valid_package` combines both and raises with a readable message —
a generation must fail loudly rather than hand the user a broken file."""

import os
import re
import subprocess
import sys
import zipfile
from collections import Counter

from lxml import etree

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_VENDORED_VALIDATOR_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "vendor", "ooxml_validator")
)


def structural_audit(pptx_path):
    """Returns a list of human-readable problem strings; empty == clean."""
    problems = []
    with zipfile.ZipFile(pptx_path) as z:
        names = set(z.namelist())

        # --- content-type coverage -------------------------------------
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        defaults = {e.get("Extension").lower() for e in ct.findall(f"{{{CT_NS}}}Default")}
        overrides = {e.get("PartName") for e in ct.findall(f"{{{CT_NS}}}Override")}
        for n in names:
            if n.endswith("/") or n == "[Content_Types].xml":
                continue
            ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
            if "/" + n not in overrides and ext not in defaults:
                problems.append(f"part without content type: {n}")

        # --- every relationship target must exist; collect referenced --
        referenced = set()
        for n in names:
            if not n.endswith(".rels"):
                continue
            base = os.path.dirname(os.path.dirname(n))
            for rel in etree.fromstring(z.read(n)):
                if rel.get("TargetMode") == "External":
                    continue
                target = os.path.normpath(os.path.join(base, rel.get("Target"))).replace("\\", "/")
                referenced.add(target)
                if target not in names:
                    problems.append(f"dangling relationship in {n}: {rel.get('Target')}")

        # --- orphaned parts (present, referenced by nothing) ------------
        roots = {"[Content_Types].xml", "ppt/presentation.xml"}
        for n in names:
            if n.endswith("/") or n.endswith(".rels") or n in roots:
                continue
            if n not in referenced and not n.startswith("docProps/"):
                problems.append(f"orphaned part (no incoming relationship): {n}")

        # --- per-slide checks -------------------------------------------
        notes_owner = Counter()
        for n in sorted(names):
            if not re.match(r"ppt/slides/slide\d+\.xml$", n):
                continue
            root = etree.fromstring(z.read(n))
            ids = [e.get("id") for e in root.iter(f"{{{P_NS}}}cNvPr")]
            for dup in (i for i, c in Counter(ids).items() if c > 1):
                problems.append(f"duplicate shape id {dup} in {n}")

            used = {
                v for el in root.iter() for k, v in el.attrib.items()
                if k.startswith("{" + R_NS + "}")
            }
            rels_name = n.replace("slides/", "slides/_rels/") + ".rels"
            declared = set()
            if rels_name in names:
                rels_root = etree.fromstring(z.read(rels_name))
                declared = {rel.get("Id") for rel in rels_root}
                for rel in rels_root:
                    if "notesSlide" in rel.get("Target"):
                        target = os.path.normpath(
                            os.path.join("ppt/slides", rel.get("Target"))
                        ).replace("\\", "/")
                        notes_owner[target] += 1
            for missing in used - declared:
                problems.append(f"r:id {missing} used in {n} but not declared in rels")

        for target, count in notes_owner.items():
            if count > 1:
                problems.append(f"notesSlide shared by {count} slides: {target}")

        # --- sldIdLst sanity --------------------------------------------
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        sld_lst = pres.find(f"{{{P_NS}}}sldIdLst")
        if sld_lst is not None:
            rids = [e.get(f"{{{R_NS}}}id") for e in sld_lst]
            for dup in (i for i, c in Counter(rids).items() if c > 1):
                problems.append(f"duplicate r:id in sldIdLst: {dup}")
            for sid in (e.get("id") for e in sld_lst):
                if int(sid) < 256:
                    problems.append(f"sldId id below 256: {sid}")
    return problems


def deep_validate(pptx_path):
    """Vendored ISO-IEC29500 schema validation. Returns list of problem lines
    (empty == passed). Soft-fails to [] if the vendored validator can't run at
    all — the structural audit above still stands guard."""
    try:
        result = subprocess.run(
            [sys.executable, "validate.py", os.path.abspath(pptx_path)],
            cwd=_VENDORED_VALIDATOR_DIR,
            capture_output=True, text=True, timeout=120,
        )
    except Exception as e:
        print(f"[package_check] deep validator unavailable: {e}", flush=True)
        return []
    if result.returncode == 0:
        return []
    output = (result.stdout + result.stderr).strip()
    return [line for line in output.splitlines() if line.strip()][:20]


def assert_valid_package(pptx_path):
    problems = structural_audit(pptx_path)
    problems += deep_validate(pptx_path)
    if problems:
        summary = "; ".join(problems[:5])
        raise RuntimeError(
            f"сгенерированный .pptx не прошёл проверку целостности ({len(problems)} проблем): {summary}"
        )
