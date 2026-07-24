"""Regression for a real user report: PowerPoint always offered to "repair"
decks generated from a template whose roles got reused (the T-Ж template,
where nearly every role needed cloning). Two independent defects, found one
after the other because the first fix alone didn't make the report go away:

1. _reorder_and_prune_slides removed a pruned slide's <p:sldId> from sldIdLst
   but never dropped its presentation.xml.rels relationship, leaving the
   underlying slide PART still physically in the .pptx — reachable via
   relationship but absent from the actual slide show.
2. _clone_slide copied EVERY relationship from the source slide, including
   notesSlide — a part meant to belong to exactly one slide (no back-edge
   exists the other way). Cloning a slide N times made N+1 different slides
   all declare the SAME notesSlide part as their notes.

Both are valid zip, valid XML, and still flagged corrupt by PowerPoint's
stricter OPC/ownership consistency checks. Verifies the actual saved package,
not python-pptx's in-memory view — both bugs only exist in what's written for
relationships prs.slides never surfaces directly."""

import re
import zipfile
from collections import Counter

from conftest import SURVEY_31, requires

from lxml import etree

from generator.generator import generate

PRES_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _package_slide_audit(pptx_path):
    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        visible_rids = {el.get("{%s}id" % PRES_NS["r"]) for el in pres.find("p:sldIdLst", PRES_NS)}
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
        slide_rels = {rel.get("Id"): rel.get("Target") for rel in rels if "slides/slide" in rel.get("Target")}
        physical_parts = {n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)}

        notes_owners = Counter()
        for slide_part in physical_parts:
            rels_name = slide_part.replace("slides/", "slides/_rels/") + ".rels"
            if rels_name not in names:
                continue
            for rel in etree.fromstring(z.read(rels_name)):
                if "notesSlide" in rel.get("Target"):
                    notes_owners[rel.get("Target")] += 1
    return visible_rids, slide_rels, physical_parts, notes_owners


@requires(SURVEY_31)
def test_clone_source_leaves_no_orphaned_parts(tmp_path):
    # Same role assigned 3 times forces the clone path (see matcher/generator
    # _clone_slide) — this is exactly the shape of plan that corrupted 16c83afc.
    # Template slide 28 carries real speaker notes — required to exercise the
    # notesSlide-sharing defect; a source slide without notes can't catch it.
    plan = [
        ({"type": "bullet_list", "title": "Первый", "bullets": ["a", "b", "c"]}, 28),
        ({"type": "bullet_list", "title": "Второй", "bullets": ["d", "e"]}, 28),
        ({"type": "bullet_list", "title": "Третий", "bullets": ["f", "g", "h", "i"]}, 28),
    ]
    out = str(tmp_path / "orphan_check.pptx")
    generate(SURVEY_31, plan, out)

    visible_rids, slide_rels, physical_parts, notes_owners = _package_slide_audit(out)

    assert len(visible_rids) == len(plan)
    # Every relationship declared in presentation.xml.rels must be visible —
    # zero dangling "part exists, but not in the slide show" relationships.
    orphaned_rels = set(slide_rels) - visible_rids
    assert orphaned_rels == set(), f"orphaned slide relationships: {orphaned_rels}"
    # And no leftover physical slide parts beyond what's actually shown.
    assert len(physical_parts) == len(plan)
    # A notesSlide part must belong to exactly one slide — never shared.
    shared = {target: n for target, n in notes_owners.items() if n > 1}
    assert shared == {}, f"notesSlide parts shared by multiple slides: {shared}"


@requires(SURVEY_31)
def test_no_clone_path_still_clean(tmp_path):
    """A plan that never triggers cloning must also produce a clean package —
    guards against the fix accidentally dropping slides that WERE meant to
    ship."""
    plan = [
        ({"type": "title", "title": "Т", "subtitle": "П"}, 23),
        ({"type": "bullet_list", "title": "Единственный", "bullets": ["a", "b", "c"]}, 28),
    ]
    out = str(tmp_path / "no_clone.pptx")
    generate(SURVEY_31, plan, out)
    visible_rids, slide_rels, physical_parts, notes_owners = _package_slide_audit(out)
    assert len(visible_rids) == 2 == len(physical_parts)
    assert set(slide_rels) == visible_rids
    assert all(n == 1 for n in notes_owners.values())
