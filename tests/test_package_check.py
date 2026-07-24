"""Package-integrity guard (plan 9.6). The two historically corrupted files'
defect classes must be caught, clean generations must pass, and generate()
must refuse to return a package that fails the audit (the gate is wired at
save time, so every other generate-based test exercises it implicitly)."""

import os
import zipfile

from conftest import SURVEY_31, TEMPLATES_DIR, requires

from qa.package_check import structural_audit

TJ = os.path.join(TEMPLATES_DIR, "custom_f496182bb15f42bb.pptx")


@requires(SURVEY_31)
def test_pristine_templates_pass():
    assert structural_audit(SURVEY_31) == []


@requires(TJ)
def test_tj_template_passes():
    assert structural_audit(TJ) == []


@requires(TJ)
def test_shared_notes_detected(tmp_path):
    """Synthesize the exact corruption class of the 16c83afc incident: two
    slides declaring the same notesSlide part."""
    broken = str(tmp_path / "broken.pptx")
    with zipfile.ZipFile(TJ) as zin, zipfile.ZipFile(broken, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == "ppt/slides/_rels/slide2.xml.rels":
                # point slide2's notes at slide1's notesSlide part
                data = data.replace(b"notesSlide2.xml", b"notesSlide1.xml")
            # write by NAME, not ZipInfo: copied ZipInfo keeps flag_bits (data
            # descriptor) that writestr doesn't emit -> unreadable archive
            zout.writestr(name, data)
    problems = structural_audit(broken)
    assert any("notesSlide shared" in p for p in problems), problems


@requires(TJ)
def test_orphaned_part_detected(tmp_path):
    """The other incident class: a slide part present in the zip but absent
    from sldIdLst/relationships (simulated by injecting an unreferenced
    part)."""
    broken = str(tmp_path / "orphan.pptx")
    with zipfile.ZipFile(TJ) as zin, zipfile.ZipFile(broken, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            zout.writestr(name, zin.read(name))
        zout.writestr("ppt/slides/slide99.xml", zin.read("ppt/slides/slide1.xml"))
    problems = structural_audit(broken)
    assert any("orphaned part" in p and "slide99" in p for p in problems), problems
