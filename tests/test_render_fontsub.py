"""Preview font substitution (rendering/render.py) without launching
LibreOffice: a deck referencing a family this machine doesn't have must get a
render copy with that family rewritten to the fallback; the original file must
never be modified."""

import zipfile

from conftest import SURVEY_31, requires

from rendering.render import (
    fallback_typeface,
    _installed_families,
    _prepare_render_copy,
    _referenced_typefaces,
)


@requires(SURVEY_31)
def test_missing_families_rewritten_in_copy(tmp_path):
    faces = _referenced_typefaces(SURVEY_31)
    assert "Aeroport" in faces  # the brand font this corpus is known to use

    copy_path = _prepare_render_copy(SURVEY_31, str(tmp_path))
    if "aeroport" in _installed_families():
        # Machine actually has the font — nothing to substitute.
        assert copy_path == SURVEY_31
        return

    assert copy_path != SURVEY_31
    with zipfile.ZipFile(copy_path) as z:
        blob = b"".join(z.read(n) for n in z.namelist() if n.endswith(".xml"))
    assert b'typeface="Aeroport"' not in blob
    assert b'typeface="%s"' % fallback_typeface().encode() in blob

    # Source untouched.
    with zipfile.ZipFile(SURVEY_31) as z:
        blob = b"".join(z.read(n) for n in z.namelist() if n.endswith(".xml"))
    assert b'typeface="Aeroport"' in blob
