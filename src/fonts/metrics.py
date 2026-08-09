"""Real font metrics for text fitting (docs/IMPROVEMENT_PLAN.md, пункт 1).

PPTX files can embed their fonts (ppt/fonts/*.fntdata — EOT-wrapped or raw
TrueType; LibreOffice's "EOT out of spec" render warnings on our test template
were the tell that they're there). Parsing those with fontTools gives exact
per-glyph advance widths, which replaces the old "average char ≈ 0.62 × size"
guess in text_fit.py — the root cause of bullet-icon drift and over-padded
title boxes.

Lookup chain per font name:
  1. a font embedded in the pptx itself whose name-table family matches;
  2. a metrically-reasonable system donor (Arial/Helvetica class);
  3. None — callers fall back to the char-count estimate.
"""

import io
import re
import zipfile
from functools import lru_cache

from fontTools.ttLib import TTFont

# sfnt signatures a TrueType/OpenType payload can start with. EOT files carry
# the raw font at some offset after a variable-length header — scanning for the
# signature and letting fontTools validate is far more robust than parsing the
# EOT header (which real-world files get wrong — see LibreOffice's complaints).
_SFNT_SIGNATURES = (b"\x00\x01\x00\x00", b"OTTO", b"true")

_REQUIRED_TABLES = ("cmap", "hmtx", "head", "hhea")

# Metric donors for when the pptx has no usable embedded font. Ordered; first
# parseable wins. These are Helvetica/Arial-class faces — not exact for every
# brand font, but far closer than a flat per-character constant.
# The BOLD donor matters as much as the regular one: our titles are drawn bold
# (synthesizer._style_paragraph passes bold=True) and Cyrillic bold runs 7-8%
# wider than the regular face — measured on Arial at 40pt: «Что мешало собирать
# отчётность вовремя» is 796pt regular and 858pt bold, against a title box
# budget of 853pt. The regular measurement promised one line and the render drew
# two, which is invisible while a layout keeps slack under the title and lands
# the body on top of it as soon as one does not.
_SYSTEM_DONOR_BOLD_CANDIDATES = (
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "/Library/Fonts/Arial Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
)

_SYSTEM_DONOR_CANDIDATES = (
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Helvetica.ttc",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
)

_LATIN_SAMPLE = (
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,!?-"
)

CYRILLIC_RANGE = range(0x0410, 0x0450)


class FontMetrics:
    """Width/height math over a parsed sfnt font. All width results are in
    points for the given point size (advance widths scaled by unitsPerEm)."""

    def __init__(self, ttfont, family_name=None):
        self._font = ttfont
        self.family_name = family_name or _family_name(ttfont)
        self._upem = ttfont["head"].unitsPerEm or 1000
        self._cmap = ttfont.getBestCmap() or {}
        self._hmtx = ttfont["hmtx"]

        hhea = ttfont["hhea"]
        raw_line = (hhea.ascent - hhea.descent + hhea.lineGap) / self._upem
        # Clamp: some display fonts declare huge line gaps that PowerPoint's
        # single spacing doesn't actually honor, and a too-tall estimate makes
        # us shrink text that would have fit.
        self.line_height_factor = min(1.4, max(1.15, raw_line))

        sample_units = [
            self._hmtx[self._cmap[ord(ch)]][0]
            for ch in _LATIN_SAMPLE
            if ord(ch) in self._cmap
        ]
        if not sample_units:
            sample_units = [m[0] for m in list(self._hmtx.metrics.values())[:200]]
        self._avg_units = sum(sample_units) / len(sample_units) if sample_units else self._upem * 0.5

    def text_width_pt(self, text, size_pt):
        total_units = 0
        for ch in text:
            glyph = self._cmap.get(ord(ch))
            if glyph is not None:
                total_units += self._hmtx[glyph][0]
            else:
                # Char the font can't render (e.g. Cyrillic in a Latin-only brand
                # font): the renderer substitutes another face; its metrics are
                # unknowable here, so this font's own average is the best guess.
                total_units += self._avg_units
        return total_units / self._upem * size_pt

    def line_height_pt(self, size_pt):
        return self.line_height_factor * size_pt

    def coverage(self, text):
        """Fraction of characters this font can actually render."""
        relevant = [ch for ch in text if not ch.isspace()]
        if not relevant:
            return 1.0
        hit = sum(1 for ch in relevant if ord(ch) in self._cmap)
        return hit / len(relevant)

    def supports_cyrillic(self):
        return any(cp in self._cmap for cp in CYRILLIC_RANGE)


def _family_name(ttfont):
    try:
        name = ttfont["name"]
        return name.getDebugName(16) or name.getDebugName(1)
    except Exception:
        return None


def _try_parse_sfnt(blob):
    """blob may be a raw sfnt or an EOT wrapper; find the payload and parse."""
    offsets = [0]
    for sig in _SFNT_SIGNATURES:
        start = 0
        while True:
            pos = blob.find(sig, start)
            if pos == -1:
                break
            offsets.append(pos)
            start = pos + 1
    for offset in offsets:
        try:
            font = TTFont(io.BytesIO(blob[offset:]), lazy=True)
            for table in _REQUIRED_TABLES:
                _ = font[table]
            return font
        except Exception:
            continue
    return None


def _normalize(name):
    return re.sub(r"[^a-z0-9а-яё]", "", (name or "").lower())


@lru_cache(maxsize=16)
def _embedded_fonts(pptx_path):
    """[(family_name, FontMetrics)] for every parseable embedded font."""
    found = []
    try:
        with zipfile.ZipFile(pptx_path) as z:
            for entry in z.namelist():
                if not entry.startswith("ppt/fonts/"):
                    continue
                font = _try_parse_sfnt(z.read(entry))
                if font is None:
                    continue
                try:
                    metrics = FontMetrics(font)
                except Exception:
                    continue
                found.append((metrics.family_name, metrics))
    except Exception:
        pass
    return found


@lru_cache(maxsize=1)
def _system_donor(bold=False):
    candidates = (_SYSTEM_DONOR_BOLD_CANDIDATES + _SYSTEM_DONOR_CANDIDATES
                  if bold else _SYSTEM_DONOR_CANDIDATES)
    for path in candidates:
        for font_number in (0, None):
            try:
                kwargs = {"lazy": True}
                if font_number is not None:
                    kwargs["fontNumber"] = font_number
                font = TTFont(path, **kwargs)
                for table in _REQUIRED_TABLES:
                    _ = font[table]
                return FontMetrics(font)
            except Exception:
                continue
    return None


def get_font_metrics(pptx_path, font_name, bold=False):
    """FontMetrics for font_name via embedded → system donor → None.

    `bold` picks the bold face where one is available. None of the five real
    templates embeds a font, so in practice this is the difference between the
    regular and the bold donor — and that difference is 7-8% of the width on
    Cyrillic, enough to turn a predicted one-line title into a drawn two-line
    one."""
    wanted = _normalize(font_name)
    embedded = _embedded_fonts(pptx_path)
    if wanted:
        matches = [(family, metrics) for family, metrics in embedded
                   if _normalize(family)
                   and (_normalize(family) == wanted
                        or _normalize(family).startswith(wanted)
                        or wanted.startswith(_normalize(family)))]
        if matches:
            if bold:
                for family, metrics in matches:
                    if "bold" in (family or "").lower():
                        return metrics
            return matches[0][1]
    # A deck usually embeds exactly its brand family — if there's precisely one
    # embedded family, it's a better metric source for any requested name than
    # a generic donor.
    families = {f for f, _ in embedded if f}
    if len(families) == 1 and embedded:
        return embedded[0][1]
    return _system_donor(bold=bold)


class FontResolver:
    """Per-template resolver: maps run-level font names (including theme tokens
    like "+mn-lt"/"+mj-lt" that python-pptx surfaces verbatim) to FontMetrics,
    with caching. One instance per generate() call."""

    def __init__(self, pptx_path, theme=None):
        self.pptx_path = pptx_path
        self.theme_fonts = (theme or {}).get("fonts") or {}
        self._cache = {}

    def resolve_name(self, font_name):
        if not font_name or font_name.startswith("+mn"):
            return self.theme_fonts.get("minorFont") or font_name
        if font_name.startswith("+mj"):
            return self.theme_fonts.get("majorFont") or font_name
        return font_name

    def metrics_for(self, font_name, bold=False):
        name = self.resolve_name(font_name)
        key = (name or "__default__", bool(bold))
        if key not in self._cache:
            self._cache[key] = get_font_metrics(self.pptx_path, name, bold=bold)
        return self._cache[key]
