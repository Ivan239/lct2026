"""What the deck actually looks like, as opposed to what its theme claims.

Synthesized slides used to take colors/fonts straight from the OOXML theme —
but decks built by hand (and our own synthetic presets) often never touch the
theme: their real style lives in explicit run colors and per-slide solid
backgrounds, while theme1.xml still holds the Office defaults. A divider
synthesized from such a theme lands as a white Calibri slide inside a navy
Georgia deck. This module takes a majority vote over the template's actual
slides and returns only the style facts it is confident about; the caller
overlays them onto the parsed theme, so decks that DO carry an honest theme
(real corporate exports, scheme-color runs) keep using it unchanged."""

from collections import Counter

from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.oxml.ns import qn

# A run at or above this size (or bold) votes for the accent/heading style;
# everything else votes for body style.
_HEADING_PT = 24

_MIN_CHAR_VOTES = 20     # ignore run-style verdicts built on a handful of characters
_MIN_SHARE = 0.5         # a style must cover half its sample to be "the deck's style"
_BG_MIN_SHARE = 0.6
_MIN_CONTRAST = 80       # |luminance delta| below this is unreadable, veto the pick


def _slide_bg_hex(slide):
    """Explicit solid slide background, read from raw XML — going through
    slide.background.fill would insert <p:bg> elements into slides that had
    none (python-pptx get-or-add behavior), mutating the deck on a read."""
    bg = slide._element.cSld.find(qn("p:bg"))
    if bg is None:
        return None
    srgb = bg.find(".//" + qn("a:srgbClr"))
    return f"#{srgb.get('val')}" if srgb is not None and srgb.get("val") else None


def _run_rgb_hex(run):
    try:
        color = run.font.color
        if color is None or color.type != MSO_COLOR_TYPE.RGB:
            return None
        return f"#{color.rgb}"
    except (AttributeError, TypeError):
        return None


def _luminance(hex_str):
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _confident_top(counter, min_votes, min_share):
    total = sum(counter.values())
    if total < min_votes:
        return None
    value, votes = counter.most_common(1)[0]
    return value if votes / total >= min_share else None


def observe_deck_style(prs):
    """Returns a dict with any of: bg, accent, text, major_font, minor_font —
    only the keys the deck's own slides give a confident majority for."""
    bg_votes = Counter()
    accent_color, body_color = Counter(), Counter()
    heading_font, body_font = Counter(), Counter()

    for slide in prs.slides:
        bg = _slide_bg_hex(slide)
        if bg:
            bg_votes[bg] += 1
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    chars = len(run.text.strip())
                    if not chars:
                        continue
                    size_pt = run.font.size.pt if run.font.size else None
                    is_heading = bool(run.font.bold) or (size_pt or 0) >= _HEADING_PT
                    rgb = _run_rgb_hex(run)
                    if rgb:
                        (accent_color if is_heading else body_color)[rgb] += chars
                    # "+mj-lt"/"+mn-lt" are theme tokens, not font names — a run
                    # carrying one already follows the theme, which is the
                    # fallback anyway; written literally it renders as a missing
                    # font (seen on custom_30e96c06e2d47ec3).
                    if run.font.name and not run.font.name.startswith("+"):
                        (heading_font if is_heading else body_font)[run.font.name] += chars

    observed = {}
    if len(prs.slides) and sum(bg_votes.values()) >= max(1, len(prs.slides) // 2):
        bg = _confident_top(bg_votes, 1, _BG_MIN_SHARE)
        if bg:
            observed["bg"] = bg
    accent = _confident_top(accent_color, _MIN_CHAR_VOTES, _MIN_SHARE)
    if accent:
        observed["accent"] = accent
    text = _confident_top(body_color, _MIN_CHAR_VOTES, _MIN_SHARE)
    if text:
        observed["text"] = text
    major = _confident_top(heading_font, _MIN_CHAR_VOTES, _MIN_SHARE)
    if major:
        observed["major_font"] = major
    minor = _confident_top(body_font, _MIN_CHAR_VOTES, _MIN_SHARE)
    if minor:
        observed["minor_font"] = minor
    return observed


def apply_observed_style(theme, observed):
    """Overlay observed style facts onto a parsed theme (deep-ish copy; the
    original stays untouched — fillers keep using it). Readability veto: an
    observed body/accent color is only taken against the EFFECTIVE background
    if the pair stays readable, and if the final text color clashes with the
    final background it is replaced with plain white/black outright."""
    merged = dict(theme)
    merged["fonts"] = dict(theme.get("fonts") or {})
    merged["palette"] = dict(theme.get("palette") or {})

    if "major_font" in observed:
        merged["fonts"]["majorFont"] = observed["major_font"]
    if "minor_font" in observed:
        merged["fonts"]["minorFont"] = observed["minor_font"]
    if "bg" in observed:
        merged["palette"]["lt1"] = observed["bg"]
    if "accent" in observed:
        merged["palette"]["accent1"] = observed["accent"]
    if "text" in observed:
        merged["palette"]["dk1"] = observed["text"]

    bg = merged["palette"].get("lt1")
    text = merged["palette"].get("dk1")
    if bg and text and abs(_luminance(bg) - _luminance(text)) < _MIN_CONTRAST:
        merged["palette"]["dk1"] = "#FFFFFF" if _luminance(bg) < 128 else "#1A1A1A"
    return merged
