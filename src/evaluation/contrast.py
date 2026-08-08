"""Deterministic text/background contrast from the RENDER — the objective
backstop for the one defect that the eye catches instantly but neither the
XML-geometry checks nor the GigaChat judge do: the deck's light ink (orange
accents, light-grey body) landing on a white breather slide, readable on the
navy slides and washed-out on the white ones.

XML can't see it — the text colour is theme-inherited, and the failure only
exists once LibreOffice has painted the actual background under it. So we measure
it on the pixels, with WCAG contrast.

Robustness note (learned the hard way on real renders): a coarse thumbnail
smears text into anti-alias blends, and taking the worst of ALL colour clusters
just measures those dark/light halos, flagging every slide. The fix: sample at a
higher resolution and look only at the DOMINANT ink cluster (>=15% of the ink
pixels) — that's the crisp glyph interior, i.e. the real text colour, never a
halo (halos spread thin across many shades). Its contrast vs the border-modal
background cleanly separates navy slides (~6:1) from the white ones (~2.4:1).
Pure PIL (no numpy in this env).
"""

from collections import Counter

from PIL import Image

_QUANT_SHIFT = 4
_RESIZE = (480, 270)       # enough that glyph interiors survive as real colour, still fast
INK_DISTANCE = 70          # RGB distance from bg to count as ink, not an anti-alias halo
MIN_INK_FRACTION = 0.15    # a real text colour is a big share of ink pixels; halos never are
LOW_CONTRAST_RATIO = 3.0   # WCAG AA large-text; the deck's ink drops under this on a white bg


def _quantize(rgb):
    r, g, b = rgb
    return (r >> _QUANT_SHIFT << _QUANT_SHIFT,
            g >> _QUANT_SHIFT << _QUANT_SHIFT,
            b >> _QUANT_SHIFT << _QUANT_SHIFT)


def _lin(c):
    c /= 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def _rel_luminance(rgb):
    r, g, b = (_lin(x) for x in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(a, b):
    la, lb = _rel_luminance(a), _rel_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _dist(a, b):
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def _pixels(png_path):
    with Image.open(png_path) as img:
        im = img.convert("RGB").resize(_RESIZE)
        raw = im.tobytes()
    w, h = _RESIZE
    px = [(raw[i], raw[i + 1], raw[i + 2]) for i in range(0, len(raw), 3)]
    return px, w, h


def _background(px, w, h):
    """Modal colour of the frame border — the border is always background, even
    on card layouts where the interior is a coloured panel."""
    border = []
    for y in (0, 1, h - 2, h - 1):
        border += [px[y * w + x] for x in range(0, w, 3)]
    for x in (0, 1, w - 2, w - 1):
        border += [px[y * w + x] for y in range(0, h, 3)]
    return Counter(_quantize(p) for p in border).most_common(1)[0][0]


def slide_text_contrast(png_path):
    """WCAG contrast of the slide's DOMINANT ink colour vs its background.
    None when the slide has no substantial ink block (blank / breather / a
    sparse divider) — a safe 'skip', never a false low-contrast flag."""
    px, w, h = _pixels(png_path)
    bg = _background(px, w, h)
    ink = [_quantize(p) for p in px if _dist(p, bg) > INK_DISTANCE]
    if not ink:
        return None
    counts = Counter(ink)
    total = len(ink)
    worst = None
    for color, n in counts.most_common(6):
        if n / total < MIN_INK_FRACTION:
            break  # most_common is descending — nothing further qualifies
        cr = contrast_ratio(color, bg)
        worst = cr if worst is None else min(worst, cr)
    return worst


def evaluate_contrast(png_paths, threshold=LOW_CONTRAST_RATIO):
    """Returns {"per_slide": [ratio|None], "low_contrast_slides": [0-based idx]}."""
    per = [slide_text_contrast(p) for p in png_paths]
    low = [i for i, c in enumerate(per) if c is not None and c < threshold]
    return {"per_slide": per, "low_contrast_slides": low}


# --- Declared-colour check -------------------------------------------------
# The pixel scan above can only see text that differs from the background by at
# least INK_DISTANCE. Text painted almost exactly in the background colour forms
# no distinct cluster at all, so the scan reports "no ink" and the slide passes
# silently — the worst case slips through. Measured on a real deck: subtitle
# #1D1C1D on a (16,16,16) slide is 22 RGB units from the background, invisible
# to the eye AND invisible to the scan.
#
# So read the colours the file DECLARES and compare them with the MEASURED
# background. Only explicit run colours are used; theme/inherited ones are left
# alone rather than guessed at.
#
# Two false-positive guards, both learned from real decks:
#
# 1. A designer may legitimately put light text on a dark button sitting on a
#    light slide, which looks identical to this test from the XML. So a slide is
#    only flagged when the near-invisible runs carry at least
#    INVISIBLE_TEXT_SHARE of its characters — a whole unreadable slide trips it,
#    a small inverted label does not.
# 2. WCAG contrast is a LUMINANCE ratio and ignores hue, so a high-chroma brand
#    pairing scores as badly as invisible text: T-Zh's red-on-blue slide measures
#    1.08, yet the text is perfectly legible because the hues are 339 RGB units
#    apart. "Invisible" must mean the colours are actually CLOSE, so a flag needs
#    low contrast AND small colour distance. The real failures sit far under this
#    (#3D3D3D on (16,16,16) is 78 apart, #1D1C1D on (16,16,16) just 22).
INVISIBLE_TEXT_SHARE = 0.5
INVISIBLE_MAX_DISTANCE = 120


def _declared_run_colours(slide):
    """[(text_length, (r,g,b))] for runs with an EXPLICIT rgb colour."""
    out = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                text = run.text or ""
                if not text.strip():
                    continue
                colour = run.font.color
                try:
                    # .rgb raises for scheme colours — those are inherited, skip.
                    if colour is None or colour.type is None or colour.rgb is None:
                        continue
                    rgb = colour.rgb
                except Exception:  # noqa: BLE001 — scheme colour, see note below
                    # Scheme colours are deliberately NOT resolved through the
                    # theme palette. Tried it and reverted: the theme claimed
                    # #EEEEEE (lt2) for text that renders BLACK, so two legible
                    # card slides came back "invisible". Scheme tokens resolve
                    # through the master's <p:clrMap>, and on top of that the
                    # theme routinely lies about the deck's real look (CLAUDE.md).
                    # Staying blind here is safer than being confidently wrong —
                    # the pixel scan still covers those slides.
                    continue
                out.append((len(text.strip()), (rgb[0], rgb[1], rgb[2])))
    return out


def evaluate_declared_contrast(pptx_path, png_paths, threshold=LOW_CONTRAST_RATIO):
    """Slides whose declared text colour is invisible against the measured
    background. Returns {"per_slide": [worst_ratio|None], "invisible_slides":
    [idx], "coverage": (slides_with_data, total)}.

    `coverage` is reported because this pass reads EXPLICIT rgb only and is
    simply blind on theme-coloured text — on a card template that can be most of
    the deck. Without it in the result, every later reader has to rediscover
    that blindness before trusting a clean verdict."""
    from pptx import Presentation

    from design_system.style_profile import page_base_color, slide_background

    prs = Presentation(pptx_path)
    slides = list(prs.slides)
    page_base = page_base_color(png_paths) if png_paths else None

    per, flagged = [], []
    for i, slide in enumerate(slides):
        bg = slide_background(png_paths[i], page_base) if i < len(png_paths) else None
        runs = _declared_run_colours(slide)
        if bg is None or not runs:
            per.append(None)
            continue
        total = sum(n for n, _ in runs) or 1
        bad_chars, worst = 0, None
        for n, rgb in runs:
            ratio = contrast_ratio(rgb, bg)
            worst = ratio if worst is None else min(worst, ratio)
            if ratio < threshold and _dist(rgb, bg) < INVISIBLE_MAX_DISTANCE:
                bad_chars += n
        per.append(worst)
        if bad_chars / total >= INVISIBLE_TEXT_SHARE:
            flagged.append(i)
    return {
        "per_slide": per,
        "invisible_slides": flagged,
        "coverage": (sum(1 for r in per if r is not None), len(per)),
    }


def reconcile(pixel_result, declared_result):
    """Combine the two passes into the slides that really are unreadable.

    The pixel scan is unreliable on card layouts: the page margin around a
    coloured card reads as a second huge colour region, so it called six of
    eight T-Zh slides low-contrast when every one was perfectly legible. The
    declared-colour pass measures the actual run colours, so where it HAS data
    for a slide and finds nothing invisible, its verdict wins and the pixel flag
    is dropped. Where the text is theme-inherited it has no colours to read,
    stays silent, and the pixel scan remains the only witness.

    Returns (low_slides, suppressed) — both 0-based, sorted."""
    invisible = set(declared_result.get("invisible_slides", []))
    declared_clear = {
        i for i, ratio in enumerate(declared_result.get("per_slide", []))
        if ratio is not None and i not in invisible
    }
    pixel_low = set(pixel_result.get("low_contrast_slides", []))
    suppressed = pixel_low & declared_clear
    return sorted((pixel_low - suppressed) | invisible), sorted(suppressed)
