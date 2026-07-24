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
