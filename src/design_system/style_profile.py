"""Template Style Profile — the measured half (plan items 9.0/9.1).

No template taxonomy exists or can exist; instead of recognizing "a type",
every template is made to describe its own design logic at upload time, and
downstream stages (matcher, synthesis, prompts) follow that description
blindly. This module builds the deterministic, LLM-free part of the profile
from the renders we already produce:

- background palette: dominant edge-band color per slide (edges are almost
  always background, whatever mechanism draws it — solid fill, theme, or a
  full-bleed decor PICTURE that XML-level inspection can't see);
- rotation: whether the deck cycles those backgrounds (T-Ж folders alternate
  yellow/green/blue/gray; a single-background deck degrades to "no rotation"
  with zero special-casing);
- bookends: the title/closing slides' colors (T-Ж opens and closes on brand
  yellow);
- breathers: full-bleed photo slides that intentionally sit outside the
  rhythm.

Verified empirically across four unrelated templates: the same code recovers
the T-Ж 4-cycle (confidence 1.0) and reports no-rotation on single-background
decks. NEVER cached to disk: rebuilding costs ~milliseconds per slide and a
stale cache has already masked fixes twice in this project's history.
"""

from collections import Counter

from PIL import Image

# 5 bits/channel: coarse enough to absorb anti-aliasing noise inside one
# folder color, fine enough to keep yellow and beige apart (measured).
_QUANT_SHIFT = 3

# Measured on the T-Ж corpus: a colored card that a human calls "the
# background" holds >=0.27 of the frame even when photos eat half the slide;
# no single quantized photo color reaches 0.22. The page base (paper-white
# margins) needs the higher bar because it must dominate outright to count.
_CARD_MIN_SHARE = 0.22
_BASE_MIN_SHARE = 0.35

# Rotation is asserted only when it's clearly the deck's own rule, not an
# accident of two slides differing.
_ROTATION_MIN_CONFIDENCE = 0.7
_ROTATION_MIN_PALETTE = 2


def _quantize(rgb):
    r, g, b = rgb
    return (r >> _QUANT_SHIFT << _QUANT_SHIFT,
            g >> _QUANT_SHIFT << _QUANT_SHIFT,
            b >> _QUANT_SHIFT << _QUANT_SHIFT)


def _frame_colors(png_path):
    with Image.open(png_path) as img:
        thumb = img.convert("RGB").resize((64, 36))
        raw = thumb.tobytes()
    pixels = [(raw[i], raw[i + 1], raw[i + 2]) for i in range(0, len(raw), 3)]
    counts = Counter(_quantize(p) for p in pixels)
    total = len(pixels)
    return [(c, n / total) for c, n in counts.most_common(6)]


def _edge_color(png_path):
    with Image.open(png_path) as img:
        im = img.convert("RGB")
        w, h = im.size
        samples = [im.getpixel((x, 1)) for x in range(0, w, max(1, w // 80))]
        samples += [im.getpixel((x, h - 2)) for x in range(0, w, max(1, w // 80))]
    return Counter(_quantize(p) for p in samples).most_common(1)[0][0]


def page_base_color(rendered_png_paths):
    """The deck's paper color — modal edge color across all slides. Card-style
    designs (T-Ж folders) keep this visible as margins on every slide; it is
    the one color that must NOT be mistaken for a slide's own background."""
    votes = Counter(_edge_color(p) for p in rendered_png_paths)
    return votes.most_common(1)[0][0]


def slide_background(png_path, page_base=None):
    """The color a human would name as this slide's background:
    the most prominent color that is NOT the deck's page base (>= 22% of the
    frame — a colored card wins over larger white margins), else the page
    base itself when it genuinely dominates, else None (full-bleed photo,
    a rhythm "breather"). Measured thresholds — see constants above."""
    ranked = _frame_colors(png_path)
    for color, share in ranked:
        if share < _CARD_MIN_SHARE:
            break
        if page_base is None or color != page_base:
            return color
    if ranked and page_base is not None:
        base_share = dict(ranked).get(page_base, 0.0)
        if base_share >= _BASE_MIN_SHARE:
            return page_base
    if page_base is None and ranked and ranked[0][1] >= _BASE_MIN_SHARE:
        return ranked[0][0]
    return None


def _infer_rotation(sequence):
    """sequence: per-slide background colors (None = breather, skipped).
    Returns (order, confidence). Confidence = share of adjacent pairs that
    CHANGE color; order = first-seen cycle of the RECURRING palette when
    confident. Colors seen exactly once (the T-Ж white "technical slide") are
    outliers, not part of the deck's rotation language — excluded from the
    cycle so generated decks never land on them."""
    colored = [c for c in sequence if c is not None]
    counts = Counter(colored)
    recurring = [c for c in dict.fromkeys(colored) if counts[c] >= 2]
    if len(recurring) < _ROTATION_MIN_PALETTE or len(colored) < 3:
        return [], 0.0
    pairs = list(zip(colored, colored[1:]))
    changes = sum(1 for a, b in pairs if a != b)
    confidence = changes / len(pairs)
    if confidence < _ROTATION_MIN_CONFIDENCE:
        return [], confidence
    return recurring, confidence


def build_measured_profile(rendered_png_paths, archetype_map=None):
    """rendered_png_paths: per-slide renders, index-aligned with the deck.
    archetype_map: {slide_idx: role} — used for bookends when available.
    Returns the measured profile dict (see module doc)."""
    page_base = page_base_color(rendered_png_paths) if rendered_png_paths else None
    backgrounds = [slide_background(p, page_base) for p in rendered_png_paths]

    order, confidence = _infer_rotation(backgrounds)
    breathers = [i for i, c in enumerate(backgrounds) if c is None]

    # Bookends by POSITION, constrained to rotation colors: the first slide of
    # a deck is its cover and the last colored one its closing, whatever the
    # archetype map thinks (roles have been observed mislabeled on merged
    # clusters, while position never lies about which slide opens the deck).
    # Outlier colors (the white technical slide) are not in `order`, so the
    # closing anchor naturally skips past them.
    bookends = {}
    anchor_pool = order if order else [c for c in dict.fromkeys(backgrounds) if c is not None]
    for i in range(len(backgrounds)):
        if backgrounds[i] in anchor_pool:
            bookends["title"] = backgrounds[i]
            break
    for i in range(len(backgrounds) - 1, -1, -1):
        if backgrounds[i] in anchor_pool:
            bookends["closing"] = backgrounds[i]
            break

    return {
        "backgrounds": {i: c for i, c in enumerate(backgrounds)},
        "palette": [c for c in dict.fromkeys(backgrounds) if c is not None],
        "rotation": {"order": order, "confidence": round(confidence, 3)},
        "bookends": bookends,
        "breathers": breathers,
    }


def rotation_targets(profile, n_slides):
    """The color each generated-slide position SHOULD get, following the
    discovered rotation with bookend anchors. Empty list when the template
    has no rotation — callers then skip color preferences entirely."""
    order = profile["rotation"]["order"]
    if not order:
        return []
    targets = []
    start = profile["bookends"].get("title")
    offset = order.index(start) if start in order else 0
    for i in range(n_slides):
        targets.append(order[(offset + i) % len(order)])
    closing = profile["bookends"].get("closing")
    if closing in order and n_slides > 1:
        targets[-1] = closing
        # keep the second-to-last different from the (possibly reassigned) last
        if n_slides > 2 and targets[-2] == targets[-1]:
            alternatives = [c for c in order if c != targets[-1]]
            if alternatives:
                targets[-2] = alternatives[0]
    return targets
