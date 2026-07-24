"""Distance-based structural clustering of slides.

Replaces exact-signature grouping (the old signature.py tuple equality): real
decks rarely repeat a layout recipe *exactly* — a bullet slide with 3 text
boxes and its sibling with 5 are the same archetype but different multisets,
so grouping has to be a similarity judgement, not an equality check.

Two signals, in order of authority:

1. Structural distance (free, from parsed OOXML): Dice overlap of the two
   slides' shape multisets at two granularities — with and without the
   position cell — so identical recipes with jittered geometry or a different
   item count still match partially instead of not at all.
2. Visual penalty (free, from renders we already produce for the UI): a coarse
   grayscale comparison of the rendered PNGs. It exists to catch what the
   parser cannot see — backgrounds and graphics that live on the layout or
   master, not the slide. Two slides whose shape lists look alike but whose
   renders are wildly different must not merge. The penalty is soft (added to
   the distance), not a hard veto: a single noisy pair — big display text at
   different spots inflates the gradient signal — must not block a merge that
   every other slide pair in the two clusters supports.

LLM classification then runs once per cluster (on the medoid), not per slide,
and each cluster carries a deterministic geometry fingerprint so results can
be reused across templates (see fingerprint_cache.py) — the fingerprint holds
no client text, only shape roles and quantized geometry.
"""

import hashlib
from collections import Counter

from PIL import Image

GRID_COLS = 4
GRID_ROWS = 3
# Average-linkage cut-off: pairs of slides whose blended Dice distance exceeds
# this never end up in one cluster. Calibrated on real uploaded decks — low
# enough to keep title/divider/content recipes apart, high enough to merge
# "same recipe, different item count" siblings.
MERGE_THRESHOLD = 0.4
# Visual difference below this is noise (same recipe, different chart values /
# photos / text positions — measured 0.17-0.39 on real same-recipe pairs);
# above it, the excess feeds into the distance steeply enough that a genuine
# master-level difference (dark vs light background: 0.5-0.8) overshoots
# MERGE_THRESHOLD on its own even when the shape lists match exactly.
VISUAL_PENALTY_START = 0.35
VISUAL_PENALTY_WEIGHT = 2.0
_VISUAL_GRID = 8


def _size_bucket(geo, slide_area):
    area = (geo["width"] or 0) * (geo["height"] or 0)
    frac = area / slide_area if slide_area else 0
    if frac < 0.02:
        return "small"  # icon-sized
    if frac < 0.2:
        return "medium"
    return "large"  # background image / full content block


def _grid_cell(geo, slide_w, slide_h):
    if geo["left"] is None or geo["top"] is None or not slide_w or not slide_h:
        return None
    cx = geo["left"] + (geo["width"] or 0) / 2
    cy = geo["top"] + (geo["height"] or 0) / 2
    col = min(int(cx / slide_w * GRID_COLS), GRID_COLS - 1)
    row = min(int(cy / slide_h * GRID_ROWS), GRID_ROWS - 1)
    return (col, row)


def _text_bucket(shape):
    if not shape.get("text"):
        return "none"
    chars = sum(len(run["text"]) for para in shape["text"] for run in para)
    if chars == 0:
        return "none"
    return "short" if chars < 100 else "long"


def slide_features(slide_struct, slide_size):
    """Two multisets of shape descriptors: "coarse" ignores position so that the
    same recipe drawn at slightly different spots still overlaps fully, "placed"
    adds the grid cell so that genuinely different arrangements drift apart."""
    slide_w = (slide_size or {}).get("width")
    slide_h = (slide_size or {}).get("height")
    slide_area = (slide_w or 0) * (slide_h or 0)

    coarse, placed = Counter(), Counter()
    for shape in slide_struct["shapes"]:
        role = shape.get("placeholder_type") or shape["shape_type"]
        geo = shape["geometry_in"]
        size = _size_bucket(geo, slide_area)
        text = _text_bucket(shape)
        coarse[(role, size, text)] += 1
        placed[(role, size, text, _grid_cell(geo, slide_w, slide_h))] += 1
    return {"coarse": coarse, "placed": placed}


def _dice_distance(a, b):
    total = sum(a.values()) + sum(b.values())
    if total == 0:
        return 0.0  # two empty slides: same (blank) recipe
    overlap = sum((a & b).values())
    return 1.0 - 2.0 * overlap / total


def structural_distance(f1, f2):
    return 0.5 * _dice_distance(f1["coarse"], f2["coarse"]) + 0.5 * _dice_distance(f1["placed"], f2["placed"])


def slide_fingerprint(features):
    """Deterministic hash of the placement-aware multiset — the cross-template
    cache key. Geometry and roles only; no text ever enters the digest."""
    items = sorted((str(key), count) for key, count in features["placed"].items())
    digest = hashlib.sha256(repr(items).encode("utf-8")).hexdigest()
    return digest[:16]


def _visual_profile(png_path):
    """Tiny grayscale thumbnail + its horizontal gradient signs. The thumbnail
    catches flat-but-different backgrounds (dark divider vs light content
    slide), the gradients catch layout skeleton; either alone has blind spots."""
    with Image.open(png_path) as img:
        gray = img.convert("L").resize((_VISUAL_GRID + 1, _VISUAL_GRID), Image.LANCZOS)
        # mode "L" bytes are row-major pixel values — same data getdata() gave,
        # without the deprecated accessor.
        pixels = list(gray.tobytes())
    thumb, gradients = [], []
    for row in range(_VISUAL_GRID):
        for col in range(_VISUAL_GRID):
            left = pixels[row * (_VISUAL_GRID + 1) + col]
            right = pixels[row * (_VISUAL_GRID + 1) + col + 1]
            thumb.append(left)
            gradients.append(left > right)
    return thumb, gradients


def visual_distance(profile_a, profile_b):
    thumb_a, grad_a = profile_a
    thumb_b, grad_b = profile_b
    brightness = sum(abs(x - y) for x, y in zip(thumb_a, thumb_b)) / (len(thumb_a) * 255)
    hamming = sum(x != y for x, y in zip(grad_a, grad_b)) / len(grad_a)
    # max, not average: either signal alone is enough to prove "these differ".
    return max(brightness, hamming)


def _distance_matrix(slides, slide_size, png_paths):
    n = len(slides)
    features = [slide_features(s, slide_size) for s in slides]
    profiles = None
    if png_paths and len(png_paths) >= n:
        try:
            profiles = [_visual_profile(png_paths[i]) for i in range(n)]
        except Exception:
            profiles = None  # renders are an optional extra signal, never a hard dependency

    dist = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            d = structural_distance(features[i], features[j])
            if profiles:
                excess = visual_distance(profiles[i], profiles[j]) - VISUAL_PENALTY_START
                if excess > 0:
                    d += excess * VISUAL_PENALTY_WEIGHT
            dist[i][j] = dist[j][i] = d
    return features, dist


def _average_linkage(n, dist, threshold):
    clusters = [[i] for i in range(n)]
    while len(clusters) > 1:
        best, best_d = None, threshold
        for a in range(len(clusters)):
            for b in range(a + 1, len(clusters)):
                pairs = [dist[i][j] for i in clusters[a] for j in clusters[b]]
                d = sum(pairs) / len(pairs)
                if d <= best_d:
                    best, best_d = (a, b), d
        if best is None:
            break
        a, b = best
        clusters[a] = clusters[a] + clusters[b]
        del clusters[b]
    return clusters


def _medoid(member_positions, dist):
    return min(
        member_positions,
        key=lambda i: sum(dist[i][j] for j in member_positions if j != i),
    )


def cluster_slides(slides, slide_size, png_paths=None, threshold=MERGE_THRESHOLD):
    """slides: list of slide_struct dicts (with an "index" key).
    Returns a list of {"indices": [slide indices], "medoid": slide index,
    "fingerprint": str} — one entry per structural cluster."""
    if not slides:
        return []
    features, dist = _distance_matrix(slides, slide_size, png_paths)
    groups = _average_linkage(len(slides), dist, threshold)

    clusters = []
    for positions in groups:
        medoid_pos = _medoid(positions, dist)
        clusters.append({
            "indices": [slides[p]["index"] for p in positions],
            "medoid": slides[medoid_pos]["index"],
            "fingerprint": slide_fingerprint(features[medoid_pos]),
        })
    clusters.sort(key=lambda c: c["indices"][0])
    return clusters
