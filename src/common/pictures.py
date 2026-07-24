"""Which big pictures make a slide unusable as a text-fill target.

The original binary rule ("any picture over 12% of the slide → exclude the
slide") was written against survey decks, where every big picture was a chart
screenshot: filling new text around someone's old data is actively misleading.
Then an illustration-driven brand template (T-Ж) arrived where EVERY slide has
a full-bleed background image or a decorative photo — the rule excluded 11 of
12 slides and the generator synthesized the whole deck from scratch, discarding
the very template the user uploaded.

The classes separate cleanly on real data (measured on both corpora, see
tests/test_picture_classes.py):

- Background art: covers ~100% of the slide and the slide's own text boxes sit
  ON TOP of it (charts: 14-34% area, zero text boxes on top). A picture the
  designer put text over is scenery, not data.
- Photos/illustrations: hundreds of quantized colors / no dominant flat color
  (charts: 80-91% near-white dominant, almost no saturated pixels).
- Chart/screenshot look: flat, near-white dominant, desaturated — and standing
  apart from the text. Only THIS combination carries stale-data risk.
"""

import io
from collections import Counter

from pptx.enum.shapes import MSO_SHAPE_TYPE

OVERSIZED_PICTURE_AREA_RATIO = 0.12

# Structure: a picture most of whose slide's text boxes lie on top of it is a
# background regardless of its pixels.
_TEXT_ON_TOP_SHARE = 0.5
_TEXT_BOX_OVERLAP = 0.5

# Pixels: measured chart signature vs photo/illustration signature. NOTE:
# quantized-color count does NOT separate the classes — a rich chart hits 149
# colors via anti-aliasing, overlapping the photos' 167-456 range. The flat
# dominant-background share does: photos peak at 0.27, charts start at 0.80.
_CHART_DOM_SHARE = 0.6    # charts: 0.80-0.91; photos: 0.08-0.27
_CHART_DOM_LUMINANCE = 200  # charts: ~240 (near-white)
_CHART_MAX_SATURATED = 0.15  # charts: 0.00-0.07; brand art: 0.4-0.85


def _pixel_stats(blob):
    from PIL import Image

    im = Image.open(io.BytesIO(blob))
    if im.mode != "RGB":
        im = im.convert("RGBA")
        background = Image.new("RGBA", im.size, (255, 255, 255, 255))
        # Transparent PNGs read as black once naively converted to RGB — every
        # stat below would then be garbage (measured before this composite).
        im = Image.alpha_composite(background, im).convert("RGB")
    im.thumbnail((128, 128))
    raw = im.tobytes()
    pixels = [(raw[i], raw[i + 1], raw[i + 2]) for i in range(0, len(raw), 3)]
    total = len(pixels)

    quantized = Counter((r >> 4, g >> 4, b >> 4) for r, g, b in pixels)
    (dom, dom_n), = quantized.most_common(1)
    dom_rgb = tuple(v << 4 for v in dom)
    return {
        "dom_share": dom_n / total,
        "dom_luminance": 0.299 * dom_rgb[0] + 0.587 * dom_rgb[1] + 0.114 * dom_rgb[2],
        "colors": len(quantized),
        "saturated": sum(1 for r, g, b in pixels if max(r, g, b) - min(r, g, b) > 60) / total,
    }


def _covers_slide_text(picture, slide):
    texts = [
        s for s in slide.shapes
        if s.has_text_frame and s.text_frame.text.strip()
        and s.left is not None and s.top is not None and s.width and s.height
    ]
    if not texts or picture.left is None or picture.top is None:
        return False
    on_top = 0
    for t in texts:
        x1 = max(picture.left, t.left)
        y1 = max(picture.top, t.top)
        x2 = min(picture.left + picture.width, t.left + t.width)
        y2 = min(picture.top + picture.height, t.top + t.height)
        if x2 > x1 and y2 > y1 and (x2 - x1) * (y2 - y1) / (t.width * t.height) > _TEXT_BOX_OVERLAP:
            on_top += 1
    return on_top / len(texts) > _TEXT_ON_TOP_SHARE


def is_stale_data_picture(picture, slide):
    """True only for the picture kind whose content goes stale next to new
    text: a chart/screenshot standing apart from the slide's text. Background
    art and photos return False — they're scenery and stay."""
    if _covers_slide_text(picture, slide):
        return False
    try:
        stats = _pixel_stats(picture.image.blob)
    except Exception:
        return True  # unreadable image on a text-apart picture: keep the safe behavior
    if stats["dom_share"] < _CHART_DOM_SHARE:
        return False  # photo / rich illustration: no flat dominant background
    return (
        stats["dom_luminance"] >= _CHART_DOM_LUMINANCE
        and stats["saturated"] <= _CHART_MAX_SATURATED
    )


def has_oversized_picture(slide, slide_width, slide_height):
    """True when the slide carries a big STALE-DATA picture (see module doc) —
    the name is kept for existing call sites, but decor no longer triggers it."""
    slide_area = (slide_width or 0) * (slide_height or 0)
    if not slide_area:
        return False
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            area = (shape.width or 0) * (shape.height or 0)
            if area / slide_area > OVERSIZED_PICTURE_AREA_RATIO and is_stale_data_picture(shape, slide):
                return True
    return False
