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

import hashlib

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

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


def _appears_on_other_slides(picture, slide):
    """True when the identical image bytes are used on another slide too.

    A chart or a screenshot is made FOR its slide and appears once; brand
    artwork is placed wherever the designer wants it. Measured over the 138
    large pictures of the real corpus that this module would otherwise call
    stale data: exactly four repeat, one per survey deck, and every one of them
    is the brand illustration that opens the deck and closes it on the contacts
    slide. Every real chart is unique.

    The pixel signature cannot make this call — the contacts illustration is two
    flat colours and so is a monochrome bar chart, 8 distinct shades against 9
    (measured). This is structural, and structure is what the deck states rather
    than what an image happens to look like."""
    try:
        blob = picture.image.blob
        slides = slide.part.package.presentation_part.presentation.slides
    except Exception:
        return False
    digest = hashlib.sha1(blob).hexdigest()
    seen = 0
    for other in slides:
        for shape in other.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            try:
                if hashlib.sha1(shape.image.blob).hexdigest() == digest:
                    seen += 1
            except Exception:
                continue
            if seen > 1:
                return True
    return False


def is_stale_data_picture(picture, slide):
    """True only for the picture kind whose content goes stale next to new
    text: a chart/screenshot standing apart from the slide's text. Background
    art and photos return False — they're scenery and stay."""
    if _covers_slide_text(picture, slide):
        return False
    if _appears_on_other_slides(picture, slide):
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


# A chart drawn with SHAPES: a row of bars is a run of same-thickness shapes
# whose lengths differ. Measured over the corpus: 11 slides, all VK Tech (its
# funnel, its Gantt chart, its dot-and-bar decorations) and zero on the other
# twelve templates — card grids do not match, their cards are equal in both
# dimensions.
SHAPE_CHART_MIN_BARS = 4
SHAPE_CHART_MIN_LENGTHS = 3
SHAPE_CHART_THICKNESS_TOLERANCE_EMU = int(Inches(0.04))


def has_shape_chart(slide):
    """True when the slide's «infographic» is a chart built out of shapes.

    Its bars carry the designer's numbers in their LENGTHS, and nothing resizes
    them (backlog item 5): VK Tech's funnel shipped as «до и после пилота» with
    our labels beside bars that mean nothing, and its Gantt chart shipped as a
    KPI board — «-5 недель» and «1,3 млн ₽» against bars of the designer's own
    proportions, one bar without a label at all (iter141, iter144). Like a
    table or a flowchart, such a slide is not offered until we can draw one."""
    def walk(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from walk(shape.shapes)
            else:
                yield shape

    plain = [s for s in walk(slide.shapes)
             if s.width and s.height and s.left is not None and s.top is not None
             and not (s.has_text_frame and s.text_frame.text.strip())]
    # Horizontal bars only. The vertical reading (equal width, varying height)
    # was measured and dropped inside this iteration: on the corpus it found no
    # chart of its own and one false positive — VK Tech 54's product cards, four
    # illustrations 1.69in wide and 1.54-2.26in tall, on a perfectly good list
    # slide. All three real cases (VK Tech's funnel and Gantt, VK Education's
    # «Диаграмма Ганта») are rows of bars.
    groups = []
    for shape in plain:
        for group in groups:
            if abs(int(group[0].height) - int(shape.height)) <= SHAPE_CHART_THICKNESS_TOLERANCE_EMU:
                group.append(shape)
                break
        else:
            groups.append([shape])
    for group in groups:
        if len(group) < SHAPE_CHART_MIN_BARS:
            continue
        lengths = {round(int(s.width) / 914400, 2) for s in group}
        rows = {int(s.top) for s in group}
        if len(lengths) >= SHAPE_CHART_MIN_LENGTHS and len(rows) >= SHAPE_CHART_MIN_BARS - 1:
            return True
    return False


def has_data_object(slide):
    """True when the slide carries a native TABLE or CHART — the designer's
    sample data in its most literal form.

    Nothing fills them yet (native tables and charts are backlog item 5), so a
    slide matched for its text boxes ships the table as the template left it:
    on VK Education «Заголовок столбца, млн», «Акцент», «Строка» and demo
    figures 15/10/14/4 stood next to our bullets (iter109). The same stale-data
    reasoning as has_oversized_picture, without a pixel heuristic — the document
    SAYS what the object is. Groups are opened: a graphic frame can sit inside
    one."""
    def walk(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                if walk(shape.shapes):
                    return True
            elif getattr(shape, "has_table", False) and shape.has_table:
                return True
            elif getattr(shape, "has_chart", False) and shape.has_chart:
                return True
        return False

    return walk(slide.shapes) or has_connector_diagram(slide) or has_shape_chart(slide)


# A flowchart is a graph: many connectors, several of them arrows. Measured over
# the corpus: VK Education's «Оформление схем» example — 15 connector lines, 9
# arrowed; the next highest slides are a Gantt chart's gridlines (6 lines, no
# arrows), T-Zh mono's column separators (4, none) and a WorkSpace timeline
# «01 → 02 → 03 → 04» (3 lines, 3 arrows — a legitimate list layout).
DIAGRAM_MIN_CONNECTORS = 6
DIAGRAM_MIN_ARROWS = 3


def _connector_lines(slide):
    from pptx.oxml.ns import qn

    def walk(shapes):
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from walk(shape.shapes)
            elif shape.shape_type == MSO_SHAPE_TYPE.LINE or shape._element.tag == qn("p:cxnSp"):
                yield shape

    return list(walk(slide.shapes))


def _is_arrowed(shape):
    from pptx.oxml.ns import qn

    ln = shape._element.find(".//" + qn("a:ln"))
    if ln is None:
        return False
    return any(end is not None and end.get("type") not in (None, "none")
               for end in (ln.find(qn("a:headEnd")), ln.find(qn("a:tailEnd"))))


def has_connector_diagram(slide):
    """True when the slide is a flowchart the designer drew as an example.

    Its boxes are the nodes of someone else's diagram, not the slots of a list:
    VK Education's «Оформление схем» (a template guide slide) was classified
    as a list, the one node with text — a 1.26in circle — took the bullets, the
    other nodes stayed empty, and the render showed «Необход/имые действия
    руководс/тва» broken inside the circle among blank boxes and arrows
    (iter130). Nothing fills diagrams yet (backlog item 5), so like a table it
    is not offered."""
    lines = _connector_lines(slide)
    return (len(lines) >= DIAGRAM_MIN_CONNECTORS
            and sum(1 for s in lines if _is_arrowed(s)) >= DIAGRAM_MIN_ARROWS)
