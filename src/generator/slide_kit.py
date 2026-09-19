"""Shared slide-level primitives used by both the filler side (generator.py)
and the synthesis side (synthesizer.py) — split out so synthesis can clone a
template slide as its canvas (plan 9.3) without a circular import.

Everything here carries hard-won invariants; see CLAUDE.md before touching:
- clone_slide must NOT copy slideLayout/notesSlide relationships (single-owner
  reltypes; sharing a notesSlide across clones made PowerPoint offer to
  "repair" every generated deck);
- chrome detection is geometric because real decks draw running headers and
  footers as PLAIN textboxes with no placeholder type to key on.
"""

import copy

from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches

_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Slide number / date / footer placeholders often render via a <a:fld> field
# code rather than a literal run — presentation chrome, never fill targets.
BORING_PLACEHOLDER_TYPES = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}

# Chrome drawn as plain text boxes ("ТЕМА ПРЕЗЕНТАЦИИ" strips, page numbers,
# "КОММЕНТАРИЙ" captions): short boxes hugging the top/bottom edge bands.
CHROME_MAX_HEIGHT_EMU = Emu(int(Inches(0.35)))
CHROME_BAND_FRACTION = 0.14


def is_boring_placeholder(shape):
    return shape.is_placeholder and shape.placeholder_format.type in BORING_PLACEHOLDER_TYPES


def text_shapes(slide):
    return [
        s for s in slide.shapes
        if s.has_text_frame and s.text_frame.text.strip() and not is_boring_placeholder(s)
    ]


class PlacedShape:
    """A shape inside a GROUP, with its box where the slide draws it.

    python-pptx reports a group member's position in the group's OWN
    coordinate space (chOff/chExt), not the slide's: of the 25 groups in VK
    Tech and VK Education only one maps 1:1. Every geometric reader — edges,
    margins, text bands, the pixels under a box — needs slide coordinates.
    Everything else (text, runs, insets, shape_id) is the real shape's. Font
    sizes are not scaled: a group transform resizes boxes, not the text."""

    def __init__(self, shape, left, top, width, height):
        self._shape = shape
        self.left, self.top = Emu(int(left)), Emu(int(top))
        self.width, self.height = Emu(int(width)), Emu(int(height))

    def __getattr__(self, name):
        return getattr(self._shape, name)


def _group_transform(group):
    """(x, y, w, h) in the group's child space -> in its parent's space."""
    xfrm = group._element.grpSpPr.xfrm
    if (xfrm is None or xfrm.off is None or xfrm.ext is None or xfrm.chOff is None
            or xfrm.chExt is None or not xfrm.chExt.cx or not xfrm.chExt.cy):
        return lambda x, y, w, h: (x, y, w, h)
    sx, sy = xfrm.ext.cx / xfrm.chExt.cx, xfrm.ext.cy / xfrm.chExt.cy
    ox, oy, cx, cy = xfrm.off.x, xfrm.off.y, xfrm.chOff.x, xfrm.chOff.y
    return lambda x, y, w, h: (ox + (x - cx) * sx, oy + (y - cy) * sy, w * sx, h * sy)


def placed_shapes(shapes, place=None):
    """Every non-group shape, groups opened recursively, each at its place on
    the slide. Top-level shapes come back as they are."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            inner = _group_transform(s)
            composed = inner if place is None else (
                lambda x, y, w, h, i=inner, o=place: o(*i(x, y, w, h)))
            yield from placed_shapes(s.shapes, composed)
        elif place is None or None in (s.left, s.top, s.width, s.height):
            yield s
        else:
            yield PlacedShape(s, *place(s.left, s.top, s.width, s.height))


def slide_height(slide):
    return slide.part.package.presentation_part.presentation.slide_height


def slide_width(slide):
    return slide.part.package.presentation_part.presentation.slide_width


def is_chrome_shape(shape, height):
    if shape.height is None or shape.top is None or shape.height > CHROME_MAX_HEIGHT_EMU:
        return False
    band = int(height * CHROME_BAND_FRACTION)
    return shape.top <= band or (shape.top + shape.height) >= height - band


def content_text_shapes(slide, claimed_ids=()):
    height = slide_height(slide)
    return [
        s for s in text_shapes(slide)
        if s.shape_id not in claimed_ids and not is_chrome_shape(s, height)
    ]


def _group_has_text(group):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for child in group.shapes:
        if child.shape_type == MSO_SHAPE_TYPE.GROUP:
            if _group_has_text(child):
                return True
        elif getattr(child, "has_text_frame", False) and child.text_frame.text.strip():
            return True
    return False


def content_groups(slide):
    """Top-level GROUPS that carry the slide's own content text.

    text_shapes() sees only the top level of slide.shapes, so a template that
    keeps each card as a group (VK Tech: icon + «Заголовок / Подзаголовок /
    Текст» ×7 per card) hands its whole content past every caller. A canvas
    cloned for synthesis then kept four placeholder cards under our title and
    image frame (iter102, iter105). A group in the chrome band is furniture and
    is not returned; chrome is judged on the GROUP's own box, since a child's
    coordinates live in the group's space, not the slide's."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    height = slide_height(slide)
    return [
        s for s in slide.shapes
        if s.shape_type == MSO_SHAPE_TYPE.GROUP
        and not is_chrome_shape(s, height)
        and _group_has_text(s)
    ]


def canvas_content_count(slide):
    """How much content a slide carries, for choosing a synthesis canvas: its
    content text boxes AND its content groups.

    Counting text boxes alone made VK Tech's four-card slide look like bare
    furniture — its cards are groups, only the title is a top-level box — so it
    was picked as the canvas for synthesized slides. Synthesis removes the
    cards (content_groups) but not the single backing shape with their
    cut-outs, and the image skeleton and two-column text were drawn across four
    empty card backings on every such deck (iter105, 111, 114, 117, 120)."""
    return len(content_text_shapes(slide)) + len(content_groups(slide))


def clone_slide(prs, source_idx):
    """Duplicates a template slide (shapes, images, formatting) as a new slide
    at the end of the deck, returns its index. Exists so a deck that needs two
    bullet-list slides from a template that only designed one gets two slides
    in the template's own visual language — synthesizing a bare substitute from
    scratch loses the icons, backgrounds and typography that make the template
    worth using in the first place."""
    source = prs.slides[source_idx]
    clone = prs.slides.add_slide(source.slide_layout)
    # add_slide pre-populates placeholders inherited from the layout — drop
    # them, the deep-copied originals below replace them wholesale.
    for shape in list(clone.shapes):
        shape._element.getparent().remove(shape._element)

    rid_map = {}
    for rid, rel in source.part.rels.items():
        if rel.reltype.endswith("/slideLayout"):
            continue  # the clone already has its own layout relationship
        if rel.reltype.endswith("/notesSlide"):
            # A notes-slide part is meant to belong to exactly ONE slide (no
            # back-relationship exists in the other direction — ownership is
            # this single edge). Blindly copying it made every clone of a
            # slide point at the SAME notesSlide part: one real deck ended up
            # with 5 different slides all declaring notesSlide3.xml as their
            # notes, which is what actually triggered PowerPoint's "found a
            # problem with content" repair prompt (the earlier sldIdLst-orphan
            # fix addressed a real but separate defect, not this one). The
            # clone's fresh generated content has no meaningful connection to
            # the template's original speaker notes anyway — drop it.
            continue
        if rel.is_external:
            rid_map[rid] = clone.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            rid_map[rid] = clone.part.relate_to(rel.target_part, rel.reltype)

    # The slide-level background <p:bg> lives in <p:cSld>, a SIBLING of the shape
    # tree — copying shapes alone drops it, so a cloned slide whose template used
    # a <p:bg> solid/gradient fill (not a full-bleed PICTURE) falls through to the
    # master's white background. That left synthesized-canvas and reused-slide
    # clones unreadable (light template text on white — the contrast backstop
    # caught it on the loop's own decks). Copy <p:bg> across, remapping any rels
    # it carries (a blipFill picture background references an image rel).
    src_cSld = source._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg")) if src_cSld is not None else None
    if src_bg is not None:
        bg = copy.deepcopy(src_bg)
        for node in bg.iter():
            for key, value in node.attrib.items():
                if key.startswith("{%s}" % _RELS_NS) and value in rid_map:
                    node.set(key, rid_map[value])
        clone_cSld = clone._element.find(qn("p:cSld"))
        existing = clone_cSld.find(qn("p:bg"))
        if existing is not None:
            clone_cSld.remove(existing)
        clone_cSld.insert(0, bg)  # schema: <p:bg> must precede <p:spTree>

    for shape in source.shapes:
        element = copy.deepcopy(shape._element)
        for node in element.iter():
            for key, value in node.attrib.items():
                if key.startswith("{%s}" % _RELS_NS) and value in rid_map:
                    node.set(key, rid_map[value])
        clone.shapes._spTree.append(element)

    return len(prs.slides._sldIdLst) - 1
