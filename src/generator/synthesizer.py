"""Builds slides from scratch, for archetypes the uploaded template simply has no
slide for at all — using the template's own theme (fonts/colors) and its own
real content geometry (see layout_bounds.py) so the result still looks like it
belongs, rather than colliding with master-level decorations we can't even see
(a logo/footer baked into the slide master isn't in slide.shapes) or using
fixed percentages that happened to work for one file and not the next."""

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

from generator.deck_style import _luminance, _slide_bg_hex
from generator.slide_kit import clone_slide, content_text_shapes
from generator.text_fit import (cap_size_to_longest_word, estimate_block_height_in,
                                fit_font_size)

# Used only when the template has too few real content shapes to infer bounds
# from (see layout_bounds.infer_content_bounds returning None).
FALLBACK_BOUNDS_PCT = {"top": 0.14, "left": 0.06, "right": 0.94, "bottom": 0.90}


def _hex_to_rgb(hex_str):
    if not hex_str:
        return None
    h = hex_str.lstrip("#")
    try:
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None


def _style_run(run, font_name, size_pt, color_hex, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    color = _hex_to_rgb(color_hex)
    if color:
        run.font.color.rgb = color


def _style_paragraph(paragraph, font_name, size_pt, color_hex, bold=False):
    """paragraph.runs is a tuple — for an empty-string paragraph (e.g. a title
    that came back "" from content parsing) it's empty, so paragraph.runs[0]
    raises "tuple index out of range". Add a run first if there isn't one."""
    if not paragraph.runs:
        paragraph.add_run()
    _style_run(paragraph.runs[0], font_name, size_pt, color_hex, bold=bold)


def _theme_palette(theme):
    return {
        "major_font": theme["fonts"].get("majorFont") or "Calibri",
        "minor_font": theme["fonts"].get("minorFont") or "Calibri",
        "accent": theme["palette"].get("accent1") or "#1F497D",
        "text": theme["palette"].get("dk1") or "#000000",
        "bg": theme["palette"].get("lt1"),
        # Measured from the template (generator._template_title_pt); None when
        # the deck gives no confident answer, and then _add_title's own default
        # stands.
        "title_pt": theme.get("title_pt"),
        "cover_pt": theme.get("cover_pt"),
    }


def _resolve_bounds(prs, bounds_in):
    """bounds_in: output of layout_bounds.infer_content_bounds (inches) or None."""
    width, height = prs.slide_width, prs.slide_height
    if bounds_in:
        return {
            "top": Emu(int(Inches(bounds_in["top_in"]))),
            "left": Emu(int(Inches(bounds_in["left_in"]))),
            "right": Emu(int(Inches(bounds_in["right_in"]))),
            "bottom": Emu(int(Inches(bounds_in["bottom_in"]))),
        }
    return {
        "top": Emu(int(height * FALLBACK_BOUNDS_PCT["top"])),
        "left": Emu(int(width * FALLBACK_BOUNDS_PCT["left"])),
        "right": Emu(int(width * FALLBACK_BOUNDS_PCT["right"])),
        "bottom": Emu(int(height * FALLBACK_BOUNDS_PCT["bottom"])),
    }




# Below this |luminance delta| text is unreadable on its background. Same
# threshold deck_style uses for its deck-wide veto — see _recolor_for_canvas for
# why that veto isn't enough on cloned canvases.
_MIN_CANVAS_CONTRAST = 80

# Below this a placeholder frame is too thin to read as a reserved image area.
MIN_FRAME_HEIGHT = int(Inches(0.6))

# Used only when the template's own title size cannot be measured.
DEFAULT_TITLE_PT = 32

# Used only when the template's own cover size cannot be measured.
DEFAULT_COVER_PT = 40
# A cover that shrinks below this is no longer a cover; a title that long is a
# content problem, not a sizing one.
_MIN_COVER_PT = 28


def _recolor_for_canvas(palette, canvas_bg):
    """deck_style already vetoes an unreadable text colour, but it compares
    against the DECK-WIDE background (the majority of slides). A cloned canvas
    keeps ITS OWN background, which the template's colour rotation may well have
    made a different colour — the deck-wide-approved text then lands on it and
    can vanish. Real defect: the T-Zh mono deck's synthesized comparison slide
    got blue #213FFF body text (correct on that deck's white slides, verified
    readable on its white closing) painted onto a blue canvas. The bullets were
    invisible, and the pixel contrast backstop couldn't see it either — text
    merged into its background forms no ink cluster at all.

    canvas_bg is the MEASURED background of the canvas slide ((r,g,b) from the
    render), not the XML one. The XML is not usable here: every slide of that
    template declares the same <a:schemeClr val="lt1"/>, which the theme maps to
    #FFFFFF, while the render shows four of them solid blue — the theme lies
    about the deck's real look exactly as CLAUDE.md warns. Falls back to an
    explicit srgb background when no measurement was supplied."""
    if canvas_bg is None:
        return palette
    if isinstance(canvas_bg, (tuple, list)):
        r, g, b = canvas_bg[:3]
        bg_hex = f"#{r:02X}{g:02X}{b:02X}"
    else:
        bg_hex = canvas_bg
    out = dict(palette)
    bg_lum = _luminance(bg_hex)
    readable = "#FFFFFF" if bg_lum < 128 else "#1A1A1A"
    for key in ("text", "accent"):
        colour = out.get(key)
        if colour and abs(_luminance(colour) - bg_lum) < _MIN_CANVAS_CONTRAST:
            out[key] = readable
    return out


# A cloned canvas's own text area is only a usable content area if it is a
# reasonable share of the layout the template designs for. Below this the slide
# was never a content slide (a divider's single line) and its band starves
# whatever role is being synthesized onto it. Measured: the three narrow
# canvases per T-Zh template sit at 20-39% of the full band, real content
# slides well above it.
_MIN_CANVAS_BAND_FRACTION = 0.4

# …and it also has to be big enough in absolute terms. A band can clear the
# fraction above and still be too small for a title plus a body: the survey-31
# canvas offers 2.82in on a 7.5in slide (47% of the template band, just over the
# cut), while a heading and three KPI pairs need about 2.9in. The content then
# overflowed the band's bottom edge — which is exactly where that canvas keeps a
# small decorative heart, so the render showed the artwork sitting on top of
# «клиентов в месяц». Measured over the corpus, this second test moves 7 more
# canvases onto the template-wide band, all of them on the wide survey decks.
_MIN_CANVAS_BAND_SLIDE_FRACTION = 0.4


def _band_height(bounds):
    return int(bounds["bottom"]) - int(bounds["top"])


# Clearance between synthesized content and the canvas's own header/footer.
_CHROME_GAP_EMU = int(Inches(0.12))

# A decor strip is only worth stepping around sideways if it runs down most of
# the band (otherwise the vertical avoidance in _centered_top handles it) and is
# narrow enough that giving it up costs little width.
_SIDE_STRIP_MIN_COVERAGE = 0.5
_SIDE_STRIP_MAX_WIDTH_FRACTION = 0.2


def _widen_to_template_if_clear(bounds, template_bounds, slide, slide_width, slide_height):
    """Grow a canvas-derived band sideways to the template's own width, but only
    into space that is actually empty.

    The canvas's own text extent is a good guide until it is simply narrower
    than the layout: the T-Zh study canvas ends its text at 7.59in on a card
    running to 9.6in, so the synthesized comparison got 3.50in columns and the
    render showed «Автоматический сбор по / расписанию» wrapping while two
    inches of card sat empty beside it. Same argument iter32 makes for a band
    that is too SHORT, applied to width.

    Blind widening would be wrong — plenty of canvases are narrow BY DESIGN,
    with art filling the other half — so each side only moves if no picture
    stands in the strip it would move into. A full-bleed background is not an
    obstacle: everything sits on it anyway."""
    left, right = int(bounds["left"]), int(bounds["right"])
    top, bottom = int(bounds["top"]), int(bounds["bottom"])
    area = (slide_width or 1) * (slide_height or 1)

    def _blocked(lo, hi):
        for shape in slide.shapes:
            if "PICTURE" not in str(shape.shape_type):
                continue
            if None in (shape.left, shape.top) or not (shape.width and shape.height):
                continue
            if (shape.width * shape.height) / area >= 0.9:
                continue  # full-bleed background: not an obstacle
            if int(shape.left) < hi and int(shape.left + shape.width) > lo \
                    and int(shape.top) < bottom and int(shape.top + shape.height) > top:
                return True
        return False

    t_left, t_right = int(template_bounds["left"]), int(template_bounds["right"])
    if t_left < left and not _blocked(t_left, left):
        left = t_left
    if t_right > right and not _blocked(right, t_right):
        right = t_right
    return dict(bounds, left=Emu(left), right=Emu(right))


def _clip_to_side_decor(bounds, slide, slide_width, slide_height):
    """Move the content's left/right edge past a decor strip running down it.

    _centered_top steps around decor VERTICALLY, which cannot help when the art
    is a column: the survey-31 canvas keeps seven small blobs at x 0.45-1.16in
    spanning y 1.96-6.71 — 80% of the band's height — and the KPI text starts at
    x 0.45, so the render showed a blob sitting on «-40%» and on its caption.
    There is no clear horizontal band to move into, only a narrower one to start
    from."""
    left, right = int(bounds["left"]), int(bounds["right"])
    band = max(1, int(bounds["bottom"]) - int(bounds["top"]))
    max_strip = int(slide_width * _SIDE_STRIP_MAX_WIDTH_FRACTION)

    # Coverage is a property of the STRIP, not of one shape: each of those seven
    # blobs spans about 10% of the band on its own, and only together do they
    # make a column worth stepping around. Merge the spans before judging.
    at_left, at_right = [], []
    for shape in slide.shapes:
        if "PICTURE" not in str(shape.shape_type):
            continue
        if None in (shape.left, shape.top) or not (shape.width and shape.height):
            continue
        if (shape.width * shape.height) / (slide_width * slide_height) >= _OBSTACLE_MAX_AREA_FRACTION:
            continue
        if shape.width > max_strip:
            continue
        span = (max(int(shape.top), int(bounds["top"])),
                min(int(shape.top + shape.height), int(bounds["bottom"])))
        if span[1] <= span[0]:
            continue
        # Tolerance, not equality: these blobs alternate between 0.4507in and
        # 0.4537in — a few EMU — and a strict "starts at or before the band's
        # left edge" dropped four of the seven, taking coverage from 70% to 31%
        # and silently disabling the whole check.
        if int(shape.left) <= left + _CHROME_GAP_EMU and int(shape.left + shape.width) > left:
            at_left.append((span, int(shape.left + shape.width)))
        elif int(shape.left + shape.width) >= right - _CHROME_GAP_EMU and int(shape.left) < right:
            at_right.append((span, int(shape.left)))

    def _covered(spans):
        merged, total = [], 0
        for lo, hi in sorted(s for s, _ in spans):
            if merged and lo <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(merged[-1][1], hi))
            else:
                merged.append((lo, hi))
        for lo, hi in merged:
            total += hi - lo
        return total / band

    if at_left and _covered(at_left) >= _SIDE_STRIP_MIN_COVERAGE:
        left = max(edge for _, edge in at_left) + _CHROME_GAP_EMU
    if at_right and _covered(at_right) >= _SIDE_STRIP_MIN_COVERAGE:
        right = min(edge for _, edge in at_right) - _CHROME_GAP_EMU
    if right - left < slide_width * 0.3:
        return bounds  # nothing usable left; leave it alone
    return dict(bounds, left=Emu(left), right=Emu(right))


def _clip_to_canvas_chrome(bounds, slide, slide_height):
    """Keep template-wide bounds clear of the furniture the CANVAS still carries.

    The clone keeps its header strip, footer rule, page number and caption slot;
    the template-wide band was measured across all slides and can run straight
    through them. Seen on the render: the image frame's dashed border crossed
    the footer rule and boxed in «КОММЕНТАРИЙ» and the page number. Only the
    edge bands are clipped — chrome by definition hugs them — so a canvas
    without furniture is unaffected."""
    from generator.slide_kit import is_chrome_shape

    top, bottom = int(bounds["top"]), int(bounds["bottom"])
    middle = slide_height // 2
    for shape in slide.shapes:
        if shape.top is None or shape.height is None:
            continue
        if not is_chrome_shape(shape, slide_height):
            continue
        if shape.top < middle:  # header band
            top = max(top, int(shape.top + shape.height) + _CHROME_GAP_EMU)
        else:                   # footer band
            bottom = min(bottom, int(shape.top) - _CHROME_GAP_EMU)
    if bottom <= top:
        return bounds
    return dict(bounds, top=Emu(top), bottom=Emu(bottom))


def _blank_ownerless_chrome(slide, slide_height):
    """Clear the canvas's leftover chrome placeholders.

    A clone keeps the whole footer/header band, and nothing on this path ever
    blanked it — so «КОММЕНТАРИЙ» shipped verbatim in the footer of EVERY
    synthesized slide of the T-Zh study deck, and the universal template's
    stranded «Название пункта» row and the mono deck's «Расскажем, что такое
    проект…» rode along the same way. The native fill path has always blanked
    exactly this class of text; the clone path simply never got the same pass.

    _is_managed_chrome is the same predicate iter31 used to decide what a LATER
    pass owns: the page number, the brand year, the running topic slot. Those
    stay, and _fill_running_topic then writes the deck's real name into the
    slot. Everything else in the band is a template placeholder with no owner.

    Multi-run digit boxes (the survey decks' page numbers) are not "managed" and
    so are cleared: _renumber_static_slide_numbers requires a single run and
    cannot fix them, and a blank corner beats confidently showing page 62."""
    from generator.generator import _is_managed_chrome
    from generator.slide_kit import is_chrome_shape

    for shape in slide.shapes:
        if not shape.has_text_frame or not shape.text_frame.text.strip():
            continue
        if not is_chrome_shape(shape, slide_height):
            continue
        if _is_managed_chrome(shape, slide_height):
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                run.text = ""


_MARKER_MAX_SIZE = Inches(1.0)
_MARKER_ALIGN_TOL = Inches(0.06)


def _drop_orphaned_marker_column(slide, placed):
    """Delete the canvas's bullet-marker icons once the text they marked is gone.

    Decks that draw bullets as little pictures leave those pictures behind when
    the clone path strips the canvas's own text boxes: they are independent
    shapes, so nothing removes them. The survey template's synthesized list
    shipped a column of SEVEN stranded blobs down the left margin while the
    three real bullets sat below with their own «•» — the icons read as bullets
    whose text had vanished.

    The native fill path has handled this since iter18 (_reposition_bullet_icons
    deletes every icon past the last line); the clone path never got the same
    pass. Same divergence CLAUDE.md records for the chrome blanking above.

    Deliberately narrow, because a picture on a canvas is usually decor worth
    keeping. A column is three or more pictures under an inch wide, matching in
    size, sharing a left edge, standing to the left of the removed text and
    overlapping it vertically — i.e. markers for text that no longer exists.
    Measured over every real template: it fires only on the two survey decks
    (whose lists really are icon-marked, verified on the template's own render)
    and never on the three T-Zh decks, whose decor is single shapes and photos."""
    if not placed:
        return []
    small = [s for s in slide.shapes
             if "PICTURE" in str(s.shape_type)
             and s.left is not None and s.top is not None and s.width and s.height
             and s.width < _MARKER_MAX_SIZE]
    # Ownership is per text box, not against the union of them: the canvas's
    # TITLE is typically further left than the body its markers belong to
    # (survey slide 25 — title at 0.45in, body at 1.11in, icons ending at
    # 1.15in), and measuring against the union asks the markers to stand left
    # of the title, which they never do.
    pics = [q for q in small
            if any(q.left + q.width <= t.left + _MARKER_ALIGN_TOL
                   and q.top < t.top + t.height and q.top + q.height > t.top
                   for t in placed)]

    dropped = []
    for i, ref in enumerate(pics):
        column = [q for q in pics[i:]
                  if abs(q.left - ref.left) < _MARKER_ALIGN_TOL
                  and abs(q.width - ref.width) < _MARKER_ALIGN_TOL
                  and abs(q.height - ref.height) < _MARKER_ALIGN_TOL]
        if len(column) >= 3 and not any(q in dropped for q in column):
            dropped.extend(column)
    for shape in dropped:
        shape._element.getparent().remove(shape._element)
    return dropped


def _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=None, canvas_bg=None):
    """Adds a new slide (appended at the end of prs.slides) and returns
    (slide, index, palette, bounds).

    canvas_idx (plan 9.3): clone THAT template slide as the canvas — its
    background art, running header/footer and page furniture survive for free
    (a from-scratch slide loses them all: the T-Ж folder backgrounds are
    slide-level PICTUREs, invisible to theme- or master-based reconstruction).
    The canvas's own content text boxes are removed to make room, and the area
    they occupied becomes the content bounds — the template's designer already
    decided where content belongs on this slide, no need to guess.

    Without canvas_idx: the old from-scratch path (layout background + theme
    fill), used when the template offers nothing to clone."""
    if canvas_idx is not None:
        idx = clone_slide(prs, canvas_idx)
        slide = prs.slides[idx]
        palette = _recolor_for_canvas(_theme_palette(theme),
                                      canvas_bg if canvas_bg is not None else _slide_bg_hex(slide))
        removed = content_text_shapes(slide)
        placed = [s for s in removed if s.left is not None and s.top is not None and s.width and s.height]
        # Before any geometry is measured: a marker column is not decor to keep
        # clear of, it is debris of the text being stripped. Dropping it later
        # would leave _clip_to_side_decor pushing content right, away from
        # pictures that no longer exist.
        _drop_orphaned_marker_column(slide, placed)
        bounds = None
        if placed:
            bounds = {
                "left": Emu(min(s.left for s in placed)),
                "top": Emu(min(s.top for s in placed)),
                "right": Emu(max(s.left + s.width for s in placed)),
                "bottom": Emu(max(s.top + s.height for s in placed)),
            }
            # …but only if that area can actually hold content. The canvas is
            # being REUSED for a different role, and a slide whose own text is
            # one low line (a divider, a closing) hands back a sliver: T-Zh
            # study slide 10 gives 1.03in against the template's 4.04in. The
            # measured consequence was a whole slide reduced to a heading — the
            # image frame needs MIN_FRAME_HEIGHT below the title, both bands
            # came out NEGATIVE, and synthesize_image_caption silently drew
            # nothing on a full-page empty card. Every T-Zh template has three
            # such canvases, and they are exactly the ones offered for synthesis.
            if _band_height(bounds) < max(
                    _MIN_CANVAS_BAND_FRACTION * _band_height(_resolve_bounds(prs, bounds_in)),
                    _MIN_CANVAS_BAND_SLIDE_FRACTION * prs.slide_height):
                bounds = None
        if bounds is None:
            bounds = _clip_to_canvas_chrome(
                _resolve_bounds(prs, bounds_in), slide, prs.slide_height)
        bounds = _widen_to_template_if_clear(
            bounds, _resolve_bounds(prs, bounds_in), slide, prs.slide_width, prs.slide_height)
        bounds = _clip_to_side_decor(bounds, slide, prs.slide_width, prs.slide_height)
        for shape in removed:
            shape._element.getparent().remove(shape._element)
        _blank_ownerless_chrome(slide, prs.slide_height)
        return slide, idx, palette, bounds

    layout = prs.slides[0].slide_layout if len(prs.slides) else prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)
    for shape in list(slide.shapes):
        shape._element.getparent().remove(shape._element)

    palette = _theme_palette(theme)
    if palette["bg"]:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = _hex_to_rgb(palette["bg"]) or RGBColor(0xFF, 0xFF, 0xFF)

    bounds = _resolve_bounds(prs, bounds_in)
    return slide, len(prs.slides) - 1, palette, bounds


def _metrics_for(resolver, font_name):
    return resolver.metrics_for(font_name) if resolver is not None else None


# A cloned canvas keeps its decorative art, and small pieces of it sit inside the
# content area. Above this share of the slide a picture is a background or a full
# photo — unavoidable, and stepping around it would mean leaving the band.
_OBSTACLE_MAX_AREA_FRACTION = 0.2


def _decor_obstacles(slide, slide_width, slide_height):
    """Vertical spans (top, bottom) of the canvas's small decorative pictures."""
    area = (slide_width or 0) * (slide_height or 0)
    if not area:
        return []
    spans = []
    for shape in slide.shapes:
        if "PICTURE" not in str(shape.shape_type):
            continue
        if shape.top is None or not shape.width or not shape.height:
            continue
        if (shape.width * shape.height) / area >= _OBSTACLE_MAX_AREA_FRACTION:
            continue
        spans.append((int(shape.top), int(shape.top + shape.height)))
    return spans


def _centered_top(content_top, bounds_bottom, est_height_in, obstacles=()):
    """Top position that vertically centers a content block of the estimated
    height between content_top and the bottom bound. Synthesized slides have no
    designer to fill the lower half — a short block left hanging right under
    the title reads as an accidentally half-empty slide (same defect class as
    generator._center_if_underfilled, but for slides we build ourselves).
    The title stays where native slides put theirs; only the body block moves.

    obstacles: vertical spans of canvas decor the block should step around.
    Measured defect: the survey-31 canvas keeps a small hand-drawn heart at
    4.53-5.46in and the centred KPI block landed across it, so the render showed
    the artwork sitting on «клиентов в месяц». Centring is a preference; not
    colliding with the template's own artwork outranks it, so the block takes the
    nearest clear position and stays put only when the band offers none.

    Order matters here, and this was measured the hard way: on the canvas's own
    2.82in band there WAS no clear position, so this did nothing until the band
    widened (see _MIN_CANVAS_BAND_SLIDE_FRACTION). The two fixes only work
    together."""
    avail = int(bounds_bottom) - int(content_top)
    est = int(Inches(max(0.0, est_height_in)))
    if est >= avail:
        return content_top
    top = int(content_top) + (avail - est) // 2
    lo, hi = int(content_top), int(bounds_bottom) - est

    def clashes(candidate):
        return any(candidate < ob_bottom and candidate + est > ob_top
                   for ob_top, ob_bottom in obstacles)

    if not clashes(top):
        return Emu(top)
    clear = [c for ob_top, ob_bottom in obstacles
             for c in (ob_top - est, ob_bottom)
             if lo <= c <= hi and not clashes(c)]
    return Emu(min(clear, key=lambda c: abs(c - top)) if clear else top)


def _title_fits_one_line(text, width_emu, size_pt, metrics, slack=0.1):
    """True when the title provably fits on ONE line with room to spare.

    Callers reserve space under the title, and the reserve carries a safety
    margin because the height estimate can under-count: the renderer wraps 2-4%
    earlier than fontTools metrics predict (CLAUDE.md). That margin is only
    earned when a wrap is actually possible. With no metrics we cannot tell, so
    we say no and the caller keeps its reserve."""
    if metrics is None or not text or not width_emu:
        return False
    budget_pt = Emu(width_emu).inches * 72 * (1 - slack)
    return metrics.text_width_pt(str(text), size_pt) <= budget_pt


def _add_title(slide, bounds, text, font_name, color_hex, size_pt=None, resolver=None):
    """Sized to its actual estimated line count rather than a fixed height, so
    content below it is positioned after wherever the title really ends."""
    size_pt = size_pt or DEFAULT_TITLE_PT
    width = Emu(bounds["right"] - bounds["left"])
    metrics = _metrics_for(resolver, font_name)
    # Height-only fitting accepts a size at which a long word has no choice but
    # to break mid-letter — "Конкурентоспособность" came out split at 32pt on two
    # real templates. iter18 capped the comparison columns for exactly this; the
    # title needs the same guard, and it is the most visible text on the slide.
    size_pt = cap_size_to_longest_word(text, width, size_pt, metrics=metrics, bold=True)
    height_in = estimate_block_height_in(text, Emu(width).inches, size_pt, metrics=metrics) + 0.25
    box = slide.shapes.add_textbox(bounds["left"], bounds["top"], width, Emu(int(Inches(height_in))))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    _style_paragraph(tf.paragraphs[0], font_name, size_pt, color_hex, bold=True)
    return box


def _add_bulleted_textbox(slide, left, top, width, height, lines, font_name, size_pt, color_hex):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if not lines:
        return box
    tf.text = f"•  {lines[0]}"
    _style_paragraph(tf.paragraphs[0], font_name, size_pt, color_hex)
    for line in lines[1:]:
        p = tf.add_paragraph()
        p.text = f"•  {line}"
        p.space_before = Pt(10)
        _style_paragraph(p, font_name, size_pt, color_hex)
    return box


def synthesize_title(prs, theme, bounds_in, data, resolver=None, canvas_idx=None, canvas_bg=None):
    """A cover/hero slide intentionally breaks from the "regular content" safe
    zone (big, vertically centered) — real cover slides do this too — so this
    one doesn't need the inferred content bounds the way the others do."""
    slide, idx, t, _ = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx, canvas_bg=canvas_bg)
    width, height = prs.slide_width, prs.slide_height

    title_text = data.get("title", "")
    title_width = Emu(int(width * 0.8))
    # The template's own cover size, capped so a long title cannot be broken
    # mid-word: at the 92pt these survey decks use, an uncapped base would do
    # exactly what iter18/iter25 caught on content titles.
    metrics = _metrics_for(resolver, t["major_font"])
    cover_pt = cap_size_to_longest_word(
        title_text, title_width, t.get("cover_pt") or DEFAULT_COVER_PT,
        metrics=metrics, bold=True)

    # Height from the text, not a fixed 22% of the slide. The subtitle used to
    # sit at a hard-coded 60% of the slide height, which was clear of a one-line
    # 40pt title and nothing else: the moment the cover took the template's own
    # size (88pt on the survey decks) the title wrapped to two lines and the
    # render showed «Итоги внедрения за квартал» printed straight through it.
    # Same lesson as iter35 — reserve from the real box, never from an offset
    # that happens to work at one size.
    title_top = int(height * 0.38)

    # …and capped again by the HEIGHT actually available below title_top. The
    # width cap above only stops mid-word breaks; nothing stopped the block
    # itself from running off the slide. The survey cover's 90pt title took
    # three lines and its box ran to 8.75in on a 7.5in slide, so «отчётности»
    # was cut by the bottom edge — and the subtitle's own safety clamp
    # (height*0.88) then placed it INSIDE the overflowing title, printing one
    # through the other. Reserve the subtitle's band first, then fit.
    subtitle = data.get("subtitle")
    sub_reserve = int(height * 0.12) + int(Inches(0.2)) if subtitle else 0
    avail_h = max(int(Inches(1.0)), int(height * 0.94) - title_top - sub_reserve)
    cover_pt = fit_font_size(title_text, title_width, Emu(avail_h), cover_pt,
                             min_size_pt=_MIN_COVER_PT, metrics=metrics).pt

    title_h_in = estimate_block_height_in(
        title_text, Emu(title_width).inches, cover_pt, metrics=metrics) + 0.15
    title_box = slide.shapes.add_textbox(
        Emu(int(width * 0.1)), Emu(title_top), title_width, Emu(int(Inches(title_h_in))))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.text = title_text
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    _style_paragraph(tf.paragraphs[0], t["major_font"], cover_pt, t["accent"], bold=True)

    if subtitle:
        # No clamp against the slide bottom here: the fit above guarantees the
        # title ends above the reserved band, and a clamp is exactly what put
        # the subtitle on top of the title when the title overflowed.
        sub_top = int(title_top + title_box.height + int(Inches(0.2)))
        sub_box = slide.shapes.add_textbox(
            Emu(int(width * 0.15)), Emu(sub_top), Emu(int(width * 0.7)), Emu(int(height * 0.12))
        )
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.text = subtitle
        stf.paragraphs[0].alignment = PP_ALIGN.CENTER
        _style_paragraph(stf.paragraphs[0], t["minor_font"], 20, t["text"])

    return idx


def synthesize_bullet_list(prs, theme, bounds_in, data, resolver=None, canvas_idx=None, canvas_bg=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx, canvas_bg=canvas_bg)
    width = Emu(b["right"] - b["left"])

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"],
                           size_pt=t.get("title_pt"), resolver=resolver)
    content_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.25))))

    bullets = data.get("bullets", [])
    est_in = estimate_block_height_in(
        [f"•  {line}" for line in bullets], Emu(width).inches, 18,
        metrics=_metrics_for(resolver, t["minor_font"]),
    ) + max(0, len(bullets) - 1) * 10 / 72  # space_before between items
    content_top = _centered_top(content_top, b["bottom"], est_in,
                                _decor_obstacles(slide, prs.slide_width, prs.slide_height))
    content_height = max(Emu(int(Inches(0.5))), Emu(b["bottom"] - content_top))

    _add_bulleted_textbox(
        slide, b["left"], content_top, width, content_height,
        bullets, t["minor_font"], 18, t["text"],
    )
    return idx


def synthesize_stats_kpi(prs, theme, bounds_in, data, resolver=None, canvas_idx=None, canvas_bg=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx, canvas_bg=canvas_bg)

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"],
                           size_pt=t.get("title_pt"), resolver=resolver)
    content_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.35))))
    # Number box (0.9") + label offset (1.0") + label box (0.8") — fixed layout.
    content_top = _centered_top(content_top, b["bottom"], 1.8,
                                _decor_obstacles(slide, prs.slide_width, prs.slide_height))

    stats = data.get("stats", [])
    if stats:
        total_width = Emu(b["right"] - b["left"])
        col_width = Emu(total_width // len(stats))
        for i, (num, label) in enumerate(stats):
            left = Emu(b["left"] + col_width * i)
            num_box = slide.shapes.add_textbox(left, content_top, col_width, Emu(int(Inches(0.9))))
            num_box.text_frame.word_wrap = True
            num_box.text_frame.text = str(num)
            _style_paragraph(num_box.text_frame.paragraphs[0], t["major_font"], 44, t["accent"], bold=True)

            label_top = Emu(content_top + Emu(int(Inches(1.0))))
            label_box = slide.shapes.add_textbox(left, label_top, col_width, Emu(int(Inches(0.8))))
            label_box.text_frame.word_wrap = True
            label_box.text_frame.text = str(label)
            _style_paragraph(label_box.text_frame.paragraphs[0], t["minor_font"], 16, t["text"])

    return idx


def synthesize_two_column_comparison(prs, theme, bounds_in, data, resolver=None, canvas_idx=None, canvas_bg=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx, canvas_bg=canvas_bg)
    margin = Emu(int(Inches(0.2)))

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"],
                           size_pt=t.get("title_pt"), resolver=resolver)
    col_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.3))))

    total_width = Emu(b["right"] - b["left"])
    col_width = Emu((total_width - margin) // 2)

    col_width_in = Emu(col_width).inches
    major_metrics = _metrics_for(resolver, t["major_font"])
    minor_metrics = _metrics_for(resolver, t["minor_font"])

    # A column is half the content width, and on a canvas whose bounds exclude a
    # photo it gets narrower still — narrow enough that the renderer breaks a
    # long word mid-letter. Seen on the T-Zh mono canvas: a 20pt heading came out
    # as "Разрозненнос/ть данных". Height-only fitting accepts that happily, so
    # cap both sizes on the LONGEST WORD, exactly what this helper exists for.
    headings = [data.get("left_heading", ""), data.get("right_heading", "")]
    points_all = list(data.get("left_points", [])) + list(data.get("right_points", []))
    heading_pt = cap_size_to_longest_word(headings, col_width, 20, metrics=major_metrics, bold=True)
    point_pt = cap_size_to_longest_word(points_all, col_width, 15, metrics=minor_metrics)

    est_in = max(
        estimate_block_height_in(heading, col_width_in, heading_pt, metrics=major_metrics)
        + estimate_block_height_in(points, col_width_in, point_pt, metrics=minor_metrics)
        for heading, points in (
            (data.get("left_heading", ""), data.get("left_points", [])),
            (data.get("right_heading", ""), data.get("right_points", [])),
        )
    )
    col_top = _centered_top(col_top, b["bottom"], est_in,
                            _decor_obstacles(slide, prs.slide_width, prs.slide_height))
    col_height = Emu(b["bottom"] - col_top)

    for left, heading, points in (
        (b["left"], data.get("left_heading", ""), data.get("left_points", [])),
        (Emu(b["left"] + col_width + margin), data.get("right_heading", ""), data.get("right_points", [])),
    ):
        box = slide.shapes.add_textbox(left, col_top, col_width, col_height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.text = heading
        _style_paragraph(tf.paragraphs[0], t["major_font"], heading_pt, t["accent"], bold=True)
        for point in points:
            p = tf.add_paragraph()
            p.text = point
            _style_paragraph(p, t["minor_font"], point_pt, t["text"])

    return idx


IMAGE_PLACEHOLDER_NAME = "img_placeholder"


def _add_image_placeholder(slide, left, top, width, height, caption, t):
    """A dashed frame with a caption — a SKELETON marking where a generated image
    will go. Images aren't generated yet (plan: skeleton first, for composition);
    this reserves the space so layout and image PLACEMENT can be judged. Named
    IMAGE_PLACEHOLDER_NAME so the evaluator can find placeholders on the slide."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR
    from pptx.oxml.ns import qn

    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shp.name = IMAGE_PLACEHOLDER_NAME
    shp.shadow.inherit = False
    shp.fill.background()  # transparent — a frame, not a filled box
    shp.line.color.rgb = _hex_to_rgb(t["accent"]) or RGBColor(0x88, 0x88, 0x88)
    shp.line.width = Pt(1.5)
    ln = shp.line._get_or_add_ln()  # dashed border reads as "reserved / to-be-filled"
    ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))

    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.text = "ИЗОБРАЖЕНИЕ"
    _style_paragraph(tf.paragraphs[0], t["minor_font"], 12, t["accent"], bold=True)
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    cap = tf.add_paragraph()
    cap.text = str(caption)
    _style_paragraph(cap, t["minor_font"], 14, t["text"])
    cap.alignment = PP_ALIGN.CENTER
    return shp


def synthesize_image_caption(prs, theme, bounds_in, data, resolver=None, canvas_idx=None, canvas_bg=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx, canvas_bg=canvas_bg)
    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"],
                           size_pt=t.get("title_pt"), resolver=resolver)
    # Reserve a two-line title's worth of clearance: the frame is large, so any
    # overlap with a title that wrapped to a second line is glaring, and the
    # estimate can under-count lines (renderer wraps ~2-4% earlier).
    #
    # But that reserve has to be EARNED. Measured on the T-Zh study canvas, a
    # title that provably cannot wrap («Платформа в работе», 32pt, 7.20in box)
    # still got the floor: 1.65in reserved against a 0.83in title, i.e. 0.47in
    # of dead band between the heading and the frame — visible on the render as
    # a hole under the title. Where a wrap is impossible with 10% to spare,
    # reserve exactly the title.
    metrics = _metrics_for(resolver, t["major_font"])
    fitted_pt = cap_size_to_longest_word(
        data.get("title", ""), Emu(b["right"] - b["left"]),
        t.get("title_pt") or DEFAULT_TITLE_PT, metrics=metrics, bold=True)
    if _title_fits_one_line(data.get("title", ""), b["right"] - b["left"], fitted_pt, metrics):
        clearance = int(title_box.height) + int(Inches(0.35))
    else:
        clearance = max(int(title_box.height), int(Inches(1.3))) + int(Inches(0.35))
    below_top = int(title_box.top) + clearance
    width = Emu(b["right"] - b["left"])

    # Put the frame wherever there is actually room. Reserving a fixed minimum
    # height below the title used to run it straight off the page: on a 5.62"
    # T-Zh slide whose title sits low, the frame started at 5.38" and ended at
    # 6.98" — over three quarters of it off-slide, and the render showed a
    # sliver of dashed border at the bottom edge. Some templates put the title
    # low by design, so the band ABOVE it can be the larger one; pick whichever
    # is bigger and never cross the content bounds.
    gap = int(Inches(0.35))
    below = int(b["bottom"]) - below_top
    above = (int(title_box.top) - gap) - int(b["top"])
    if above > below:
        top, height = int(b["top"]), above
    else:
        top, height = below_top, below

    # Never force a minimum height: doing that just moves the problem from
    # "frame hangs off the slide" to "frame sits on top of the title" — both
    # tried on the same T-Zh slide, both visibly wrong. The frame gets exactly
    # the band that is free, however short that is; below MIN_FRAME_HEIGHT
    # there is nothing worth drawing and the slide keeps just its title.
    height = min(height, int(b["bottom"]) - top)
    if height < MIN_FRAME_HEIGHT:
        return idx

    caption = data.get("image") or data.get("caption") or "иллюстрация по теме слайда"
    _add_image_placeholder(slide, b["left"], Emu(top), width, Emu(height), caption, t)
    return idx


SYNTHESIZERS = {
    "title": synthesize_title,
    "section_divider": synthesize_title,
    "closing": synthesize_title,
    "bullet_list": synthesize_bullet_list,
    "stats_kpi": synthesize_stats_kpi,
    "two_column_comparison": synthesize_two_column_comparison,
    "image_caption": synthesize_image_caption,
}
