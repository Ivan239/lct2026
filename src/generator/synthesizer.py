"""Builds slides from scratch, for archetypes the uploaded template simply has no
slide for at all — using the template's own theme (fonts/colors) and its own
real content geometry (see layout_bounds.py) so the result still looks like it
belongs, rather than colliding with master-level decorations we can't even see
(a logo/footer baked into the slide master isn't in slide.shapes) or using
fixed percentages that happened to work for one file and not the next."""

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

from generator.slide_kit import clone_slide, content_text_shapes
from generator.text_fit import estimate_block_height_in

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




def _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=None):
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
        palette = _theme_palette(theme)
        removed = content_text_shapes(slide)
        placed = [s for s in removed if s.left is not None and s.top is not None and s.width and s.height]
        if placed:
            bounds = {
                "left": Emu(min(s.left for s in placed)),
                "top": Emu(min(s.top for s in placed)),
                "right": Emu(max(s.left + s.width for s in placed)),
                "bottom": Emu(max(s.top + s.height for s in placed)),
            }
        else:
            bounds = _resolve_bounds(prs, bounds_in)
        for shape in removed:
            shape._element.getparent().remove(shape._element)
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


def _centered_top(content_top, bounds_bottom, est_height_in):
    """Top position that vertically centers a content block of the estimated
    height between content_top and the bottom bound. Synthesized slides have no
    designer to fill the lower half — a short block left hanging right under
    the title reads as an accidentally half-empty slide (same defect class as
    generator._center_if_underfilled, but for slides we build ourselves).
    The title stays where native slides put theirs; only the body block moves."""
    avail = int(bounds_bottom) - int(content_top)
    est = int(Inches(max(0.0, est_height_in)))
    if est >= avail:
        return content_top
    return Emu(int(content_top) + (avail - est) // 2)


def _add_title(slide, bounds, text, font_name, color_hex, size_pt=32, resolver=None):
    """Sized to its actual estimated line count rather than a fixed height, so
    content below it is positioned after wherever the title really ends."""
    width = Emu(bounds["right"] - bounds["left"])
    metrics = _metrics_for(resolver, font_name)
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


def synthesize_title(prs, theme, bounds_in, data, resolver=None, canvas_idx=None):
    """A cover/hero slide intentionally breaks from the "regular content" safe
    zone (big, vertically centered) — real cover slides do this too — so this
    one doesn't need the inferred content bounds the way the others do."""
    slide, idx, t, _ = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx)
    width, height = prs.slide_width, prs.slide_height

    title_box = slide.shapes.add_textbox(
        Emu(int(width * 0.1)), Emu(int(height * 0.38)), Emu(int(width * 0.8)), Emu(int(height * 0.22))
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.text = data.get("title", "")
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    _style_paragraph(tf.paragraphs[0], t["major_font"], 40, t["accent"], bold=True)

    subtitle = data.get("subtitle")
    if subtitle:
        sub_box = slide.shapes.add_textbox(
            Emu(int(width * 0.15)), Emu(int(height * 0.6)), Emu(int(width * 0.7)), Emu(int(height * 0.12))
        )
        stf = sub_box.text_frame
        stf.word_wrap = True
        stf.text = subtitle
        stf.paragraphs[0].alignment = PP_ALIGN.CENTER
        _style_paragraph(stf.paragraphs[0], t["minor_font"], 20, t["text"])

    return idx


def synthesize_bullet_list(prs, theme, bounds_in, data, resolver=None, canvas_idx=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx)
    width = Emu(b["right"] - b["left"])

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"], resolver=resolver)
    content_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.25))))

    bullets = data.get("bullets", [])
    est_in = estimate_block_height_in(
        [f"•  {line}" for line in bullets], Emu(width).inches, 18,
        metrics=_metrics_for(resolver, t["minor_font"]),
    ) + max(0, len(bullets) - 1) * 10 / 72  # space_before between items
    content_top = _centered_top(content_top, b["bottom"], est_in)
    content_height = max(Emu(int(Inches(0.5))), Emu(b["bottom"] - content_top))

    _add_bulleted_textbox(
        slide, b["left"], content_top, width, content_height,
        bullets, t["minor_font"], 18, t["text"],
    )
    return idx


def synthesize_stats_kpi(prs, theme, bounds_in, data, resolver=None, canvas_idx=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx)

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"], resolver=resolver)
    content_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.35))))
    # Number box (0.9") + label offset (1.0") + label box (0.8") — fixed layout.
    content_top = _centered_top(content_top, b["bottom"], 1.8)

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


def synthesize_two_column_comparison(prs, theme, bounds_in, data, resolver=None, canvas_idx=None):
    slide, idx, t, b = _prepare_blank_slide(prs, theme, bounds_in, canvas_idx=canvas_idx)
    margin = Emu(int(Inches(0.2)))

    title_box = _add_title(slide, b, data.get("title", ""), t["major_font"], t["accent"], resolver=resolver)
    col_top = Emu(title_box.top + title_box.height + Emu(int(Inches(0.3))))

    total_width = Emu(b["right"] - b["left"])
    col_width = Emu((total_width - margin) // 2)

    col_width_in = Emu(col_width).inches
    major_metrics = _metrics_for(resolver, t["major_font"])
    minor_metrics = _metrics_for(resolver, t["minor_font"])
    est_in = max(
        estimate_block_height_in(heading, col_width_in, 20, metrics=major_metrics)
        + estimate_block_height_in(points, col_width_in, 15, metrics=minor_metrics)
        for heading, points in (
            (data.get("left_heading", ""), data.get("left_points", [])),
            (data.get("right_heading", ""), data.get("right_points", [])),
        )
    )
    col_top = _centered_top(col_top, b["bottom"], est_in)
    col_height = Emu(b["bottom"] - col_top)

    for left, heading, points in (
        (b["left"], data.get("left_heading", ""), data.get("left_points", [])),
        (Emu(b["left"] + col_width + margin), data.get("right_heading", ""), data.get("right_points", [])),
    ):
        box = slide.shapes.add_textbox(left, col_top, col_width, col_height)
        tf = box.text_frame
        tf.word_wrap = True
        tf.text = heading
        _style_paragraph(tf.paragraphs[0], t["major_font"], 20, t["accent"], bold=True)
        for point in points:
            p = tf.add_paragraph()
            p.text = point
            _style_paragraph(p, t["minor_font"], 15, t["text"])

    return idx


SYNTHESIZERS = {
    "title": synthesize_title,
    "section_divider": synthesize_title,
    "closing": synthesize_title,
    "bullet_list": synthesize_bullet_list,
    "stats_kpi": synthesize_stats_kpi,
    "two_column_comparison": synthesize_two_column_comparison,
}
