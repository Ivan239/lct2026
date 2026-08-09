"""Zero-token half of the rubric scorer.

The rubric's technical/typographic criteria are exactly the ones this project's
whole philosophy says to enforce in code, not hand to a model: overflow, out-of-
bounds boxes, orphan lines, title length, duplicate slides, pacing. So they're
scored here from the real .pptx geometry and the same font metrics the generator
fits against — no LLM, no tokens, and the numbers are reproducible.

Deliberately NOT here: rectangle-overlap detection. qa/geometry.py documents why
(designers overlap generous text rectangles by intent — 51 false hits on a
pristine 69-slide deck), so "no overlaps" (9.2) is left to the vision judge,
which sees what actually renders.

Each check returns {"score": 1..5 or None, "detail": str}. Rates map to scores
via _rate_to_score so "half the slides are affected" lands mid-scale, not at the
floor.
"""

import re
from collections import Counter

from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Emu, Inches

from fonts.metrics import FontResolver
from generator.generator import _pick_body_shape, _pick_title_shape
from generator.text_fit import (
    estimate_block_height_in,
    horizontal_margins_in,
    vertical_insets_emu,
)
from generator.slide_kit import is_chrome_shape, slide_height
from qa.geometry import SPARSE_EXEMPT_ROLES, _is_content_shape
from template_parser.parser import extract_theme

OVERFLOW_TOLERANCE = 1.05          # same slack qa/geometry uses before "it overflows"
MIN_READABLE_PT = 11.0             # below this, body text reads as fine print…
# …on a 16:9 slide of the standard 13.33in width. A point is absolute on paper,
# but a deck is projected to fill a screen, so what decides readability is the
# size RELATIVE to the canvas: 10pt on a 10in-wide slide renders exactly as
# large as 13.3pt on a 13.33in one. All three T-Zh templates are authored at
# 10x5.62in, and the flat threshold reported NINE boxes of pristine T-Zh mono
# as fine print — every one of them the deck's own 10pt body copy.
# In EMU, exactly: the inch value 13.333 is SHORT of the real 12192000 EMU, and
# the resulting ratio of 1.000025 pushed the threshold to 11.0003pt — enough to
# reclassify 29 boxes set at exactly 11pt in survey-69 as fine print.
REFERENCE_SLIDE_WIDTH_EMU = 12192000
TITLE_MAX_CHARS = 60
TITLE_MAX_WORDS = 10
WALL_OF_TEXT_CHARS = 600           # one content box past this is a wall of text
NOISE_SHAPE_COUNT = 12             # more content boxes than this reads as clutter
EDGE_SAFE_FRACTION = 0.02          # inside 2% of a slide edge = not in the safe area
OOB_TOLERANCE_FRACTION = 0.02      # past the edge by >2% = genuinely clipped
WIDOW_MAX_FRACTION = 0.40          # lone last-line word under 40% of the line = widow


SUBSTANTIVE_PICTURE_FRACTION = 0.03  # bigger than a bullet-marker icon


def deck_media(pptx_path):
    """What visual media the deck actually contains — so image/infographics
    criteria can be forced to N/A when there's nothing to judge, instead of
    trusting the model (which scores 'no images' as 1 about half the time).
    Bullet-marker icons are excluded by the area threshold: they're not the
    'изображения' the rubric's section 4 is about.

    Also counts image PLACEHOLDERS (dashed skeleton frames the generator draws
    where an image will go): a deck with placeholders has no real picture yet,
    but its image PLACEMENT/composition is real and should be scored."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    from generator.synthesizer import IMAGE_PLACEHOLDER_NAME

    prs = Presentation(pptx_path)
    area = (prs.slide_width or 1) * (prs.slide_height or 1)
    pictures = charts = placeholders = 0
    for slide in prs.slides:
        for s in slide.shapes:
            if s.name == IMAGE_PLACEHOLDER_NAME:
                placeholders += 1
                continue
            if getattr(s, "has_chart", False):
                charts += 1
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if (s.width or 0) * (s.height or 0) / area > SUBSTANTIVE_PICTURE_FRACTION:
                    pictures += 1
    return {"substantive_pictures": pictures, "charts": charts, "placeholders": placeholders}


def _rate_to_score(rate):
    """Fraction-of-items-affected (0..1) -> 1..5. 0 bad -> 5, everything bad -> 1."""
    return max(1, min(5, round(5 - 4 * rate)))


def _cv_to_score(cv):
    """Coefficient of variation of a per-slide quantity -> evenness score."""
    if cv <= 0.35:
        return 5
    if cv <= 0.55:
        return 4
    if cv <= 0.80:
        return 3
    if cv <= 1.10:
        return 2
    return 1


def _norm_text(s):
    return re.sub(r"\s+", " ", s or "").strip().lower()


def _content_shapes(slide):
    """The slide's CONTENT boxes — chrome excluded.

    Chrome is the designer's furniture, not our text, and judging it as body
    copy makes the harness wrong about a deck that is fine. Measured on the
    T-Zh universal deck: 17 of 35 "content" boxes counted as unreadable fine
    print, every one of them the template's own 8pt «2025», running topic or
    page number, and zero real overflows. Readability came out at 3 for a deck
    with nothing wrong with it.

    Sharpened by an earlier fix of ours: until iter31 that furniture was blanked
    on the way out, so it never reached this count. Repairing the product made
    the measurement look worse."""
    height = slide_height(slide)
    return [s for s in slide.shapes
            if _is_content_shape(s) and not is_chrome_shape(s, height)]


_MARGIN_BUCKET_EMU = int(Emu(int(Inches(0.05))))


def _designed_margins(slides, W, H):
    """The margins the DECK itself keeps, as fractions of the slide.

    A fixed "2% of the slide" safe area penalises a designer who chose a tighter
    one: the T-Zh mono template sets its content at 0.19-0.20in on a 10in slide
    — 1.9% — so nine boxes were reported outside the safe area for doing exactly
    what the template does everywhere. Measured across the corpus, every deck's
    smallest offset equals its 10th percentile, i.e. boxes sit ON a designed
    margin rather than scattering towards the edge (mono 1.9%, the others
    2.5-3.9%).

    So take the MODE of the offsets, bucketed to 0.05in: a margin the deck
    repeats is a decision, while a single box shoved to the edge stays a lone
    value and is still caught."""
    sides = {"l": [], "t": [], "r": [], "b": []}
    for slide in slides:
        for s in _content_shapes(slide):
            if s.left is None or s.top is None or not s.width or not s.height:
                continue
            sides["l"].append(s.left)
            sides["t"].append(s.top)
            sides["r"].append(W - (s.left + s.width))
            sides["b"].append(H - (s.top + s.height))
    out = {}
    for side, values in sides.items():
        if not values:
            out[side] = None
            continue
        buckets = Counter(max(0, v) // _MARGIN_BUCKET_EMU for v in values)
        out[side] = (buckets.most_common(1)[0][0] * _MARGIN_BUCKET_EMU) / (
            W if side in ("l", "r") else H)
    return out


def _shape_runs(shape):
    return [run for p in shape.text_frame.paragraphs for run in p.runs]


def _shape_max_size(shape):
    sizes = [r.font.size.pt for r in _shape_runs(shape) if r.font.size and r.text.strip()]
    return max(sizes) if sizes else None


def _text_band_emu(shape):
    """(top, bottom) of the shape's TEXT, not of its box.

    A box is routinely taller than what it holds, and the difference is not a
    defect: survey-69 slide 6 runs a box to 7.83in on a 7.5in slide while its
    three paragraphs end at 5.11in. Measuring the box would call the designer's
    own slide broken — the mistake CLAUDE.md records three times over.

    Only the anchor decides where the text sits inside the box: top-anchored
    text starts at the top inset, bottom-anchored ends at the bottom one,
    middle-anchored is centred. Without metrics the estimate is unreliable, so
    the caller is told nothing rather than something wrong."""
    top_inset, bottom_inset = vertical_insets_emu(shape)
    box_top, box_bottom = int(shape.top) + top_inset, int(shape.top + shape.height) - bottom_inset
    return box_top, box_bottom


# A text block buried under another one. Calibrated, not guessed: across the 217
# slides of the real corpus the designers' worst overlap is 0.51 (survey-69 tucks
# «*new question in survey» right under its question, on purpose), while the real
# defect of iter59 — a subtitle printed through the cover title — scores 1.00,
# the smaller block sitting entirely inside the larger one's text band.
TEXT_COLLISION_FRACTION = 0.7


def _text_rect(shape, metrics_for):
    """The rectangle the shape's TEXT occupies: box width (text wraps inside it)
    by the estimated text band. None when the estimate is not trustworthy."""
    size_pt = _shape_max_size(shape)
    font_name = next((r.font.name for r in _shape_runs(shape) if r.font.name), None)
    metrics = metrics_for(font_name) if font_name else None
    texts = [t for t in ("".join(r.text for r in p.runs)
                         for p in shape.text_frame.paragraphs) if t.strip()]
    if not size_pt or metrics is None or not texts or not shape.width:
        return None
    if shape.left is None or shape.top is None or not shape.height:
        return None
    line_spacing = next(
        (p.line_spacing for p in shape.text_frame.paragraphs if p.line_spacing is not None), None)
    est = int(Inches(estimate_block_height_in(
        texts, Emu(shape.width).inches, size_pt, metrics=metrics,
        line_spacing=line_spacing, margins_in=horizontal_margins_in(shape))))
    top, _ = _text_band_emu(shape)
    return int(shape.left), top, int(shape.left + shape.width), top + est


def _text_collisions(content, metrics_for):
    """Pairs of content blocks whose TEXT lands on top of other text.

    Not box rectangles: CLAUDE.md records that those intersect by design (51
    hits on the pristine 69-slide original). Text bands are a different measure —
    zero hits on four of the five real templates, and on the fifth only the
    footnotes the designer tucked under a heading on purpose."""
    rects = [r for r in (_text_rect(s, metrics_for) for s in content) if r]
    hits = 0
    for i, a in enumerate(rects):
        for b in rects[i + 1:]:
            dx = min(a[2], b[2]) - max(a[0], b[0])
            dy = min(a[3], b[3]) - max(a[1], b[1])
            if dx <= 0 or dy <= 0:
                continue
            smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
            if smaller and dx * dy / smaller >= TEXT_COLLISION_FRACTION:
                hits += 1
    return hits


def _text_past_edge(shape, W, H, metrics_for):
    """EMU by which the shape's text crosses a slide edge, 0 when it doesn't.

    Left/right come from the box: text wraps inside it, so the box edge IS the
    text edge. Top/bottom come from the estimated text height — see
    _text_band_emu for why the box would lie there."""
    if shape.left is None or shape.top is None or not shape.width or not shape.height:
        return 0
    over = max(0, -int(shape.left), int(shape.left + shape.width) - int(W))

    size_pt = _shape_max_size(shape)
    font_name = next((r.font.name for r in _shape_runs(shape) if r.font.name), None)
    metrics = metrics_for(font_name) if font_name else None
    texts = [t for t in ("".join(r.text for r in p.runs)
                         for p in shape.text_frame.paragraphs) if t.strip()]
    if not size_pt or metrics is None or not texts:
        return over  # blind on the vertical axis: say so by measuring nothing

    line_spacing = next(
        (p.line_spacing for p in shape.text_frame.paragraphs if p.line_spacing is not None), None)
    est_emu = int(Inches(estimate_block_height_in(
        texts, Emu(shape.width).inches, size_pt, metrics=metrics,
        line_spacing=line_spacing, margins_in=horizontal_margins_in(shape))))
    band_top, band_bottom = _text_band_emu(shape)
    anchor = shape.text_frame.vertical_anchor
    if anchor == MSO_ANCHOR.BOTTOM:
        text_top, text_bottom = band_bottom - est_emu, band_bottom
    elif anchor == MSO_ANCHOR.MIDDLE:
        mid = (band_top + band_bottom) // 2
        text_top, text_bottom = mid - est_emu // 2, mid + est_emu // 2
    else:
        text_top, text_bottom = band_top, band_top + est_emu
    return max(over, -text_top, text_bottom - int(H), 0)


def _overflows(shape, metrics_for):
    """True if the shape's text is estimated to exceed its box (would be clipped
    or auto-shrunk by the renderer). Mirrors qa.geometry.enforce_text_fits but
    read-only — no mutation, we're only measuring."""
    if not shape.width or not shape.height:
        return False
    size_pt = _shape_max_size(shape)
    if not size_pt:
        return False
    font_name = next((r.font.name for r in _shape_runs(shape) if r.font.name), None)
    metrics = metrics_for(font_name) if font_name else None
    if metrics is None:
        return False
    texts = ["".join(r.text for r in p.runs) for p in shape.text_frame.paragraphs]
    while texts and not texts[-1].strip():
        texts.pop()
    if not any(t.strip() for t in texts):
        return False
    line_spacing = next(
        (p.line_spacing for p in shape.text_frame.paragraphs if p.line_spacing is not None),
        None,
    )
    top_inset, bottom_inset = vertical_insets_emu(shape)
    usable_h_in = Emu(max(0, shape.height - top_inset - bottom_inset)).inches
    est = estimate_block_height_in(
        texts, Emu(shape.width).inches, size_pt,
        metrics=metrics, line_spacing=line_spacing,
        margins_in=horizontal_margins_in(shape),
    )
    return est > usable_h_in * OVERFLOW_TOLERANCE


NBSP = " "


def _tokens(raw):
    """Split into wrap tokens on breakable whitespace only — a non-breaking space
    (U+00A0) keeps its two words as ONE token, matching how the renderer wraps
    text glued by the widow guard (two_phase._guard_widow). No-op for text with
    no NBSP, so every existing deck tokenises exactly as raw.split() did."""
    return [t for t in re.split(r"[^\S ]+", raw) if t]


def _wrap_words(text, width_in, size_pt, metrics, margins_in):
    """Greedy word-wrap into visual lines (list of token-lists), matching the
    renderer. Used to spot orphan last lines."""
    budget_pt = max(1.0, (width_in - (margins_in or 0)) * 72)
    space_pt = metrics.text_width_pt(" ", size_pt)
    lines = []
    for raw in text.split("\n"):
        words = _tokens(raw)
        if not words:
            continue
        cur, cur_w = [], 0.0
        for w in words:
            # NBSP renders as a space width; measure it as one so the glued
            # token's width is right even if the font lacks a U+00A0 glyph.
            ww = metrics.text_width_pt(w.replace(NBSP, " "), size_pt)
            step = ww if not cur else space_pt + ww
            if cur and cur_w + step > budget_pt:
                lines.append((cur, cur_w))
                cur, cur_w = [w], ww
            else:
                cur.append(w)
                cur_w += step
        if cur:
            lines.append((cur, cur_w))
    return budget_pt, lines


def _is_widow(text, width_in, size_pt, metrics, margins_in):
    """A widow/orphan is a wrapped paragraph whose LAST line is a single short
    word — a lone word using less than WIDOW_MAX_FRACTION of the line is what
    reads as a dangling straggler; one long word filling most of its line is
    just a normal wrap, not a defect. A single token that is itself an NBSP-glued
    pair is TWO visible words, so it isn't a widow."""
    budget_pt, lines = _wrap_words(text, width_in, size_pt, metrics, margins_in)
    if len(lines) < 2:
        return False, len(lines)
    last_words, last_w = lines[-1]
    lone = len(last_words) == 1 and NBSP not in last_words[0]
    return (lone and last_w < WIDOW_MAX_FRACTION * budget_pt), len(lines)


def evaluate(pptx_path, slide_roles=None):
    """slide_roles: optional {0-based position: role_str} so completeness can
    check the deck actually ends on a closing/summary. Returns
    {criterion_id: {"score": .., "detail": ..}} for the deterministic criteria."""
    prs = Presentation(pptx_path)
    resolver = FontResolver(pptx_path, extract_theme(pptx_path))
    metrics_for = resolver.metrics_for
    W, H = prs.slide_width, prs.slide_height

    slides = list(prs.slides)
    n = len(slides)
    per_slide = []  # collected features
    all_titles = []
    all_slide_texts = []
    all_line_sets = []  # per-slide set of normalized content lines, for near-dup

    min_readable_pt = MIN_READABLE_PT * W / REFERENCE_SLIDE_WIDTH_EMU
    designed = _designed_margins(slides, W, H)
    for slide in slides:
        content = _content_shapes(slide)
        title_shape = _pick_title_shape(slide, set())
        title = _norm_text(title_shape.text_frame.text) if title_shape else ""
        all_titles.append(title)
        slide_text = _norm_text(" ".join(s.text_frame.text for s in content))
        all_slide_texts.append(slide_text)
        lines = set()
        for s in content:
            for p in s.text_frame.paragraphs:
                ln = _norm_text("".join(r.text for r in p.runs))
                if len(ln) > 3:  # skip bare numbers / one-char spacers
                    lines.add(ln)
        all_line_sets.append(lines)

        overflow = sum(1 for s in content if _overflows(s, metrics_for))
        collide = _text_collisions(content, metrics_for)
        tiny = sum(1 for s in content if (_shape_max_size(s) or 99) < min_readable_pt)
        oob = 0
        near_edge = 0
        for s in content:
            if s.left is None or s.top is None or not s.width or not s.height:
                continue
            l, t, r, b = s.left, s.top, s.left + s.width, s.top + s.height
            if _text_past_edge(s, W, H, metrics_for) > OOB_TOLERANCE_FRACTION * min(W, H):
                oob += 1
            # Tighter than BOTH the generic safe area and the deck's own
            # designed margin — a box level with the rest of the deck is not a
            # margin violation, however tight that margin is.
            def _tight(offset, limit, designed):
                bound = limit if designed is None else min(limit, designed)
                return offset < bound
            if (_tight(l / W, EDGE_SAFE_FRACTION, designed["l"])
                    or _tight(t / H, EDGE_SAFE_FRACTION, designed["t"])
                    or _tight((W - r) / W, EDGE_SAFE_FRACTION, designed["r"])
                    or _tight((H - b) / H, EDGE_SAFE_FRACTION, designed["b"])):
                near_edge += 1

        # orphan last lines across content paragraphs — the TITLE excluded.
        # A widow is a body-copy defect. A display title wrapping to two lines
        # with one word on the second is normal typography, and these templates
        # do it themselves: T-Zh mono's own title slide reads «Заголовок /
        # слайда», and our «Платформа / Поток» at 68pt was the single orphan the
        # whole corpus produced — looked at on the render, it is the template's
        # own treatment, not a defect.
        title_id = title_shape.shape_id if title_shape is not None else None
        orphan = 0
        multiline = 0
        for s in content:
            if not s.width or s.shape_id == title_id:
                continue
            fn = next((r.font.name for r in _shape_runs(s) if r.font.name), None)
            m = metrics_for(fn) if fn else None
            if m is None:
                continue
            for p in s.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs).strip()
                if not txt:
                    continue
                # The PARAGRAPH's own size, not the shape's maximum. A box can
                # legitimately hold two sizes — a display figure over its
                # caption — and measuring the caption's wrapping at the figure's
                # size invents lines that do not exist: a 18pt caption that
                # renders on one line was reported as a two-line widow because
                # the number above it is 24pt.
                sizes = [r.font.size.pt for r in p.runs if r.font.size and r.text.strip()]
                size_pt = max(sizes) if sizes else _shape_max_size(s)
                if not size_pt:
                    continue
                widow, nlines = _is_widow(txt, Emu(s.width).inches, size_pt, m,
                                          horizontal_margins_in(s))
                if nlines >= 2:
                    multiline += 1
                    if widow:
                        orphan += 1

        per_slide.append({
            "n_content": len(content),
            "chars": len(slide_text),
            "overflow": overflow,
            "collide": collide,
            "tiny": tiny,
            "oob": oob,
            "near_edge": near_edge,
            "orphan": orphan,
            "multiline": multiline,
            "has_title": bool(title),
            "wall": len(slide_text) > WALL_OF_TEXT_CHARS,
        })

    total_content = sum(p["n_content"] for p in per_slide) or 1
    n_titled = sum(1 for t in all_titles if t) or 1

    def rate_over_content(key):
        return sum(p[key] for p in per_slide) / total_content

    scores = {}

    # 1.1 readability — fine print (a rate: it degrades a deck gradually) and
    # text printed through other text (categorical: one such slide cannot be
    # shown, so it caps the score the way 9.1 does).
    #
    # The old formula counted _overflows instead, and that is not a defect
    # measure: 32 of the 79 boxes of the pristine survey-31 "overflow", 25 of
    # them marked noAutofit — the designer keeps the frame smaller than the text
    # on purpose and the renderer simply draws past it. What the spill can
    # actually do is leave the slide (9.1 measures that since iter60) or land on
    # other text, which is what is counted here.
    tiny_total = sum(p["tiny"] for p in per_slide)
    collided = sum(p["collide"] for p in per_slide)
    slides_collided = sum(1 for p in per_slide if p["collide"])
    base = _rate_to_score(tiny_total / total_content)
    scores["1.1"] = {
        "score": min(base, 2 if slides_collided == 1 else 1) if collided else base,
        "detail": (f"{tiny_total} из {total_content} контентных боксов мелким шрифтом"
                   + (f"; {collided} блоков текста поверх другого текста "
                      f"на {slides_collided} слайдах" if collided else "")),
    }
    # dop_wrap — same overflow signal, framed as wrapping correctness
    of = sum(p["overflow"] for p in per_slide)
    scores["dop_wrap"] = {
        "score": _rate_to_score(of / total_content),
        "detail": f"{of} боксов с некорректным переносом (текст не влезает)",
    }
    # dop_orphans
    total_ml = sum(p["multiline"] for p in per_slide)
    orph = sum(p["orphan"] for p in per_slide)
    scores["dop_orphans"] = {
        "score": _rate_to_score(orph / total_ml) if total_ml else 5,
        "detail": f"{orph} висячих строк из {total_ml} многострочных абзацев",
    }
    # 2.1 title length
    long_titles = sum(
        1 for t in all_titles if t and (len(t) > TITLE_MAX_CHARS or len(t.split()) > TITLE_MAX_WORDS)
    )
    scores["2.1"] = {
        "score": _rate_to_score(long_titles / n_titled),
        "detail": f"{long_titles} заголовков длиннее {TITLE_MAX_CHARS} симв./{TITLE_MAX_WORDS} слов",
    }
    # 2.3 title uniqueness
    titles = [t for t in all_titles if t]
    dup_titles = len(titles) - len(set(titles))
    scores["2.3"] = {
        "score": _rate_to_score(dup_titles / len(titles)) if titles else 5,
        "detail": f"{dup_titles} повторяющихся заголовков из {len(titles)}",
    }
    # 3.1 text volume
    walls = sum(1 for p in per_slide if p["wall"])
    scores["3.1"] = {
        "score": _rate_to_score(walls / n) if n else 5,
        "detail": f"{walls} из {n} слайдов со «стеной текста» (>{WALL_OF_TEXT_CHARS} симв.)",
    }
    # 9.1 out-of-bounds — by SEVERITY, not by rate. A rate is right for defects
    # that degrade a deck gradually; text running off the slide is categorical.
    # The survey cover ran its title to 8.75in on a 7.5in slide with its last
    # line cut off, and 9.1 scored 5/5 — one bad box among fifteen good ones is
    # a 6% rate, which rounds to perfect. The deck reported 98.8 while its cover
    # was sliced in half, for three iterations running.
    off_slide = sum(p["oob"] for p in per_slide)
    slides_hit = sum(1 for p in per_slide if p["oob"])
    scores["9.1"] = {
        "score": 5 if not off_slide else (2 if slides_hit == 1 else 1),
        "detail": (f"{off_slide} боксов с текстом за границей слайда "
                   f"на {slides_hit} слайдах" if off_slide
                   else "текст нигде не выходит за границы слайда"),
    }
    # dop_safe_margins
    scores["dop_safe_margins"] = {
        "score": _rate_to_score(rate_over_content("near_edge")),
        "detail": (f"{sum(p['near_edge'] for p in per_slide)} боксов ближе к краю, "
                   f"чем {EDGE_SAFE_FRACTION:.0%} и чем собственное поле деки"),
    }
    # dop_noise
    noisy = sum(1 for p in per_slide if p["n_content"] > NOISE_SHAPE_COUNT)
    scores["dop_noise"] = {
        "score": _rate_to_score(noisy / n) if n else 5,
        "detail": f"{noisy} из {n} слайдов с >{NOISE_SHAPE_COUNT} контентных боксов",
    }
    # dop_no_dup_slides — exact text dups AND semantic near-dups. The near-dup
    # arm exists because the LLM judge caught two stat slides sharing all three
    # metric LABELS (different numbers, so not byte-identical) that this check
    # used to wave through — a real defect on iteration #1's deck (slides 5/7).
    nonempty = [t for t in all_slide_texts if t]
    dup_slides = len(nonempty) - len(set(nonempty))
    near_dup = _near_duplicate_slides(all_line_sets)
    affected = min(n, dup_slides + near_dup)
    scores["dop_no_dup_slides"] = {
        "score": _rate_to_score(affected / n) if n else 5,
        "detail": f"{dup_slides} точных + {near_dup} смысловых дублей слайдов",
    }
    # dop_text_split — overflow means content should have been split further
    scores["dop_text_split"] = {
        "score": _rate_to_score((of + walls) / (n or 1)),
        "detail": f"{of} переполнений + {walls} стен текста как признак плохого разбиения",
    }
    # dop_distribution / dop_pacing — evenness across CONTENT slides only. Title,
    # divider and closing are sparse BY DESIGN (a divider is one phrase); counting
    # them as "uneven" punishes a deck for having structure — the same reason
    # qa.geometry.find_sparse_slides exempts these roles. Without roles, fall back
    # to all slides. (Also: char/shape counts are now read from the SAME slides,
    # not filtered independently and zipped — that could misalign the pairs.)
    if slide_roles:
        content_slides = [p for i, p in enumerate(per_slide)
                          if slide_roles.get(i) not in SPARSE_EXEMPT_ROLES]
    else:
        content_slides = per_slide
    char_counts = [p["chars"] for p in content_slides if p["chars"] > 0]
    scores["dop_distribution"] = _evenness_score(char_counts, "символов контентных слайдов")
    paced = [p["chars"] + p["n_content"] * 40 for p in content_slides if p["chars"] > 0]
    scores["dop_pacing"] = _evenness_score(paced or char_counts, "объёма контентных слайдов")
    # 6.3 completeness
    if slide_roles:
        last_role = slide_roles.get(n - 1)
        closing_anywhere = any(r in ("closing",) for r in slide_roles.values())
        if last_role == "closing":
            scores["6.3"] = {"score": 5, "detail": "дек завершается closing-слайдом"}
        elif closing_anywhere:
            scores["6.3"] = {"score": 4, "detail": "closing есть, но не в конце"}
        else:
            scores["6.3"] = {"score": 2, "detail": "нет завершающего/итогового слайда"}
    else:
        scores["6.3"] = {"score": None, "detail": "роли слайдов не переданы"}

    return scores


NEAR_DUP_SHARED_LINES = 3   # this many identical content lines shared = a dup
NEAR_DUP_JACCARD = 0.6      # ...or this much line-set overlap for shorter slides


def _near_duplicate_slides(line_sets):
    """Count slides that are semantic near-duplicates of an earlier slide:
    they share a lot of identical content lines even if the full text differs
    (e.g. two KPI slides with the same metric labels but different numbers).
    Each slide is counted at most once."""
    flagged = set()
    for i in range(len(line_sets)):
        for j in range(i + 1, len(line_sets)):
            a, b = line_sets[i], line_sets[j]
            if not a or not b:
                continue
            shared = len(a & b)
            if shared == 0:
                continue
            jac = shared / len(a | b)
            if shared >= NEAR_DUP_SHARED_LINES or jac >= NEAR_DUP_JACCARD:
                flagged.add(j)
    return len(flagged)


def _evenness_score(values, label):
    if len(values) < 2:
        return {"score": 5, "detail": f"недостаточно слайдов для оценки {label}"}
    mean = sum(values) / len(values)
    if mean == 0:
        return {"score": 5, "detail": f"нет {label}"}
    var = sum((v - mean) ** 2 for v in values) / len(values)
    cv = (var ** 0.5) / mean
    return {"score": _cv_to_score(cv), "detail": f"разброс {label} по слайдам CV={cv:.2f}"}
