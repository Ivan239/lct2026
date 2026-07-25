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

from pptx import Presentation
from pptx.util import Emu

from fonts.metrics import FontResolver
from generator.generator import _pick_body_shape, _pick_title_shape
from generator.text_fit import (
    estimate_block_height_in,
    horizontal_margins_in,
    vertical_insets_emu,
)
from qa.geometry import _is_content_shape
from template_parser.parser import extract_theme

OVERFLOW_TOLERANCE = 1.05          # same slack qa/geometry uses before "it overflows"
MIN_READABLE_PT = 11.0             # below this, body text reads as fine print
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
    'изображения' the rubric's section 4 is about."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(pptx_path)
    area = (prs.slide_width or 1) * (prs.slide_height or 1)
    pictures = charts = 0
    for slide in prs.slides:
        for s in slide.shapes:
            if getattr(s, "has_chart", False):
                charts += 1
            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if (s.width or 0) * (s.height or 0) / area > SUBSTANTIVE_PICTURE_FRACTION:
                    pictures += 1
    return {"substantive_pictures": pictures, "charts": charts}


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
    return [s for s in slide.shapes if _is_content_shape(s)]


def _shape_runs(shape):
    return [run for p in shape.text_frame.paragraphs for run in p.runs]


def _shape_max_size(shape):
    sizes = [r.font.size.pt for r in _shape_runs(shape) if r.font.size and r.text.strip()]
    return max(sizes) if sizes else None


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
        tiny = sum(1 for s in content if (_shape_max_size(s) or 99) < MIN_READABLE_PT)
        oob = 0
        near_edge = 0
        for s in content:
            if s.left is None or s.top is None or not s.width or not s.height:
                continue
            l, t, r, b = s.left, s.top, s.left + s.width, s.top + s.height
            if (l < -OOB_TOLERANCE_FRACTION * W or t < -OOB_TOLERANCE_FRACTION * H
                    or r > W * (1 + OOB_TOLERANCE_FRACTION) or b > H * (1 + OOB_TOLERANCE_FRACTION)):
                oob += 1
            if (l < EDGE_SAFE_FRACTION * W or t < EDGE_SAFE_FRACTION * H
                    or r > W * (1 - EDGE_SAFE_FRACTION) or b > H * (1 - EDGE_SAFE_FRACTION)):
                near_edge += 1

        # orphan last lines across content paragraphs
        orphan = 0
        multiline = 0
        for s in content:
            if not s.width:
                continue
            size_pt = _shape_max_size(s)
            fn = next((r.font.name for r in _shape_runs(s) if r.font.name), None)
            m = metrics_for(fn) if fn else None
            if not size_pt or m is None:
                continue
            for p in s.text_frame.paragraphs:
                txt = "".join(r.text for r in p.runs).strip()
                if not txt:
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

    # 1.1 readability — overflow + fine print
    bad = sum(p["overflow"] + p["tiny"] for p in per_slide)
    scores["1.1"] = {
        "score": _rate_to_score(bad / total_content),
        "detail": f"{bad} из {total_content} контентных боксов переполнены/мелкий шрифт",
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
    # 9.1 out-of-bounds
    scores["9.1"] = {
        "score": _rate_to_score(rate_over_content("oob")),
        "detail": f"{sum(p['oob'] for p in per_slide)} боксов выходят за границы слайда",
    }
    # dop_safe_margins
    scores["dop_safe_margins"] = {
        "score": _rate_to_score(rate_over_content("near_edge")),
        "detail": f"{sum(p['near_edge'] for p in per_slide)} боксов ближе {EDGE_SAFE_FRACTION:.0%} к краю",
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
    # dop_distribution / dop_pacing — evenness of per-slide char counts / shapes
    char_counts = [p["chars"] for p in per_slide if p["chars"] > 0]
    scores["dop_distribution"] = _evenness_score(char_counts, "символов")
    shape_counts = [p["n_content"] for p in per_slide if p["n_content"] > 0]
    scores["dop_pacing"] = _evenness_score(
        [c + s * 40 for c, s in zip(char_counts, shape_counts)] or char_counts, "объёма"
    )
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
