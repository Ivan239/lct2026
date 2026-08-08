"""Clustering regressions: the 69-slide survey deck once produced 69
singleton groups (exact-tuple signature matching); preset decks must never
over-merge; identical geometry across templates must fingerprint identically
(that's what makes the cross-template cache work)."""

import glob

from conftest import PRESET_A, PRESET_B, SURVEY_31, SURVEY_69, TEMPLATES_DIR, requires

from design_system.clustering import cluster_slides, slide_features, slide_fingerprint
from template_parser.parser import extract_template


def _cluster(path):
    ts = extract_template(path)
    pngs = sorted(glob.glob(path.replace("templates", "rendered").replace(".pptx", "-*.png")))
    return ts, cluster_slides(ts["slides"], ts["slide_size_in"], pngs or None)


@requires(SURVEY_69)
def test_survey_69_clusters_into_families():
    ts, clusters = _cluster(SURVEY_69)
    assert len(ts["slides"]) == 69
    # Human count is ~12-15 visual families; anything near 69 means clustering
    # degenerated back to per-slide singletons, anything near 1 means collapse.
    assert 10 <= len(clusters) <= 25
    # The four purple section dividers must land in one cluster (indices from
    # the visually verified run — see docs/IMPROVEMENT_PLAN.md item 3).
    by_index = {idx: tuple(sorted(c["indices"])) for c in clusters for idx in c["indices"]}
    assert by_index[56] == by_index[61] == by_index[63] == by_index[65]


@requires(PRESET_A)
def test_preset_slides_stay_separate():
    ts, clusters = _cluster(PRESET_A)
    # 4 slides, 4 distinct archetypes — merging any two is over-merging.
    assert len(clusters) == len(ts["slides"]) == 4


@requires(PRESET_A)
@requires(PRESET_B)
def test_identical_geometry_shares_fingerprints():
    fps = []
    for path in (PRESET_A, PRESET_B):
        ts = extract_template(path)
        fps.append([
            slide_fingerprint(slide_features(s, ts["slide_size_in"])) for s in ts["slides"]
        ])
    # The two presets are the same deck re-themed: geometry identical, so the
    # cache key must match slide-for-slide (this is the "second customer with
    # a standard layout costs zero LLM calls" property).
    assert fps[0] == fps[1]


@requires(SURVEY_31)
def test_slide_description_carries_font_size_and_canvas_size():
    """The classifier judges from this description alone, and it used to drop
    the two facts that decide "title vs caption". Measured on this real customer
    template: its opening slide is one 8.15x3.2in box reading "Employee Short
    Survey Results" — set at 92pt — and it came back "other", which makes the
    deck's most visible slide unusable. Clustering is not the culprit: that
    slide is its own cluster, so the classifier judged it on its own.

    A box size means nothing without the canvas (8.15in is 61% of a 13.33in
    slide and all of an 8in one), so the slide size is stated too."""
    from design_system.extractor import _describe_slide

    struct = extract_template(SURVEY_31)
    title = _describe_slide(struct["slides"][0], struct["slide_size_in"])
    assert "13.33x7.5" in title, title
    assert "font 92pt" in title, title

    # A content slide's caption sits an order of magnitude lower — that contrast
    # is the whole point of stating the size.
    chart = _describe_slide(struct["slides"][4], struct["slide_size_in"])
    assert "font 40pt" in chart and "font 18pt" in chart, chart


@requires(TEMPLATES_DIR)
def test_description_never_invents_a_font_size():
    """Google-Slides exports leave most runs unsized, and the "title = biggest
    font" heuristic is already blind on them. Stating a guessed number here
    would push that blindness into the classifier too, so sized runs get a font
    clause and unsized ones get nothing."""
    import os

    from design_system.extractor import _describe_slide

    mono = os.path.join(TEMPLATES_DIR, "custom_838830368dac3116.pptx")
    if not os.path.exists(mono):
        return
    struct = extract_template(mono)
    described = _describe_slide(struct["slides"][2], struct["slide_size_in"])
    shape_lines = [ln for ln in described.splitlines() if ln.startswith("- ")]
    with_font = [ln for ln in shape_lines if ", font " in ln]
    assert 0 < len(with_font) < len(shape_lines), (
        "expected SOME shapes sized and some not on a Google-Slides export")
    assert all("font 0pt" not in ln and "font Nonept" not in ln for ln in shape_lines)
