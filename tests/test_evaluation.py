"""Evaluation harness — deterministic checks, weighting, and Claude's own
score-merge step (claude_review). No LLM API is ever called to judge a deck."""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from pptx import Presentation
from pptx.util import Emu, Inches, Pt

from evaluation import deterministic, rubric


def _all(score):
    return {cid: {"score": score, "detail": ""} for cid in rubric.CRITERIA}


def test_weighted_total_all_fives_is_100():
    assert rubric.weighted_total(_all(5)) == 100.0


def test_weighted_total_all_fours_is_80():
    assert rubric.weighted_total(_all(4)) == 80.0


def test_na_bucket_is_renormalised_not_zeroed():
    """A text-only deck has no images to score — the images bucket must drop out
    and the remaining weights renormalise, not drag the total down."""
    scores = _all(5)
    for cid in rubric.BUCKETS["images_infographics"]["criteria"]:
        scores[cid] = {"score": None, "detail": "N/A"}
    assert rubric.weighted_total(scores) == 100.0
    assert rubric.bucket_breakdown(scores)["images_infographics"]["applicable"] is False


def test_one_weak_bucket_pulls_total_by_its_weight():
    scores = _all(5)
    # technical bucket (5%) all 1s: total = 95%*100 + 5%*20 = 96.0
    for cid in rubric.BUCKETS["technical"]["criteria"]:
        scores[cid] = {"score": 1, "detail": ""}
    assert rubric.weighted_total(scores) == 96.0


def test_every_criterion_is_in_exactly_one_bucket():
    assigned = [c for b in rubric.BUCKETS.values() for c in b["criteria"]]
    assert sorted(assigned) == sorted(rubric.CRITERIA)
    assert len(assigned) == len(set(assigned))


def _deck(tmp_path, slides):
    """slides: list of (title, [body_lines]). Returns a saved .pptx path."""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    for title, body in slides:
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
        tb.text_frame.text = title
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
        box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(4))
        tf = box.text_frame
        for i, line in enumerate(body):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line
            p.runs[0].font.size = Pt(18)
    out = str(tmp_path / "deck.pptx")
    prs.save(out)
    return out


def test_duplicate_titles_scored_low(tmp_path):
    path = _deck(tmp_path, [("Вывод", ["a"]), ("Вывод", ["b"]), ("Вывод", ["c"])])
    scores = deterministic.evaluate(path)
    assert scores["2.3"]["score"] <= 2  # 2 of 3 titles duplicated


def test_unique_titles_scored_high(tmp_path):
    path = _deck(tmp_path, [("Проблема", ["a"]), ("Решение", ["b"]), ("Итог", ["c"])])
    scores = deterministic.evaluate(path)
    assert scores["2.3"]["score"] == 5


def test_overlong_title_flagged(tmp_path):
    long_title = "Основные преимущества внедрения современных технологий в процесс управления"
    path = _deck(tmp_path, [(long_title, ["a"]), ("Кратко", ["b"])])
    scores = deterministic.evaluate(path)
    assert scores["2.1"]["score"] < 5


def test_deck_media_counts_image_placeholders(tmp_path):
    """Image skeletons (named frames) are counted so the evaluator scores image
    PLACEMENT even though there's no real picture yet."""
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    from generator.synthesizer import IMAGE_PLACEHOLDER_NAME

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(6), Inches(3))
    shp.name = IMAGE_PLACEHOLDER_NAME
    path = str(tmp_path / "ph.pptx")
    prs.save(path)
    media = deterministic.deck_media(path)
    assert media["placeholders"] == 1
    assert media["substantive_pictures"] == 0


def test_text_only_deck_has_no_substantive_media(tmp_path):
    """The image-criteria N/A gate rests on this: a text deck reports zero
    substantive pictures and zero charts, so section 4/7 get forced to N/A
    instead of the model's occasional bogus '1 — нет изображений'."""
    path = _deck(tmp_path, [("Заголовок", ["строка один", "строка два"])])
    media = deterministic.deck_media(path)
    assert media["substantive_pictures"] == 0
    assert media["charts"] == 0


def test_semantic_near_duplicate_slides_flagged(tmp_path):
    """Two KPI slides sharing the same metric LABELS (different numbers) are a
    near-duplicate the byte-equality check misses — the near-dup arm must catch
    them. Grounded in iteration #1's real defect (slides 5 & 7)."""
    labels = ["Рост конверсии сделок", "Экономия времени менеджеров",
              "Точность прогнозирования спроса"]
    path = _deck(tmp_path, [
        ("Ключевые возможности", ["+25%"] + labels),
        ("Другой слайд", ["Совсем иной", "самостоятельный", "контент здесь"]),
        ("Результаты внедрения", ["+15%"] + labels),
    ])
    scores = deterministic.evaluate(path)
    assert "1 смысловых" in scores["dop_no_dup_slides"]["detail"]
    assert scores["dop_no_dup_slides"]["score"] < 5


def test_distinct_slides_not_flagged_as_duplicates(tmp_path):
    path = _deck(tmp_path, [
        ("Проблема", ["данные разрознены", "отчёты вручную"]),
        ("Решение", ["единый дашборд", "автоотчёты"]),
        ("Итог", ["запустите пилот", "свяжитесь с нами"]),
    ])
    scores = deterministic.evaluate(path)
    assert scores["dop_no_dup_slides"]["score"] == 5


def test_pacing_excludes_structural_sparse_slides(tmp_path):
    """A sparse divider is sparse BY DESIGN — it must not drag down the
    distribution/pacing evenness of the content slides."""
    body = ["строка текста здесь " * 4]
    path = _deck(tmp_path, [
        ("Слайд один", body),
        ("Слайд два", body),
        ("Раздел", ["итог"]),      # tiny divider
        ("Слайд три", body),
    ])
    without = deterministic.evaluate(path)["dop_distribution"]["score"]
    with_roles = deterministic.evaluate(
        path, slide_roles={2: "section_divider"})["dop_distribution"]["score"]
    assert with_roles > without  # excluding the sparse divider lifts even-content score


def test_guard_widow_glues_last_two_words():
    from content_parser.two_phase import _guard_widow
    assert _guard_widow("Рост точности принятия решений") == "Рост точности принятия решений"
    assert _guard_widow("Быстрая автоматизация") == "Быстрая автоматизация"  # 2 words untouched
    assert _guard_widow("Итог") == "Итог"


def test_enforce_budgets_guards_stat_labels():
    from content_parser.two_phase import _enforce_text_budgets
    block = {"type": "stats_kpi", "title": "Итоги",
             "stats": [["+25%", "Рост конверсии продаж"], ["95%", "Точность"]]}
    out = _enforce_text_budgets(block, "stats_kpi")
    assert out["stats"][0][0] == "+25%"                                    # number untouched
    assert " " in out["stats"][0][1]                                  # last two words glued
    assert out["stats"][0][1].replace(" ", " ") == "Рост конверсии продаж"
    assert out["stats"][1][1] == "Точность" and " " not in out["stats"][1][1]  # 1 word untouched


def test_tokens_nbsp_aware_is_noop_without_nbsp():
    from evaluation.deterministic import _tokens
    assert _tokens("a b c") == ["a", "b", "c"]           # identical to .split()
    assert _tokens("one  two\tthree") == ["one", "two", "three"]
    assert _tokens("aa bb cc") == ["aa", "bb cc"]  # NBSP keeps the pair


def test_is_widow_respects_nbsp_glue():
    """A lone last word is a widow; the same words glued with NBSP are one token
    (two visible words) and must NOT be flagged."""
    from evaluation.deterministic import _is_widow

    class _FakeMetrics:
        def text_width_pt(self, s, size_pt):
            return len(s) * 10.0

    m = _FakeMetrics()
    plain_widow, _ = _is_widow("aa bb cc dd", 1.4, 18, m, 0)
    glued_ok, _ = _is_widow("aa bb cc dd", 1.4, 18, m, 0)
    assert plain_widow is True
    assert glued_ok is False


def test_gen_models_adds_max_safety_net():
    """A weak tier that returns an invalid block must fall back to GigaChat-2-Max
    instead of crashing the iteration — but only for GigaChat models (a non-
    GigaChat client can't serve a GigaChat name)."""
    from evaluation.loop import _gen_models
    assert _gen_models("GigaChat-2") == ["GigaChat-2", "GigaChat-2-Max"]
    assert _gen_models("GigaChat-3-Ultra") == ["GigaChat-3-Ultra", "GigaChat-2-Max"]
    # Max gets a net too (Pro) — forcing Max used to have no fallback and could crash
    assert _gen_models("GigaChat-2-Max") == ["GigaChat-2-Max", "GigaChat-2-Pro"]
    assert _gen_models("rtx") == ["rtx"]                        # non-GigaChat: unchanged


def _slide_png(path, bg, ink):
    """A synthetic 'slide': a solid bg with a big block of ink 'text' — enough
    to exercise the dominant-ink contrast logic without LibreOffice."""
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (960, 540), bg)
    d = ImageDraw.Draw(img)
    for i in range(6):  # several fat bars = a substantial ink block
        d.rectangle([120, 120 + i * 40, 840, 145 + i * 40], fill=ink)
    img.save(path)
    return str(path)


def test_contrast_flags_light_ink_on_white(tmp_path):
    from evaluation import contrast

    # Same brand ink (orange), different background — exactly the real defect:
    # the deck's palette is designed for navy and washes out on a white breather.
    good = _slide_png(tmp_path / "navy.png", (16, 32, 64), (240, 128, 32))     # orange on navy ~6:1
    bad = _slide_png(tmp_path / "white.png", (240, 240, 240), (240, 128, 32))  # orange on white ~2.4:1
    res = contrast.evaluate_contrast([good, bad])
    assert res["low_contrast_slides"] == [1]  # only the grey-on-white slide
    assert res["per_slide"][0] > contrast.LOW_CONTRAST_RATIO
    assert res["per_slide"][1] < contrast.LOW_CONTRAST_RATIO


def test_contrast_ratio_matches_wcag_black_white():
    from evaluation import contrast
    # black vs white is the canonical 21:1
    assert round(contrast.contrast_ratio((0, 0, 0), (255, 255, 255)), 1) == 21.0


def test_out_of_bounds_box_flagged(tmp_path):
    path = _deck(tmp_path, [("Ок", ["a"])])
    prs = Presentation(path)
    box = prs.slides[0].shapes.add_textbox(Emu(prs.slide_width - Inches(0.2)),
                                           Inches(2), Inches(4), Inches(1))
    box.text_frame.text = "уехал за правый край"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(18)
    prs.save(path)
    scores = deterministic.evaluate(path)
    assert scores["9.1"]["score"] < 5


def test_outline_always_gets_one_image_slide():
    """Asking the model for an image slide works ~half the time — the rule is
    enforced in code (same lesson as the divider cap)."""
    from content_parser.two_phase import MAX_BLOCKS, _enforce_outline_rules

    def roles(outline):
        return [i["role"] for i in _enforce_outline_rules(outline)]

    # none requested -> exactly one inserted, just before the closing
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "bullet_list", "theme": "b", "count": 3},
                 {"role": "closing", "theme": "c", "count": None}])
    assert out == ["title", "bullet_list", "image_caption", "closing"]

    # model omitted the closing -> one is enforced, and the image lands before it
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "bullet_list", "theme": "b", "count": 3}])
    assert out == ["title", "bullet_list", "image_caption", "closing"]

    # already present -> kept, not duplicated
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "image_caption", "theme": "i", "count": None},
                 {"role": "closing", "theme": "c", "count": None}])
    assert out.count("image_caption") == 1

    # more than one requested -> capped at one
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "image_caption", "theme": "i1", "count": None},
                 {"role": "image_caption", "theme": "i2", "count": None}])
    assert out.count("image_caption") == 1

    # a full outline stays within MAX_BLOCKS
    full = [{"role": "title", "theme": "t", "count": None}]
    full += [{"role": "bullet_list", "theme": f"b{i}", "count": 3} for i in range(MAX_BLOCKS + 3)]
    assert len(_enforce_outline_rules(full)) <= MAX_BLOCKS


def test_apply_claude_scores_recomputes_total():
    from evaluation.claude_review import apply_claude_scores

    result = {"scores": {cid: {"title": t, "mode": m, "score": None, "detail": "ожидает оценки Клода"}
                         for cid, (t, m) in rubric.CRITERIA.items()}}
    # deterministic criteria already scored (as evaluate_deck would leave them)
    for cid in rubric.CRITERIA:
        if rubric.CRITERIA[cid][1] == "det":
            result["scores"][cid]["score"] = 5

    llm_scores = {cid: (5, "ок") for cid, (_, mode) in rubric.CRITERIA.items() if mode == "llm"}
    out = apply_claude_scores(result, llm_scores)
    assert out["total_100"] == 100.0
    assert out["llm_evaluated"] is True
    assert out["judge"] == "claude"


def test_apply_claude_scores_rejects_deterministic_ids():
    from evaluation.claude_review import apply_claude_scores
    result = {"scores": {cid: {"title": t, "mode": m, "score": None, "detail": ""}
                         for cid, (t, m) in rubric.CRITERIA.items()}}
    det_id = next(cid for cid, (_, mode) in rubric.CRITERIA.items() if mode == "det")
    try:
        apply_claude_scores(result, {det_id: (5, "не моя забота")})
        assert False, "should have rejected a deterministic-mode criterion id"
    except ValueError:
        pass


def test_apply_claude_scores_partial_leaves_rest_na():
    from evaluation.claude_review import apply_claude_scores
    result = {"scores": {cid: {"title": t, "mode": m, "score": None, "detail": "ожидает оценки Клода"}
                         for cid, (t, m) in rubric.CRITERIA.items()}}
    out = apply_claude_scores(result, {"1.3": (2, "верх плотный, низ пустой")})
    assert out["scores"]["1.3"]["score"] == 2
    assert out["scores"]["1.6"]["score"] is None  # untouched llm criterion stays N/A


def test_stat_fingerprints_and_duplicate_rejection():
    """Two stat slides showing the same numbers read as one slide shown twice
    (real defect: slides 5/6 both +25% / -40% / 92%). Uniqueness is enforced by
    validation+retry, not by asking the model nicely."""
    from content_parser.two_phase import _reject_duplicate_stats, stat_fingerprints

    first = {"stats": [["+25%", "Рост конверсии в сделку"], ["-40%", "Экономия времени"]]}
    nums, labels = stat_fingerprints(first)
    assert nums == {"+25%", "-40%"}
    assert "рост конверсии в сделку" in labels

    # a later slide repeating a NUMBER must be rejected (-> retried)
    dup_num = {"stats": [["+25%", "Совсем другая метрика"]]}
    try:
        _reject_duplicate_stats(dup_num, "stats_kpi", nums, labels)
        assert False, "duplicate number should have been rejected"
    except ValueError:
        pass

    # repeating a LABEL must be rejected too
    dup_label = {"stats": [["+99%", "Рост конверсии в сделку"]]}
    try:
        _reject_duplicate_stats(dup_label, "stats_kpi", nums, labels)
        assert False, "duplicate label should have been rejected"
    except ValueError:
        pass

    # genuinely different metrics pass
    fresh = {"stats": [["3 дня", "Срок внедрения"]]}
    _reject_duplicate_stats(fresh, "stats_kpi", nums, labels)

    # non-stat roles are untouched
    _reject_duplicate_stats({"bullets": ["a"]}, "bullet_list", nums, labels)


def test_outline_always_ends_with_a_closing():
    """A deck must end. Asking for a closing "if the brief suits one" got taken
    as optional — a real deck finished on a stats slide with no wrap-up."""
    from content_parser.two_phase import MAX_BLOCKS, _enforce_outline_rules

    def roles(outline):
        return [i["role"] for i in _enforce_outline_rules(outline)]

    # model omitted the closing -> one is appended at the end
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "stats_kpi", "theme": "s", "count": 3}])
    assert out[-1] == "closing"
    assert out.count("closing") == 1

    # already present -> not duplicated, stays last
    out = roles([{"role": "title", "theme": "t", "count": None},
                 {"role": "bullet_list", "theme": "b", "count": 3},
                 {"role": "closing", "theme": "c", "count": None}])
    assert out.count("closing") == 1 and out[-1] == "closing"

    # a full outline still gets a closing and stays within MAX_BLOCKS
    full = [{"role": "title", "theme": "t", "count": None}]
    full += [{"role": "bullet_list", "theme": f"b{i}", "count": 3} for i in range(MAX_BLOCKS + 3)]
    out = roles(full)
    assert len(out) <= MAX_BLOCKS
    assert out[-1] == "closing"


def test_kpi_number_slot_must_be_a_bare_figure():
    """The big KPI slot is a display figure, not a sentence. The model kept
    stuffing the metric name into it ("+18% конверсии"), which renders as three
    orange sentences with the labels below repeating the same words."""
    from content_parser.two_phase import _is_display_figure, _reject_wordy_figures

    for good in ("+25%", "-30 часов", "95%", "3 дня", "8 из 10", "x10"):
        assert _is_display_figure(good), good
    for bad in ("+18% конверсии", "-20% затрат времени", "90% точность прогноза",
                "рост эффективности", ""):
        assert not _is_display_figure(bad), bad

    try:
        _reject_wordy_figures({"stats": [["+18% конверсии", "рост"]]}, "stats_kpi")
        assert False, "wordy figure should have been rejected"
    except ValueError:
        pass

    # clean block passes, non-stat roles untouched
    _reject_wordy_figures({"stats": [["+25%", "Рост конверсии"]]}, "stats_kpi")
    _reject_wordy_figures({"bullets": ["a"]}, "bullet_list")
