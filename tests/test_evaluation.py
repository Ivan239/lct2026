"""Evaluation harness — the LLM-free parts (weighting + deterministic checks).
The vision/content judge needs GigaChat and is exercised manually, not here."""

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


def test_guard_widow_glues_last_two_words():
    from content_parser.two_phase import _guard_widow
    assert _guard_widow("Рост точности принятия решений") == "Рост точности принятия решений"
    assert _guard_widow("Быстрая автоматизация") == "Быстрая автоматизация"  # 2 words untouched
    assert _guard_widow("Итог") == "Итог"


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


def test_unavailable_ids_selects_only_judge_failures():
    from evaluation.judge import JUDGE_UNAVAILABLE, _unavailable_ids
    scores = {
        "a": {"score": 4, "detail": "ok"},                                # scored
        "b": {"score": None, "detail": "неприменимо: в деке нет медиа"},  # inapplicable N/A
        "c": {"score": None, "detail": f"{JUDGE_UNAVAILABLE}: ConnectionError"},  # judge failure
        "d": {"score": None, "detail": f"{JUDGE_UNAVAILABLE}: Timeout"},          # judge failure
    }
    assert _unavailable_ids(scores, ["a", "b", "c", "d"]) == ["c", "d"]


def test_judge_recovers_transient_failure(monkeypatch):
    """A transient tunnel drop that N/A's one chunk on the first pass must be
    recovered by the second pass — not left permanently N/A (a real run lost 6
    criteria to a momentary ConnectionError)."""
    import json as _json
    import requests
    from evaluation import judge as J
    from evaluation.judge import JUDGE_UNAVAILABLE

    monkeypatch.setattr("common.model_fallback.NETWORK_RETRY_DELAY_SECONDS", 0)
    payload = _json.dumps({cid: {"score": 4, "note": "ok"} for cid in J.VISUAL_IDS + J.CONTENT_IDS})

    class Flaky:
        def __init__(self):
            self.calls = 0

        def upload_file(self, path, purpose="general"):
            return {"id": "img"}

        def chat(self, messages, model=None, **kw):
            self.calls += 1
            if self.calls <= 5:  # exhaust the first single-model chunk's retries
                raise requests.exceptions.ConnectionError("tunnel drop")
            return {"choices": [{"message": {"content": payload}}]}

    scores = J.judge(Flaky(), ["a.png"], "brief", ["s1", "s2"],
                     vision_models=["M"], text_models=["M"])
    still_failed = [c for c, v in scores.items()
                    if v["score"] is None and str(v["detail"]).startswith(JUDGE_UNAVAILABLE)]
    assert still_failed == [], f"recovery left criteria unavailable: {still_failed}"
    assert scores["1.2"]["score"] == 4  # first (failed) visual chunk recovered


def test_gen_models_adds_max_safety_net():
    """A weak tier that returns an invalid block must fall back to GigaChat-2-Max
    instead of crashing the iteration — but only for GigaChat models (a non-
    GigaChat client can't serve a GigaChat name)."""
    from evaluation.loop import _gen_models
    assert _gen_models("GigaChat-2") == ["GigaChat-2", "GigaChat-2-Max"]
    assert _gen_models("GigaChat-3-Ultra") == ["GigaChat-3-Ultra", "GigaChat-2-Max"]
    assert _gen_models("GigaChat-2-Max") == ["GigaChat-2-Max"]  # no self-duplicate
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
