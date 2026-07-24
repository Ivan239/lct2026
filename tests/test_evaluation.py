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
