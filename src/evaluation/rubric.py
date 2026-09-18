"""Structured form of docs/evaluation_rubric.md.

The prose rubric is the source of truth for humans; this module is the machine
form the loop scores against. Every criterion is assigned to exactly one of the
seven weighted buckets from the rubric's final scoring table, so the weighted
100-point total is computed deterministically from per-criterion 1..5 scores —
the LLM judge never gets to invent the weighting, matching the project rule that
hard numbers are enforced in code, not left to prompt wording.

`mode` records how a criterion is scored:
  - "det"  -> a deterministic geometric/textual check (zero tokens)
  - "llm"  -> the vision/text LLM judge decides
Splitting them here (not in the judge) keeps the token spend to only the
criteria a machine genuinely can't call.
"""

# id -> (human title, mode). ids match the section numbers in the prose rubric,
# with dop_* for the "дополнительные критерии" that have no section number.
CRITERIA = {
    # 1. Визуальный дизайн
    "1.1": ("Читаемость текста", "det"),
    "1.2": ("Перекрытие текста изображениями", "llm"),
    "1.3": ("Баланс композиции", "llm"),
    "1.4": ("Использование свободного пространства", "llm"),
    "1.5": ("Выравнивание элементов", "llm"),
    "1.6": ("Визуальная иерархия", "llm"),
    "1.7": ("Единый стиль", "llm"),
    "1.8": ("Цветовая гармония", "llm"),
    # 2. Заголовки
    "2.1": ("Длина заголовка", "det"),
    "2.2": ("Информативность заголовка", "llm"),
    "2.3": ("Уникальность заголовков", "det"),
    # 3. Текст
    "3.1": ("Количество текста", "det"),
    "3.2": ("Структура текста", "llm"),
    "3.3": ("Ясность", "llm"),
    "3.4": ("Отсутствие повторов", "llm"),
    "3.5": ("Грамматика", "llm"),
    # 4. Изображения
    "4.1": ("Соответствие изображений теме", "llm"),
    "4.2": ("Качество изображений", "llm"),
    "4.3": ("Соотношение размеров изображений", "llm"),
    "4.4": ("Подписи к изображениям", "llm"),
    "4.5": ("Избыточность изображений", "llm"),
    # 5. Контент
    "5.1": ("Соответствие запросу", "llm"),
    "5.2": ("Полнота", "llm"),
    "5.3": ("Логика", "llm"),
    "5.4": ("Фактическая корректность", "llm"),
    "5.5": ("Последовательность терминов", "llm"),
    "5.6": ("Отсутствие галлюцинаций", "llm"),
    # 6. Структура презентации
    "6.1": ("Логичный порядок слайдов", "llm"),
    "6.2": ("Связность", "llm"),
    "6.3": ("Завершённость", "det"),
    # 7. Инфографика
    "7.1": ("Корректность диаграмм", "llm"),
    "7.2": ("Понятность инфографики", "llm"),
    "7.3": ("Читаемость подписей инфографики", "llm"),
    # 8. Адаптация под аудиторию
    "8.1": ("Соответствие уровню аудитории", "llm"),
    "8.2": ("Соответствие стилю", "llm"),
    # 9. Техническое качество
    "9.1": ("Нет выходящих за границы объектов", "det"),
    "9.2": ("Нет наложений элементов", "llm"),
    "9.3": ("Консистентность оформления", "llm"),
    # 10. Общая полезность
    "10.1": ("Готовность без ручной правки", "llm"),
    # 11. Дополнительные критерии
    "dop_prompt": ("Следование пользовательскому промпту", "llm"),
    "dop_brand": ("Соблюдение фирменного стиля", "llm"),
    "dop_template": ("Корректное использование шаблона", "llm"),
    "dop_distribution": ("Равномерное распределение информации", "det"),
    "dop_no_dup_slides": ("Отсутствие дублирующихся слайдов", "det"),
    "dop_text_split": ("Корректное разбиение текста по слайдам", "det"),
    "dop_emphasis": ("Уместность выделений", "llm"),
    "dop_wrap": ("Корректность переносов строк", "det"),
    "dop_icons": ("Качественные иконки вместо клипарта", "llm"),
    "dop_image_style": ("Согласованность стиля изображений", "llm"),
    "dop_orphans": ("Отсутствие «сирот» и «висячих» строк", "det"),
    "dop_safe_margins": ("Соблюдение безопасных полей", "det"),
    "dop_noise": ("Умеренное количество объектов", "det"),
    "dop_pacing": ("Единый темп презентации", "det"),
    "dop_no_placeholders": ("Нет текста-заглушки шаблона", "det"),
}

# Seven weighted buckets from the rubric's final table. weight sums to 1.0.
# Each bucket lists the criterion ids that roll up into it; a criterion belongs
# to exactly one bucket. Bucket score = mean of its non-N/A criteria.
BUCKETS = {
    "visual_design": {
        "title": "Визуальный дизайн",
        "weight": 0.25,
        "criteria": ["1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8",
                     "9.2", "dop_emphasis", "dop_noise", "dop_safe_margins"],
    },
    "readability_typography": {
        "title": "Читаемость и типографика",
        "weight": 0.15,
        "criteria": ["1.1", "3.1", "3.2", "3.3", "3.4", "3.5",
                     "dop_wrap", "dop_orphans"],
    },
    "content": {
        "title": "Контент",
        "weight": 0.25,
        "criteria": ["2.1", "2.2", "2.3", "5.2", "5.3", "5.4", "5.5", "5.6"],
    },
    "structure": {
        "title": "Структура презентации",
        "weight": 0.10,
        "criteria": ["6.1", "6.2", "6.3", "dop_distribution",
                     "dop_no_dup_slides", "dop_text_split", "dop_pacing"],
    },
    "images_infographics": {
        "title": "Изображения и инфографика",
        "weight": 0.10,
        "criteria": ["4.1", "4.2", "4.3", "4.4", "4.5", "7.1", "7.2", "7.3",
                     "dop_icons", "dop_image_style"],
    },
    "prompt_adherence": {
        "title": "Следование запросу",
        "weight": 0.10,
        "criteria": ["5.1", "8.1", "8.2", "10.1",
                     "dop_prompt", "dop_brand", "dop_template"],
    },
    "technical": {
        "title": "Техническое качество",
        "weight": 0.05,
        "criteria": ["9.1", "9.3", "dop_no_placeholders"],
    },
}

# Every criterion must live in exactly one bucket — guard against a typo silently
# dropping a criterion from the weighted total.
_assigned = [c for b in BUCKETS.values() for c in b["criteria"]]
assert sorted(_assigned) == sorted(CRITERIA), (
    "bucket/criteria mismatch: "
    f"unassigned={sorted(set(CRITERIA) - set(_assigned))}, "
    f"unknown={sorted(set(_assigned) - set(CRITERIA))}"
)
assert abs(sum(b["weight"] for b in BUCKETS.values()) - 1.0) < 1e-9, "weights must sum to 1.0"

# Score used for a criterion the deck genuinely doesn't exercise (e.g. image
# criteria on a text-only deck). N/A criteria are dropped from the mean, never
# scored as 0 — the rubric says N/A is "неприменимо", not "плохо".
NA = None


def bucket_score(bucket_key, scores):
    """Mean 1..5 of a bucket's criteria, skipping N/A / missing. None if the
    whole bucket is inapplicable (all criteria N/A)."""
    vals = [
        scores[c]["score"]
        for c in BUCKETS[bucket_key]["criteria"]
        if scores.get(c) and scores[c].get("score") is not NA
    ]
    return sum(vals) / len(vals) if vals else None


def weighted_total(scores):
    """Weighted 0..100 total. Buckets that are entirely N/A are excluded and the
    remaining weights are renormalised, so a text-only deck isn't punished for
    having no images to score."""
    parts = []
    for key, b in BUCKETS.items():
        bs = bucket_score(key, scores)
        if bs is not None:
            parts.append((b["weight"], bs))
    if not parts:
        return 0.0
    wsum = sum(w for w, _ in parts)
    normed = sum((w / wsum) * bs for w, bs in parts)
    return round(normed / 5.0 * 100.0, 1)


def bucket_breakdown(scores):
    """Per-bucket {title, weight, score_1_5, applicable} for reporting."""
    out = {}
    for key, b in BUCKETS.items():
        bs = bucket_score(key, scores)
        out[key] = {
            "title": b["title"],
            "weight": b["weight"],
            "score_1_5": round(bs, 2) if bs is not None else None,
            "applicable": bs is not None,
        }
    return out
