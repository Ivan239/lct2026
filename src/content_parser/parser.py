import json

from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, call_with_model_fallback

SCHEMA_PROMPT = """Ты помогаешь превратить сырой бриф о продукте в структурированный план презентации.

Разбей текст на блоки контента. Каждый блок должен иметь один из типов:
- title: {"type": "title", "title": "...", "subtitle": "..."}
- section_divider: {"type": "section_divider", "title": "..."} — короткая крупная фраза-переход
  между смысловыми разделами презентации. НЕ БОЛЕЕ ОДНОГО такого блока на всю презентацию,
  и только если содержание реально делится на два крупных, непохожих раздела (например
  "Проблема" и "Решение"). Для короткого питча (как большинство брифов) этот тип обычно вообще
  не нужен — пропусти его, если не уверен, что он оправдан. Не вставляй его перед каждым блоком.
- bullet_list: {"type": "bullet_list", "title": "...", "bullets": ["...", "...", "..."]} (ровно 3 буллета)
- stats_kpi: {"type": "stats_kpi", "title": "...", "stats": [["число", "подпись"], ...]} (ровно 3 пары)
- two_column_comparison: {"type": "two_column_comparison", "title": "...",
  "left_heading": "...", "left_points": ["...", "...", "..."],
  "right_heading": "...", "right_points": ["...", "...", "..."]} (ровно по 3 пункта в колонке)
- closing: {"type": "closing", "title": "...", "subtitle": "..."} — финальный слайд презентации
  (благодарность, призыв к действию, контакты). Используй как последний блок, если брифу
  естественно подходит завершение питча.

Всегда начинай с одного блока title. Порядок остальных блоков — как логично для питча.
Если брифу естественно подходит финальный слайд — заверши блоком closing.

ВАЖНО: у каждого блока bullet_list, stats_kpi, two_column_comparison и closing ОБЯЗАТЕЛЬНО
должно быть непустое поле "title" со своим конкретным заголовком — даже если непосредственно
перед этим блоком стоит section_divider. section_divider — это только переход между разделами,
он НЕ заменяет и не отменяет заголовок следующего за ним слайда.

Ответь ТОЛЬКО валидным JSON-массивом блоков, без пояснений и markdown-разметки.

Бриф:
---
__BRIEF__
---
"""


MAX_SECTION_DIVIDERS = 1

# Which type carries "a list of [num, label] pairs" — the one shape the model
# has been observed to get wrong (a bare ["x10"] with no label), unlike plain
# string lists (bullets, left_points, right_points) which have no inner shape
# to violate. Mirrors content_parser.two_phase._validate_block's stricter
# per-item check, which this legacy path historically lacked.
PAIR_FIELD_BY_TYPE = {"stats_kpi": "stats"}


def _sanitize_pairs(block):
    field = PAIR_FIELD_BY_TYPE.get(block.get("type"))
    if field in block:
        block[field] = [item for item in block[field] if isinstance(item, (list, tuple)) and len(item) == 2]
    return block


def _cap_section_dividers(blocks, max_count=MAX_SECTION_DIVIDERS):
    """The model doesn't reliably respect a soft "at most N" instruction in the
    prompt (observed inserting one before nearly every block despite being told
    not to) — a hard cardinality limit like this needs to be enforced in code,
    not just requested nicely."""
    kept = []
    seen = 0
    for block in blocks:
        if block["type"] == "section_divider":
            seen += 1
            if seen > max_count:
                continue
        kept.append(block)
    return kept


def parse_brief(client, brief_text, models=TEXT_MODELS):
    prompt = SCHEMA_PROMPT.replace("__BRIEF__", brief_text)

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=2000)
        return extract_json(result["choices"][0]["message"]["content"])

    blocks = call_with_model_fallback(call, models)
    blocks = _cap_section_dividers(blocks)
    blocks = [_sanitize_pairs(b) for b in blocks]
    blocks = [_ensure_title(client, b, models) for b in blocks]
    for i, block in enumerate(blocks):
        block["order"] = i
    return blocks


TITLE_PROMPT_TEMPLATE = """Вот блок контента презентации в формате JSON, без заголовка:
---
__BLOCK__
---
Придумай короткий, конкретный заголовок (title) для этого слайда, отражающий его содержание.
Ответь ТОЛЬКО текстом заголовка, одной строкой, без кавычек и пояснений.
"""


def _ensure_title(client, block, models=TEXT_MODELS):
    """The model doesn't reliably follow even an explicit, forceful instruction
    to always fill "title" (observed leaving it blank on every block right
    after a section_divider, as if the divider already covered it) — rather
    than keep re-wording the prompt, patch it up after the fact for whichever
    specific blocks actually came back without one."""
    if block.get("title") or block["type"] == "title":
        return block

    prompt = TITLE_PROMPT_TEMPLATE.replace("__BLOCK__", json.dumps(block, ensure_ascii=False))

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=100)
        return result["choices"][0]["message"]["content"].strip().strip('"').strip()

    try:
        title = call_with_model_fallback(call, models)
    except Exception:
        title = ""
    block["title"] = title
    return block


RESIZE_PROMPT_TEMPLATE = """Вот блок контента презентации в формате JSON:
---
__BLOCK__
---

На конкретном слайде шаблона реально помещается ровно __COUNT__ элемент(ов) в поле "__FIELD__"
(а не столько, сколько сейчас в блоке). Перепиши этот блок так, чтобы в поле "__FIELD__" было
ТОЧНО __COUNT__ элементов — не больше и не меньше, сохранив общий смысл и стиль. Если нужно
больше элементов, чем есть сейчас — раздели существующие мысли на более конкретные или добавь
логичное продолжение по смыслу блока. Если нужно меньше — объедини наиболее важные.
Остальные поля блока оставь как есть.

Ответь ТОЛЬКО валидным JSON-объектом блока (тот же формат, что во входе), без пояснений и markdown.
"""

# Which field holds "the list whose length must match slide capacity" per type.
RESIZE_FIELD_BY_TYPE = {
    "bullet_list": "bullets",
    "stats_kpi": "stats",
}


def resize_block(client, block, target_count, models=TEXT_MODELS):
    """Regenerates just this one block so its item count matches what the
    matched template slide can actually hold (e.g. 4 icon-bullet slots instead
    of the usual 3) — instead of silently dropping/leaving slots blank at fill
    time. Best-effort: falls back to the original block on any failure, since
    the filler already degrades gracefully on a mismatch."""
    field = RESIZE_FIELD_BY_TYPE.get(block["type"])
    if field is None or target_count == len(block.get(field, [])):
        return block

    prompt = (
        RESIZE_PROMPT_TEMPLATE
        .replace("__BLOCK__", json.dumps(block, ensure_ascii=False))
        .replace("__COUNT__", str(target_count))
        .replace("__FIELD__", field)
    )

    def call(model):
        result = client.chat([{"role": "user", "content": prompt}], model=model, max_tokens=1000)
        return extract_json(result["choices"][0]["message"]["content"])

    try:
        resized = call_with_model_fallback(call, models)
    except Exception:
        return block

    resized["type"] = block["type"]
    resized["order"] = block.get("order")
    resized = _sanitize_pairs(resized)
    if len(resized.get(field, [])) != target_count:
        return block
    return resized
