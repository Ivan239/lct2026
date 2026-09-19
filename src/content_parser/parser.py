import json

from common.json_utils import extract_json
from common.model_fallback import TEXT_MODELS, call_with_model_fallback
from common.prompt_files import load_prompt

SCHEMA_PROMPT = load_prompt("schema_blocks.v1.txt")


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


TITLE_PROMPT_TEMPLATE = load_prompt("title_for_block.v1.txt")


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


RESIZE_PROMPT_TEMPLATE = load_prompt("resize_block.v1.txt")

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
