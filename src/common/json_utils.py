import json
import re


def _try_repair(candidate):
    """Targets one specific, observed GigaChat failure mode: it sometimes forgets to
    close the last object in an array before closing the array itself (ending in "]]"
    instead of "}]"). If brace/bracket counts are off by exactly one closing "}",
    insert it right before the final "]" and retry — otherwise give up."""
    if candidate.count("{") - candidate.count("}") != 1:
        return None
    stripped = candidate.rstrip()
    if not stripped.endswith("]"):
        return None
    idx = len(candidate) - len(candidate) + candidate.rfind("]")
    repaired = candidate[:idx] + "}" + candidate[idx:]
    try:
        return json.loads(repaired)
    except json.JSONDecodeError:
        return None


# Гибридные reasoning-модели (Qwen3, открытые веса до 35B) кладут рассуждение в
# <think>…</think> перед ответом. В рассуждении бывают скобки и черновики JSON,
# а поиск ниже берёт ПЕРВУЮ скобку — черновик выдавался бы за ответ. Незакрытый
# <think> (ответ обрезан по max_tokens) — ответа нет вовсе.
_THINK_RE = re.compile(r"<think>.*?(</think>|$)", re.DOTALL | re.IGNORECASE)


def extract_json(text):
    text = _THINK_RE.sub("", text)
    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    match = re.search(r"(\{.*\}|\[.*\])", cleaned, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON found in model response: {text!r}")

    candidate = match.group(0)
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        repaired = _try_repair(candidate)
        if repaired is not None:
            return repaired
        raise
