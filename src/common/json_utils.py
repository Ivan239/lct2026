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


def extract_json(text):
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
