"""The cross-template economics claim, verified: classifying a template whose
geometry was already seen must cost ZERO LLM calls — the fingerprint cache
answers instead. Uses the two presets (same deck re-themed, identical
geometry) and a call-counting fake client."""

import json

from conftest import PRESET_A, PRESET_B, requires

from design_system.extractor import build_archetype_map
from design_system.fingerprint_cache import FingerprintCache
from template_parser.parser import extract_template


class CountingClient:
    def __init__(self):
        self.calls = 0

    def chat(self, messages, model=None, **kwargs):
        self.calls += 1
        return {"choices": [{"message": {"content": json.dumps(
            {"archetype": "bullet_list", "confidence": "high"}
        )}}]}


@requires(PRESET_A)
@requires(PRESET_B)
def test_second_template_with_known_geometry_is_free(tmp_path):
    cache = FingerprintCache(str(tmp_path / "fp.json"))

    first = CountingClient()
    build_archetype_map(first, extract_template(PRESET_A), fingerprint_cache=cache)
    assert first.calls > 0  # the first customer pays

    second = CountingClient()
    result = build_archetype_map(second, extract_template(PRESET_B), fingerprint_cache=cache)
    assert second.calls == 0  # the second one rides the cache
    assert set(result.values()) == {"bullet_list"}  # propagated from cached answers


@requires(PRESET_A)
def test_low_confidence_answers_never_poison_the_cache(tmp_path):
    cache = FingerprintCache(str(tmp_path / "fp.json"))

    class HesitantClient:
        def chat(self, messages, model=None, **kwargs):
            return {"choices": [{"message": {"content": json.dumps(
                {"archetype": "quote", "confidence": "low"}
            )}}]}

    build_archetype_map(HesitantClient(), extract_template(PRESET_A), fingerprint_cache=cache)
    assert cache._data == {}  # hesitant guesses must not be memorized


def test_an_all_other_parse_is_a_failure_not_a_cache_entry():
    """build_archetype_map degrades a slide to "other" on ANY exception, so a
    run that loses the network mid-way produces a full map of shrugs. Writing
    that to archetypes.json freezes the non-answer for that (model, template)
    pair forever, because later runs reuse the cache.

    Measured on this machine when the offline probe stumbled into it: 9 of 20
    cached parses were entirely "other", including three of the five real
    templates under one model — those pairs could only ever produce decks with
    zero native slides, and nothing said so.

    The project already applies this reasoning one level down: extractor's
    fingerprint cache refuses to store "other" because a shrug is not knowledge.
    This is the same rule for the per-template cache."""
    from evaluation.loop import _is_failed_parse

    assert _is_failed_parse({0: "other", 1: "other", 2: "other"})
    # survey-31's real parse is almost all "other" — and must NOT be discarded,
    # because the one real answer in it is knowledge.
    assert not _is_failed_parse({i: "other" for i in range(30)} | {30: "closing"})
    assert not _is_failed_parse({0: "title", 1: "bullet_list"})
    assert not _is_failed_parse({}), "an empty map is not the same claim"
