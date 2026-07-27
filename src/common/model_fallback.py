import time

import requests

# llm_clients.gigachat.API_BASE moved from gigachat.devices.sberbank.ru to
# api.giga.chat (2026-07) — that host serves a different model line entirely;
# the old names ("GigaChat", "GigaChat-Pro", ...) 404 there outright (verified
# directly against /v1/models and /v1/chat/completions, not assumed). 404 isn't
# in SKIPPABLE_STATUS_CODES, so a stale old name here doesn't even fail over
# to the next model — it raises immediately and takes the whole call down.
# GigaChat-2 (base tier) flatly rejects image attachments (422), same pattern
# as the old line's base tier — kept out of the vision chain for the same reason.
TEXT_MODELS = ["GigaChat-2", "GigaChat-2-Pro"]
VISION_MODELS = ["GigaChat-2-Pro", "GigaChat-2-Max"]

# Kept for backwards compatibility with existing imports.
DEFAULT_MODELS = TEXT_MODELS

SKIPPABLE_STATUS_CODES = {402, 422}

# Same flaky path as TRANSIENT_NETWORK_ERRORS below, but the failure arrives as
# an HTTP status instead of a socket error. Measured: a run died on
# "403 Forbidden" from /chat/completions while the token balance was untouched
# (250M base / 39M Pro / 25M Max) and a fresh call seconds later succeeded on
# the first try — so it was a momentary gateway rejection, not an auth or quota
# verdict. 403 is not skippable (skipping models wouldn't help — it isn't about
# the model) and re-raising killed the whole run, so it belongs here: retry the
# same model after a pause, bounded by retries_per_model.
TRANSIENT_STATUS_CODES = {403, 429, 500, 502, 503, 504}

# Sber's endpoints resolve through whatever local network path reaches Russia
# (observed resolving to addresses in the RFC 2544 benchmark range — i.e. some
# local proxy/tunnel, not a direct route) and that path drops connections
# intermittently: a request can fail with ConnectionError/SSLError and the very
# next attempt, seconds later, succeeds. This is a transient network fault, not
# a signal about the request itself — worth a short retry, unlike a real HTTP
# error from the server.
TRANSIENT_NETWORK_ERRORS = (
    requests.exceptions.ConnectionError,
    requests.exceptions.SSLError,
    requests.exceptions.Timeout,
)
NETWORK_RETRY_DELAY_SECONDS = 1.5


def call_with_model_fallback(call_fn, models, retries_per_model=5):
    """call_fn(model) -> result, where call_fn is expected to raise ValueError/KeyError
    if the model's response wasn't valid/parseable JSON. Tries each model in order:
    - 402/422 (out of quota / doesn't support the request) -> move to the next model
      immediately, retrying won't help.
    - malformed JSON (ValueError/KeyError) -> LLM output is occasionally malformed on a
      given sample; retry the same model a few times before giving up on it.
    - transient network fault (connection dropped, SSL EOF, timeout) -> same model,
      after a short pause — the network path itself needs a moment, not the model.

    If every model ultimately fails, raise whichever error actually reflects a model
    that was tried and produced bad output, rather than a "no quota" error from a
    model that never had a chance to begin with — that's the more useful message.
    """
    last_quota_error = None
    last_content_error = None
    last_network_error = None
    for model in models:
        for attempt in range(retries_per_model):
            try:
                return call_fn(model)
            except requests.HTTPError as e:
                code = e.response.status_code if e.response is not None else None
                if code in SKIPPABLE_STATUS_CODES:
                    last_quota_error = e
                    break
                if code in TRANSIENT_STATUS_CODES:
                    last_network_error = e
                    if attempt < retries_per_model - 1:
                        time.sleep(NETWORK_RETRY_DELAY_SECONDS)
                    continue
                raise
            except (ValueError, KeyError) as e:
                last_content_error = e
                continue
            except TRANSIENT_NETWORK_ERRORS as e:
                last_network_error = e
                if attempt < retries_per_model - 1:
                    time.sleep(NETWORK_RETRY_DELAY_SECONDS)
                continue
    raise last_content_error or last_network_error or last_quota_error
