"""429 от провайдера — подождать и повторить, а не уронить вариант.

Три варианта вёрстки идут параллельно, Cloud.ru отвечал «Too Many Requests»,
а пять попыток по 1,5 с (7 секунд) были короче окна лимита — вариант «не
собрался» прямо на записи демо."""

import requests

from common import model_fallback as mf


def _http_error(code, retry_after=None):
    response = requests.Response()
    response.status_code = code
    if retry_after is not None:
        response.headers["Retry-After"] = str(retry_after)
    return requests.HTTPError(f"{code}", response=response)


def test_rate_limit_waits_longer_and_keeps_the_content_retries(monkeypatch):
    sleeps = []
    monkeypatch.setattr(mf.time, "sleep", sleeps.append)
    answers = [_http_error(429)] * 4 + [ValueError("битый JSON")] * 4 + ["ok"]

    def call(model):
        answer = answers.pop(0)
        if isinstance(answer, Exception):
            raise answer
        return answer

    assert mf.call_with_model_fallback(call, ["m"], retries_per_model=5) == "ok"
    assert sleeps[:4] == [2, 4, 8, 16], sleeps
    assert sum(sleeps) >= 30, "ожидание короче окна лимита"


def test_retry_after_from_the_provider_is_honoured(monkeypatch):
    sleeps = []
    monkeypatch.setattr(mf.time, "sleep", sleeps.append)
    answers = [_http_error(429, retry_after=7), "ok"]

    def call(model):
        answer = answers.pop(0)
        if isinstance(answer, Exception):
            raise answer
        return answer

    assert mf.call_with_model_fallback(call, ["m"]) == "ok"
    assert sleeps == [7.0]


def test_an_endless_rate_limit_still_ends(monkeypatch):
    monkeypatch.setattr(mf.time, "sleep", lambda s: None)
    calls = []

    def call(model):
        calls.append(model)
        raise _http_error(429)

    try:
        mf.call_with_model_fallback(call, ["m"], retries_per_model=3)
    except requests.HTTPError:
        pass
    else:
        raise AssertionError("бесконечный 429 должен закончиться ошибкой")
    assert len(calls) < 20
