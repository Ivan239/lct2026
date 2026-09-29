"""Backend registry for the improvement loop — "generate from different LLMs".

A backend is one place a model runs. GigaChat tiers all share one client;
RTX (and any OpenAI-compatible host) is configured by env. resolve() hands the
loop a ready (gen_client, model_name) pair, or raises BackendUnavailable so the
loop can QUEUE that model and move on — which is exactly RTX's state right now
(box offline), not an error.

Env for the open-weights backend (the one the brief requires — Apache 2.0 / MIT,
up to 35B — served by vLLM, Ollama, TGI or any OpenAI-compatible provider):
    LLM_BASE_URL   e.g. https://api.provider.tld/v1 or http://localhost:8000/v1
    LLM_MODEL      model name the server serves, e.g. Qwen/Qwen2.5-32B-Instruct
    LLM_API_KEY    optional

RTX_* are still read as aliases: that was the local GPU box's name during
development.
"""

import os

import requests

from llm_clients.gigachat import OAUTH_URL, GigaChatClient
from llm_clients.openai_compat import OpenAICompatClient

GIGACHAT_MODELS = ["GigaChat-2", "GigaChat-2-Pro", "GigaChat-2-Max", "GigaChat-3-Ultra"]


class BackendUnavailable(Exception):
    """The backend for this model exists in config but can't be reached now —
    queue the model, don't fail the run."""


_gigachat_singleton = None


def _gigachat():
    global _gigachat_singleton
    if _gigachat_singleton is None:
        _gigachat_singleton = GigaChatClient(verify_ssl=False)
    return _gigachat_singleton


# Short on purpose: this is a liveness probe, not the request.
REACH_TIMEOUT = 4


def _reachable(url, timeout=4):
    try:
        requests.get(url, timeout=timeout, verify=False)
        return True
    except requests.RequestException:
        return False


def open_weights_config():
    """The configured open-weights endpoint, or None when it is not set up."""
    base = os.environ.get("LLM_BASE_URL") or os.environ.get("RTX_BASE_URL")
    if not base:
        return None
    return {
        "base_url": base,
        "model": os.environ.get("LLM_MODEL") or os.environ.get("RTX_MODEL", "open"),
        "api_key": os.environ.get("LLM_API_KEY") or os.environ.get("RTX_API_KEY"),
        # Проверка TLS включена: у провайдеров (Cloud.ru и т.п.) сертификат
        # валидный. Выключать — только для самоподписанного локального vLLM.
        "verify_ssl": os.environ.get("LLM_VERIFY_SSL", "1").lower() not in ("0", "false", "no"),
    }


def open_weights_client():
    """(client, model_name) for the configured open-weights endpoint, or None.

    The default provider of the service: the brief allows only open weights up
    to 35B under Apache 2.0 / MIT. GigaChat stays reachable as a development
    fallback and is never used when this endpoint is configured."""
    cfg = open_weights_config()
    if not cfg:
        return None
    client = OpenAICompatClient(cfg["base_url"], api_key=cfg["api_key"],
                                default_model=cfg["model"], verify_ssl=cfg["verify_ssl"])
    return client, cfg["model"]


_rtx_config = open_weights_config  # исторический псевдоним


def resolve(model):
    """Return (gen_client, model_name) for `model`, or raise BackendUnavailable.
    `model` is a GigaChat tier name, or "rtx" / "rtx:<name>" for the RTX box."""
    if model in GIGACHAT_MODELS:
        # Same treatment RTX already gets: an unreachable backend is "queued",
        # not a run that dies ten minutes later. The OAuth host has been
        # unreachable for 34 consecutive loop iterations, and each one spent
        # 5-16 minutes working through connect timeouts and the model fallback
        # chain before failing — a probe answers in 3 seconds. On a reachable
        # host nothing changes: the probe passes and the real request follows.
        if not _reachable(OAUTH_URL, timeout=REACH_TIMEOUT):
            raise BackendUnavailable(
                f"GigaChat недоступен по сети: {OAUTH_URL} не отвечает")
        return _gigachat(), model

    if model == "rtx" or model.startswith("rtx:"):
        cfg = _rtx_config()
        if not cfg:
            raise BackendUnavailable("RTX не настроен: задайте RTX_BASE_URL (сейчас в очереди)")
        if not _reachable(cfg["base_url"]):
            raise BackendUnavailable(f"RTX недоступен по {cfg['base_url']} (в очереди до появления)")
        name = model.split(":", 1)[1] if ":" in model else cfg["model"]
        client = OpenAICompatClient(cfg["base_url"], api_key=cfg["api_key"],
                                    default_model=name, verify_ssl=False)
        return client, name

    raise BackendUnavailable(f"неизвестный бэкенд для модели {model!r}")


def available(models):
    """Split a candidate model list into (ready, queued) — queued are configured
    but unreachable right now (e.g. RTX offline). Used to report loop status."""
    ready, queued = [], []
    for m in models:
        try:
            resolve(m)
            ready.append(m)
        except BackendUnavailable as e:
            queued.append((m, str(e)))
    return ready, queued
