"""Backend registry for the improvement loop — "generate from different LLMs".

A backend is one place a model runs. GigaChat tiers all share one client;
RTX (and any OpenAI-compatible host) is configured by env. resolve() hands the
loop a ready (gen_client, model_name) pair, or raises BackendUnavailable so the
loop can QUEUE that model and move on — which is exactly RTX's state right now
(box offline), not an error.

Env for an OpenAI-compatible backend (e.g. RTX):
    RTX_BASE_URL   e.g. http://192.168.31.50:8000/v1
    RTX_MODEL      model name the server serves (default "rtx")
    RTX_API_KEY    optional
"""

import os

import requests

from llm_clients.gigachat import GigaChatClient
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


def _reachable(url, timeout=4):
    try:
        requests.get(url, timeout=timeout, verify=False)
        return True
    except requests.RequestException:
        return False


def _rtx_config():
    base = os.environ.get("RTX_BASE_URL")
    if not base:
        return None
    return {
        "base_url": base,
        "model": os.environ.get("RTX_MODEL", "rtx"),
        "api_key": os.environ.get("RTX_API_KEY"),
    }


def resolve(model):
    """Return (gen_client, model_name) for `model`, or raise BackendUnavailable.
    `model` is a GigaChat tier name, or "rtx" / "rtx:<name>" for the RTX box."""
    if model in GIGACHAT_MODELS:
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
