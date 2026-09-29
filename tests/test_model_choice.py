"""Модель из запроса — только та, что есть у провайдера.

Браузер хранил последний выбор в localStorage, и «GigaChat-2-Max» со времён
разработки уходил в Cloud.ru: 404 на плане, 500 и «Не удалось разобрать бриф
через GigaChat» — у человека, который записывал видео."""

from api import main as api_main


def test_an_unknown_model_falls_back_to_the_configured_one(monkeypatch):
    monkeypatch.setattr(api_main, "DEFAULT_MODEL", "Qwen/Qwen3-32B")
    monkeypatch.setattr(api_main, "DEFAULT_MODELS", ["Qwen/Qwen3-32B"])
    assert api_main._models_for("GigaChat-2-Max") == ["Qwen/Qwen3-32B"]
    assert api_main._models_for("Qwen/Qwen3-32B") == ["Qwen/Qwen3-32B"]
    assert api_main._models_for(None) == ["Qwen/Qwen3-32B"]


def test_without_a_configured_model_the_request_decides(monkeypatch):
    monkeypatch.setattr(api_main, "DEFAULT_MODEL", None)
    assert api_main._models_for("GigaChat-2-Max") == ["GigaChat-2-Max"]
