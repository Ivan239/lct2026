"""Прогресс трёх параллельных вариантов не перемешивается.

Одна общая ячейка на три потока давала в кнопке то «11/11», то «10/14», а
первый закончивший вариант гасил прогресс двух других."""

import threading

import api.main as api_main


def _in_lane(name, fn):
    t = threading.Thread(target=lambda: (setattr(api_main._lane, "name", name), fn()))
    t.start()
    t.join()


def test_each_variant_keeps_its_own_counter():
    api_main._progress_lanes.clear()
    _in_lane("compact", lambda: api_main._set_progress("Пишем контент: а", done=2, total=10))
    _in_lane("detailed", lambda: api_main._set_progress("Пишем контент: б", done=5, total=14))
    _in_lane("visual", lambda: api_main._set_progress("Собираем .pptx в стиле шаблона"))
    snap = api_main._progress_snapshot()
    assert snap["active"] and snap["total"] == 0
    assert "Компактный 3/10" in snap["stage"] and "Подробный 6/14" in snap["stage"]
    assert "Визуальный: собираем .pptx" in snap["stage"], snap["stage"]


def test_a_finished_variant_does_not_stop_the_others():
    api_main._progress_lanes.clear()
    _in_lane("compact", lambda: api_main._set_progress("Пишем контент: а", done=1, total=10))
    _in_lane("visual", lambda: api_main._set_progress("Пишем контент: в", done=1, total=11))
    _in_lane("compact", lambda: api_main._set_progress("", active=False))
    snap = api_main._progress_snapshot()
    assert snap["active"] and snap["stage"] == "Пишем контент: в" and snap["total"] == 11
    api_main._progress_lanes.clear()
    assert not api_main._progress_snapshot()["active"]
