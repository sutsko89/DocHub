"""Perf-журнал производительности (DIAG-1): флаг по умолчанию выключен.

Контракты карточки DIAG-1:
  * флаг EWDOCHUB_PERF_LOG по умолчанию выключен — ни файла, ни записей;
  * при включённом флаге записи буферизуются и пишутся в logs/perf.log;
  * applog.timed дублирует точки в perf-журнал без порога 500 мс,
    а app.log продолжается вести себя как раньше;
  * при выключенном флаге поведение идентично прежнему.
"""
from __future__ import annotations

import logging
import os
import time

import applog
import perflog


def test_flag_disabled_by_default(appdata, monkeypatch):
    monkeypatch.delenv(perflog.ENV_FLAG, raising=False)
    perflog.reset_for_tests()
    assert perflog.enabled() is False


def test_flag_enabled_from_env(appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    assert perflog.enabled() is True
    assert perflog.perf_path().endswith("perf.log")


def test_disabled_record_writes_nothing(appdata, monkeypatch):
    monkeypatch.delenv(perflog.ENV_FLAG, raising=False)
    perflog.reset_for_tests()
    perflog.record("ui.refresh", 123.0)
    perflog.flush()
    assert not os.path.exists(perflog.perf_path()), "файл не создаётся без флага"


def test_enabled_record_writes_buffer(appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    perflog.record("ui.refresh", 12.5, kind="duration")
    perflog.record("pointer.poller", 16.2, kind="tick")
    perflog.flush()
    with open(perflog.perf_path(), "r", encoding="utf-8") as fh:
        lines = fh.read().strip().splitlines()
    assert len(lines) == 2
    assert " duration ui.refresh 12.500" in lines[0]
    assert " tick pointer.poller 16.200" in lines[1]


def test_record_flushes_automatically(appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    for _ in range(perflog.FLUSH_EVERY + 1):
        perflog.record("pointer.poller", 1.0, kind="tick")
    # Буфер очищается порциями — в файле не меньше FLUSH_EVERY строк.
    with open(perflog.perf_path(), "r", encoding="utf-8") as fh:
        assert len(fh.read().strip().splitlines()) >= perflog.FLUSH_EVERY


def test_timed_duplicates_into_perf_log(appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    logger = logging.getLogger("test-perflog")
    with applog.timed("Быстрая операция", 500, "test"):
        pass  # быстрее порога — в app.log не попадает
    perflog.flush()
    with open(perflog.perf_path(), "r", encoding="utf-8") as fh:
        content = fh.read()
    assert "duration Быстрая операция" in content, (
        "timed-точка дублируется в perf-журнал без порога 500 мс"
    )


def test_timed_below_threshold_keeps_applog_quiet(appdata, monkeypatch):
    monkeypatch.delenv(perflog.ENV_FLAG, raising=False)
    perflog.reset_for_tests()
    logger = logging.getLogger(applog.LOGGER_NAME + ".test")
    records: list[logging.LogRecord] = []
    handler = logging.Handler()
    handler.emit = records.append
    logger.addHandler(handler)
    with applog.timed("Быстрая операция", 500, "test"):
        pass
    assert not records, "порог app.log не изменён (500 мс)"


def test_tick_meter_measures_intervals(appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    meter = perflog.TickMeter("unit.meter")
    meter.tick()
    time.sleep(0.001)
    meter.tick()
    perflog.flush()
    with open(perflog.perf_path(), "r", encoding="utf-8") as fh:
        lines = [line for line in fh.read().strip().splitlines() if "unit.meter" in line]
    assert len(lines) == 1, "первый тик — только отсчёт, dt пишется со второго"


def test_canvas_meter_not_bound_when_disabled(appdata, monkeypatch):
    class FakeWidget:
        def __init__(self):
            self.bindings = []

        def bind(self, sequence, handler, add=""):
            self.bindings.append((sequence, add))

    monkeypatch.delenv(perflog.ENV_FLAG, raising=False)
    perflog.reset_for_tests()
    widget = FakeWidget()
    assert perflog.bind_canvas_meter(widget, "registry.canvas") is None
    assert widget.bindings == [], "без флага обработчиков нет вовсе"


def test_canvas_meter_binds_when_enabled(appdata, monkeypatch):
    class FakeWidget:
        def __init__(self):
            self.bindings = []

        def bind(self, sequence, handler, add=""):
            self.bindings.append((sequence, add))
            self.handler = handler

    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    widget = FakeWidget()
    meter = perflog.bind_canvas_meter(widget, "registry.canvas")
    assert meter is not None
    assert widget.bindings == [("<Configure>", "+")]
    widget.handler()
    widget.handler()
    perflog.flush()
    with open(perflog.perf_path(), "r", encoding="utf-8") as fh:
        assert "tick registry.canvas" in fh.read()
