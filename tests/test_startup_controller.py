"""Юнит-тесты StartupFlow — START-1 (R-07): стартовые задачи в рабочем потоке.

Контракты карточки (этап 5 roadmap rev2):

* guard повторного входа: второй run() не порождает второй воркер;
* воркер выполняет задачи по порядку (очистки → план корзины → сверка
  папки → сбор сроков), кладёт результаты в СВОЮ очередь и завершается
  ("done"); между шагами и файлами проверяется флаг остановки;
* результаты применяются ТОЛЬКО в главном потоке (drain из tick):
  очистка корзины — мутация главного потока (TRASH-WARN), сверка —
  через scan_consistency(orphans=…), уведомление — через
  _apply_expiry_notice; клики по реестру во время сверки обрабатываются
  (главный поток свободен — критерий карточки);
* include_notice=False: сбор сроков не выполняется (суточный гейт);
* при завершении работы работник останавливается по флагу.
"""
from __future__ import annotations

import logging
import os
import queue
from types import SimpleNamespace

import pytest

import startup_controller as sc
import storage
from document_builder import F


class FakeRoot:
    def __init__(self):
        self.after_calls: list = []

    def after(self, ms, fn):
        self.after_calls.append((ms, fn))


class FakeSettingsFlow:
    def __init__(self, documents_dir: str):
        self._documents_dir = documents_dir

    def documents_dir(self, interactive: bool = True) -> str:
        return self._documents_dir


class FakeRegistryFlow:
    def __init__(self):
        self.calls: list = []

    def scan_consistency(self, orphans=None):
        self.calls.append(orphans)


class FakeApp:
    def __init__(self, store, documents_dir: str):
        self.store = store
        self.root = FakeRoot()
        self._log = logging.getLogger("startup-tests")
        self._settings_flow = FakeSettingsFlow(documents_dir)
        self._registry_flow = FakeRegistryFlow()
        self.trash_calls: list = []
        self.notice_calls: list = []

    def _cleanup_expired_trash(self, planned=None):
        self.trash_calls.append(planned)

    def _apply_expiry_notice(self, text):
        self.notice_calls.append(text)


@pytest.fixture()
def fake_app(store, tmp_path):
    docs_dir = tmp_path / "docs"
    docs_dir.mkdir()
    return FakeApp(store, str(docs_dir))


def _sync_factory(worker_calls: list):
    """Синхронный «поток»: воркер выполняется сразу в вызывающем потоке."""
    def factory(target, args=(), daemon=True):
        def run():
            worker_calls.append(args)
            target(*args)
        return SimpleNamespace(start=run)
    return factory


def _add_expired_agreement(store):
    party = store.add_counterparty(
        "ooo", {F.INN: "7701234567", F.SHORT_NAME: "ООО Ромашка"}, save=False,
    )
    store.add_document(
        storage.KIND_ADDITIONAL, party["id"], "1", "01.01.2026",
        data={}, extra={"valid_until": "01.01.2020"}, text="", save=True,
    )


def _drain_all(flow: sc.StartupFlow):
    """Дренирует очередь воркера до 'done' (без Tk-цикла)."""
    assert flow._drain() is True
    flow.running = False


# --------------------------------------------------------------------------- #
#  Порядок задач, применение результатов в главном потоке
# --------------------------------------------------------------------------- #
def test_worker_emits_all_results_and_done(fake_app, tmp_path):
    _add_expired_agreement(fake_app.store)
    (tmp_path / "docs" / "мусор.docx.tmp").write_text("x")
    flow = sc.StartupFlow(fake_app, thread_factory=lambda **kw: SimpleNamespace(
        target=None, args=None, start=None,
    ))
    flow._thread_factory = None  # не используется: worker вызывается напрямую
    result_queue = queue.Queue()
    flow.worker(result_queue)
    kinds = []
    while True:
        kind, payload = result_queue.get_nowait()
        if kind == "done":
            break
        kinds.append((kind, payload))
    assert [kind for kind, _ in kinds] == ["temp", "trash_plan", "orphans", "notice"]
    assert dict(kinds)["temp"] == 1  # .docx.tmp удалён воркером
    assert dict(kinds)["trash_plan"] == 0
    assert dict(kinds)["orphans"] == []
    assert dict(kinds)["notice"] and "истекло" in dict(kinds)["notice"]


def test_apply_applies_results_in_main_thread(fake_app, tmp_path):
    _add_expired_agreement(fake_app.store)
    flow = sc.StartupFlow(fake_app)
    flow.running = True
    flow._queue = queue.Queue()
    flow._emit(flow._queue, "temp", 3)
    flow._emit(flow._queue, "trash_plan", 2)
    flow._emit(flow._queue, "orphans", ["c:/файл.docx"])
    flow._emit(flow._queue, "notice", "Просроченные доп. соглашения:\n• …")
    flow._emit(flow._queue, "done", None)
    assert flow._drain() is True
    assert fake_app.trash_calls == [2]  # мутация корзины — главный поток
    assert fake_app._registry_flow.calls == [["c:/файл.docx"]]
    assert fake_app.notice_calls == ["Просроченные доп. соглашения:\n• …"]


def test_apply_skips_empty_results(fake_app):
    flow = sc.StartupFlow(fake_app)
    flow.running = True
    flow._queue = queue.Queue()
    flow._emit(flow._queue, "temp", 0)
    flow._emit(flow._queue, "trash_plan", 0)
    flow._emit(flow._queue, "orphans", [])
    flow._emit(flow._queue, "notice", None)
    flow._emit(flow._queue, "done", None)
    assert flow._drain() is True
    assert fake_app.trash_calls == []
    assert fake_app._registry_flow.calls == [[]]  # сверка применяется даже без сирот
    assert fake_app.notice_calls == [None]  # None доходит до приложения и игнорируется


def test_run_guard_and_scheduling(fake_app):
    calls: list = []
    flow = sc.StartupFlow(fake_app, thread_factory=lambda target, args, daemon: SimpleNamespace(
        target=target, args=args, start=lambda: calls.append(args),
    ))
    assert flow.run(include_notice=False) is True
    assert flow.running is True
    assert flow.run() is False  # guard повторного входа
    assert len(calls) == 1
    assert fake_app.root.after_calls and fake_app.root.after_calls[0][0] == 80


def test_run_without_notice_skips_expiry_collection(fake_app):
    calls: list = []

    def factory(target, args, daemon=True):
        def run():
            calls.append(args)
            target(*args)
        return SimpleNamespace(start=run)

    flow = sc.StartupFlow(fake_app, thread_factory=factory)
    _add_expired_agreement(fake_app.store)
    flow.run(include_notice=False)
    _drain_all(flow)
    # Сбор сроков не выполнялся: уведомление не применялось.
    assert fake_app.notice_calls == [None]
    assert fake_app._registry_flow.calls == [[]]


def test_stop_flag_interrupts_worker_between_steps(fake_app):
    _add_expired_agreement(fake_app.store)
    flow = sc.StartupFlow(fake_app)
    flow.stop()
    result_queue = queue.Queue()
    flow.worker(result_queue)
    kinds = []
    while True:
        kind, _payload = result_queue.get_nowait()
        if kind == "done":
            break
        kinds.append(kind)
    # Очистки успели выполниться (первый шаг), остальное прервано флагом.
    assert kinds == ["temp"]


def test_worker_swallows_scan_errors(fake_app, tmp_path):
    fake_app._settings_flow = FakeSettingsFlow("")  # папка не задана
    flow = sc.StartupFlow(fake_app)
    result_queue = queue.Queue()
    flow.worker(result_queue)
    payloads = {}
    while True:
        kind, payload = result_queue.get_nowait()
        if kind == "done":
            break
        payloads[kind] = payload
    assert payloads["orphans"] == []
    assert payloads["notice"] is None  # сроков нет — тишина


# --------------------------------------------------------------------------- #
#  Очистка .docx.tmp (чистый диск, без реестра)
# --------------------------------------------------------------------------- #
def test_cleanup_temp_documents_removes_only_tmp(fake_app, tmp_path):
    (tmp_path / "docs" / "подпапка").mkdir()
    (tmp_path / "docs" / "a.docx.tmp").write_text("x")
    (tmp_path / "docs" / "подпапка" / "b.docx.tmp").write_text("x")
    (tmp_path / "docs" / "документ.docx").write_text("x")
    flow = sc.StartupFlow(fake_app)
    assert flow._cleanup_temp_documents() == 2
    assert not (tmp_path / "docs" / "a.docx.tmp").exists()
    assert (tmp_path / "docs" / "документ.docx").exists()


def test_cleanup_temp_documents_missing_dir(fake_app):
    fake_app._settings_flow = FakeSettingsFlow(str(fake_app.store) + "-нет-такой")
    flow = sc.StartupFlow(fake_app)
    assert flow._cleanup_temp_documents() == 0


def test_worker_uses_real_thread_and_queue(fake_app):
    """Интеграционный вариант: настоящий daemon-поток + drain главного потока."""
    flow = sc.StartupFlow(fake_app)  # фабрика по умолчанию — threading.Thread
    assert flow.run(include_notice=False) is True
    deadline_queue = flow._queue
    assert deadline_queue is not None
    # Воркер в фоне; главный поток дренирует до 'done'.
    import time
    for _ in range(200):
        if flow._drain():
            break
        time.sleep(0.005)
    flow.running = False
    assert fake_app._registry_flow.calls  # сверка применена в главном потоке


# --------------------------------------------------------------------------- #
#  E-4 (контрольные замеры этапа 5): перф-каналы старта
# --------------------------------------------------------------------------- #
def test_startup_steps_write_perf_records(fake_app, monkeypatch, tmp_path):
    """EWDOCHUB_PERF_LOG=1: шаги воркера и применения пишут каналы startup.*.

    Воркер — информационные каналы; применения — главнопоточные, они
    участвуют в критерии E-4 «первые 5 секунд без операций > 100 мс».
    Имена источников однотоковые латинские: парсер bench_store_1a
    --report требует ровно 5 токенов в строке perf-журнала. Без флага
    записи не создаются (контракт DIAG-1 — проверен в test_perflog).
    """
    import perflog
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    try:
        _add_expired_agreement(fake_app.store)
        (tmp_path / "docs" / "мусор.docx.tmp").write_text("x")
        flow = sc.StartupFlow(fake_app)
        result_queue = queue.Queue()
        flow.worker(result_queue)      # каналы шагов воркера
        flow.running = True
        flow._queue = result_queue
        _drain_all(flow)               # каналы применений (главный поток)

        # Буфер общий для сеанса: кроме startup.* в нём могут быть записи
        # других timed-каналов (например, «Запись реестра» от хранилища
        # фикстуры) — фильтруем только стартовые.
        startup_sources, startup_kinds = set(), set()
        for line in perflog._buffer:
            tokens = line.split()
            if len(tokens) >= 5 and tokens[3].startswith("startup."):
                assert len(tokens) == 5, "источники startup.* однотоковые"
                startup_kinds.add(tokens[2])
                startup_sources.add(tokens[3])
        assert {"startup.temp", "startup.trash_count", "startup.orphans",
                "startup.notice", "startup.apply_scan"} <= startup_sources
        assert startup_kinds == {"duration"}, "события startup.built/ready добавляет ContractApp.run"
    finally:
        perflog.reset_for_tests()


def test_startup_apply_trash_perf_record(fake_app, monkeypatch):
    """Мутация корзины (payload > 0) пишет главнопоточный канал apply_trash."""
    import perflog
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    try:
        flow = sc.StartupFlow(fake_app)
        flow.running = True
        flow._queue = queue.Queue()
        flow._emit(flow._queue, "trash_plan", 2)
        flow._emit(flow._queue, "done", None)
        assert flow._drain() is True
        assert fake_app.trash_calls == [2]
        sources = {line.split()[3] for line in perflog._buffer}
        assert "startup.apply_trash" in sources
    finally:
        perflog.reset_for_tests()
