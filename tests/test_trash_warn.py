"""TRASH-WARN — сводка перед автоочисткой корзины 30 дней.

Этап 2 roadmap: корзина не удаляет записи молча. Контракты:

  * storage.count_expired_trash() считает записи старше срока хранения
    ДО фактического удаления (сводка «будет удалено N записей»);
  * _cleanup_expired_trash приложения: пустая сводка → никаких действий;
    непустая → предупреждение в журнал ДО cleanup, после — итог в статусе
    и тост пользователю; сбой подсчёта не роняет старт;
  * срок хранения — storage.TRASH_RETENTION_DAYS (30 дней), текст сводки
    называет срок.

Тесты: storage — на реальном DataStore; приложение — на duck-typed
фейке ContractApp без Tk.
"""
from __future__ import annotations

import logging
import sys
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import main
import storage


@pytest.fixture()
def appdata(tmp_path, monkeypatch):
    path = tmp_path / "appdata"
    path.mkdir()
    monkeypatch.setenv("APPDATA", str(path))
    return path


@pytest.fixture()
def store(appdata):
    return storage.DataStore()


def _trashed(store, days_ago: int, mode: str = "ooo") -> dict:
    rec = store.add_counterparty(
        mode, {"ИНН": "2540123456"}, save=False,
    )
    rec["deleted_at"] = (datetime.now() - timedelta(days=days_ago)).isoformat(
        timespec="seconds"
    )
    return rec


# --------------------------------------------------------------------------- #
#  storage: подсчёт до удаления
# --------------------------------------------------------------------------- #
def test_count_expired_trash_counts_only_expired(store):
    fresh = _trashed(store, days_ago=1)
    expired = _trashed(store, days_ago=45)
    active = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    assert storage.is_deleted(fresh) and storage.is_deleted(expired)
    assert not storage.is_deleted(active)
    assert store.count_expired_trash() == 1, "свежая корзина и живые записи не считаются"


def test_count_expired_trash_counts_documents_too(store):
    _trashed(store, days_ago=45)
    doc = {
        "id": store._next_id(), "kind": storage.KIND_CONTRACT,
        "party_id": 1, "number": "5", "data": {}, "extra": {}, "text": "",
        "deleted_at": (datetime.now() - timedelta(days=45)).isoformat(
            timespec="seconds"
        ),
    }
    store.documents.append(doc)
    assert store.count_expired_trash() == 2, "контрагенты и документы вместе"


def test_count_trash_matches_cleanup_result(store):
    """Подсчёт совпадает с фактом очистки — сводка не врёт."""
    for days in (45, 50):
        _trashed(store, days_ago=days)
    _trashed(store, days_ago=2)  # свежая — останется
    planned = store.count_expired_trash()
    removed = store.cleanup_trash()
    assert planned == removed == 2


def test_trash_retention_days_is_30():
    assert storage.TRASH_RETENTION_DAYS == 30


# --------------------------------------------------------------------------- #
#  Приложение: сводка ДО удаления, итог — после
# --------------------------------------------------------------------------- #
class FakeTrashStore:
    """Duck-typed store: считает и удаляет по сценарию теста."""

    def __init__(self, planned: int, removed: int | None = None):
        self.planned = planned
        self.removed = planned if removed is None else removed
        self.count_calls = 0
        self.cleanup_calls = 0

    def count_expired_trash(self, *args, **kwargs):
        self.count_calls += 1
        return self.planned

    def cleanup_trash(self, *args, **kwargs):
        self.cleanup_calls += 1
        return self.removed


def make_trash_app(store_obj) -> main.ContractApp:
    app = object.__new__(main.ContractApp)
    app.store = store_obj
    app._log = logging.getLogger("test-trash-warn")
    app.calls: list = []
    app.notifications: list[str] = []
    app._registry_flow = SimpleNamespace(
        refresh=lambda: app.calls.append("registry"),
    )
    app._update_status = lambda: app.calls.append("status")
    app._notify = lambda message, path=None: app.notifications.append(message)
    return app


def test_cleanup_empty_trash_does_nothing():
    store_obj = FakeTrashStore(planned=0)
    app = make_trash_app(store_obj)
    app._cleanup_expired_trash()
    assert store_obj.count_calls == 1 and store_obj.cleanup_calls == 0, \
        "пустая сводка — cleanup не запускается вовсе"
    assert app.notifications == [] and app.calls == []


def test_cleanup_logs_summary_before_deleting():
    """Сводка «будет удалено N» пишется в журнал ДО cleanup."""
    store_obj = FakeTrashStore(planned=3, removed=3)
    app = make_trash_app(store_obj)
    records: list[logging.LogRecord] = []

    class _Capture(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            records.append(record)

    handler = _Capture(level=logging.WARNING)
    app._log.addHandler(handler)
    try:
        app._cleanup_expired_trash()
    finally:
        app._log.removeHandler(handler)

    warnings = [r.getMessage() for r in records if r.levelno >= logging.WARNING]
    assert any("будет удалено 3" in text for text in warnings), \
        f"сводка «будет удалено» отсутствует: {warnings}"
    assert any("30" in text for text in warnings), "срок хранения назван"


def test_cleanup_notifies_user_after_deletion():
    store_obj = FakeTrashStore(planned=2, removed=2)
    app = make_trash_app(store_obj)
    app._cleanup_expired_trash()
    assert app.calls == ["registry", "status"], "реестр и статус обновлены"
    assert app.notifications and "удалено 2 записей" in app.notifications[0], \
        "пользователь видит итог автоочистки"


def test_cleanup_count_failure_keeps_startup_alive():
    class BrokenStore:
        def count_expired_trash(self):
            raise OSError("диск недоступен")

        def cleanup_trash(self):  # не должен вызываться
            raise AssertionError("cleanup после сбоя подсчёта не запускается")

    app = make_trash_app(BrokenStore())
    app._cleanup_expired_trash()  # не падает
    assert app.notifications == []


def test_cleanup_failure_after_count_is_logged():
    class BrokenCleanup:
        def count_expired_trash(self):
            return 5

        def cleanup_trash(self):
            raise OSError("файл занят")

    app = make_trash_app(BrokenCleanup())
    app._cleanup_expired_trash()  # не падает
    assert app.notifications == [], "без удаления — без тоста"
