"""Уборка .bad-* копий повреждённого реестра (аудит 2.0.1, P3).

Свежие копии не трогаются, старше BAD_COPY_RETENTION_DAYS — удаляются
при загрузке реестра.
"""
from __future__ import annotations

import os
import time

import storage


def _create(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("{}")


def test_stale_bad_copies_removed_fresh_kept(appdata):
    path = storage.db_path()
    old = path + ".bad-20200101-000000"
    fresh = path + ".bad-29991231-235959"
    _create(old)
    _create(fresh)
    past = time.time() - storage.BAD_COPY_RETENTION_DAYS * 86400 - 3600
    os.utime(old, (past, past))

    storage.DataStore()  # _load() запускает уборку

    assert not os.path.exists(old)
    assert os.path.exists(fresh)


def test_young_bad_copies_are_kept(appdata):
    path = storage.db_path()
    young = path + ".bad-20990101-000000"
    _create(young)
    recent = time.time() - 3600  # час назад — в пределах срока хранения
    os.utime(young, (recent, recent))

    storage.DataStore()

    assert os.path.exists(young)


def test_backup_file_gets_fsync_call(appdata, monkeypatch):
    """fsync у .bak-копии вызывается (аудит: бэкап обязан пережить сбой питания)."""
    calls: list[int] = []
    real_fsync = os.fsync

    def spy_fsync(fd):
        calls.append(fd)
        return real_fsync(fd)

    monkeypatch.setattr(os, "fsync", spy_fsync)
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    assert store.save() is True
    assert calls, "os.fsync вызван при ротации бэкапов"
    assert os.path.exists(storage.db_path() + ".bak")
