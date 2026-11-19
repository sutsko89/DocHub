"""Уборка .bad-* копий повреждённого реестра (аудит 2.0.1, P3) и надёжность
записи базы (п. 11 этапа 5).

Свежие .bad-копии не трогаются, старше BAD_COPY_RETENTION_DAYS — удаляются
при загрузке реестра. Долговечность записи (сбой питания) в SQLite
гарантирует PRAGMA synchronous=FULL.
"""
from __future__ import annotations

import os
import sqlite3
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


def test_synchronous_full_for_power_failure_safety(appdata):
    """synchronous=FULL: коммит доходит до диска до подтверждения записи.

    Прежний тест проверял fsync у .bak-копии JSON; при записи транзакцией
    в SQLite ту же гарантию даёт PRAGMA synchronous=FULL.
    """
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    assert store.save() is True
    conn = sqlite3.connect(storage.db_path())
    try:
        mode = conn.execute("PRAGMA synchronous").fetchone()[0]
    finally:
        conn.close()
    assert mode == 2, "synchronous=FULL (2): запись переживает сбой питания"
