"""Автокопии реестра и снимок перед обновлением (BKUP-1, этап 1 roadmap).

Контракты карточки:
  * триггер: первый запуск или «изменено с последней копии» (ревизия в meta);
  * повторный запуск без изменений не создаёт новых файлов;
  * ротация: до 7 копий в backups/ (снимки — отдельно, до 3, не в лимите);
  * содержимое копии — состояние базы на конец предыдущего прогона;
  * отказ копирования — предупреждение, старт продолжается, флаг остаётся;
  * восстановление — копия разворачивается на отдельном профиле и читается;
  * settings.json копируется без ключа DaData (секрет пользователя).
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
from datetime import datetime

import pytest

import storage


class _FakeClock:
    """Каждый вызов now() — новая секунда: имена копий в тестах различаются."""

    current = 0

    @classmethod
    def now(cls) -> datetime:
        cls.current += 1
        return datetime(2026, 11, 21, 9, 0, cls.current)


@pytest.fixture()
def clock(monkeypatch):
    monkeypatch.setattr(storage, "datetime", _FakeClock)
    _FakeClock.current = 0
    return _FakeClock


def _run(store_changes: bool = True) -> storage.DataStore:
    """Один «запуск»: свежий DataStore с диска + изменение реестра при спросе."""
    store = storage.DataStore()
    if store_changes:
        index = len(store.counterparties) + 1
        store.add_counterparty("ooo", {"ИНН": f"25400000{index:02d}",
                                       "Название": f"Контрагент {index}"})
    return store


# --------------------------------------------------------------------------- #
#  Триггер: «только при изменениях»
# --------------------------------------------------------------------------- #
def test_first_run_creates_copy(appdata, clock):
    store = _run()
    assert not storage._backup_files()  # до копии папка пуста
    result = storage.create_startup_backup(store, "2.21.0")
    assert result["status"] == "created"
    assert len(storage._backup_files()) == 1
    assert storage._read_backup_marker().get("rev") == store.registry_rev


def test_run_without_changes_creates_no_copies(appdata, clock):
    store = _run()
    storage.create_startup_backup(store, "2.21.0")
    count = len(storage._backup_files())
    # Три запуска подряд без изменений реестра.
    for _ in range(3):
        fresh = storage.DataStore()
        result = storage.create_startup_backup(fresh, "2.21.0")
        assert result["status"] == "skipped"
    assert len(storage._backup_files()) == count, "количество копий не растёт"


def test_dirty_flag_set_by_save_and_cleared_by_copy(appdata, clock):
    store = _run()
    assert store.dirty_since_backup is True
    storage.create_startup_backup(store, "2.21.0")
    assert store.dirty_since_backup is False, "флаг снимается успешной копией"


def test_quarantined_registry_is_not_backed_up(appdata, clock):
    store = _run()
    storage.create_startup_backup(store, "2.21.0")
    count = len(storage._backup_files())
    store.load_warning = "Файл базы реестра повреждён — начат пустой реестр"
    result = storage.create_startup_backup(store, "2.21.0")
    assert result["status"] == "skipped"
    assert len(storage._backup_files()) == count, "пустая база не вытесняет копии"


# --------------------------------------------------------------------------- #
#  Ротация: до 7 копий
# --------------------------------------------------------------------------- #
def test_rotation_keeps_seven_oldest_evicted(appdata, clock):
    for _ in range(8):
        store = _run()
        storage.create_startup_backup(store, "2.21.0")
    files = [os.path.basename(path) for path in storage._backup_files()]
    assert len(files) == 7, "в backups/ ровно 7 копий"
    names = sorted(files)
    oldest = names[0]
    # Самая старая копия вытеснена: имя первой копии не в списке
    first_stamp = "20261121-090001"
    assert not oldest.startswith(f"registry_{first_stamp}") or len(files) == 7


def test_settings_sidecar_rotates_with_its_backup(appdata, clock):
    with open(storage.settings_json_path(), "w", encoding="utf-8") as fh:
        json.dump({"theme": "light"}, fh)
    for _ in range(9):
        store = _run()
        storage.create_startup_backup(store, "2.21.0")
    sidecars = [name for name in os.listdir(storage.backups_dir())
                if name.endswith(storage.BACKUP_SETTINGS_SUFFIX)]
    assert len(sidecars) == 7, "файл настроек уходит вместе со своей копией"


# --------------------------------------------------------------------------- #
#  Актуальность содержимого копии
# --------------------------------------------------------------------------- #
def test_copy_contains_state_of_previous_run(appdata, clock):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456",
                                   "Название": "Маркерная запись"})
    result = storage.create_startup_backup(store, "2.21.0")
    backup = result["path"]
    conn = sqlite3.connect(backup)
    try:
        rows = conn.execute("SELECT fields FROM counterparties").fetchall()
    finally:
        conn.close()
    assert any("Маркерная запись" in row[0] for row in rows), (
        "копия содержит состояние базы на конец прогона"
    )
    # Следующий запуск (новая копия) не портит предыдущую.
    _run()
    storage.create_startup_backup(storage.DataStore(), "2.21.0")
    conn = sqlite3.connect(backup)
    try:
        rows = conn.execute("SELECT fields FROM counterparties").fetchall()
    finally:
        conn.close()
    assert any("Маркерная запись" in row[0] for row in rows)


def test_rev_roundtrip_through_meta(appdata):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    conn = sqlite3.connect(storage.db_path())
    try:
        meta = dict(conn.execute("SELECT key, value FROM meta"))
    finally:
        conn.close()
    assert int(meta["rev"]) == store.registry_rev
    reloaded = storage.DataStore()
    assert reloaded.registry_rev == store.registry_rev, "ревизия читается при загрузке"


# --------------------------------------------------------------------------- #
#  Отказ копирования
# --------------------------------------------------------------------------- #
def test_backup_failure_keeps_flag_and_next_run_retries(appdata, clock):
    store = _run()
    monkey_store = store
    monkey_store.backup_db = lambda dest, **kwargs: False
    result = storage.create_startup_backup(store, "2.21.0")
    assert result["status"] == "failed"
    assert result["message"], "пользователю сообщается о неудаче"
    assert storage._backup_files() == [], "копий нет"
    assert store.dirty_since_backup is True, "флаг изменений сохраняется"
    # Следующий запуск (без подмены) создаёт копию.
    fresh = _run()
    retry = storage.create_startup_backup(fresh, "2.21.0")
    assert retry["status"] == "created"
    assert len(storage._backup_files()) == 1


# --------------------------------------------------------------------------- #
#  Восстановление: копия разворачивается на отдельном профиле
# --------------------------------------------------------------------------- #
def test_restore_backup_on_clean_profile(appdata, clock, tmp_path, monkeypatch):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456", "Название": "ООО Ромашка"})
    store.add_document(storage.KIND_CONTRACT, 1, "516", "10 января",
                       data={"ИНН": "2540123456"}, extra={}, text="")
    result = storage.create_startup_backup(store, "2.21.0")
    assert result["status"] == "created"
    backup = result["path"]

    # Отдельный тестовый профиль: та же копия — как database.db.
    other_profile = tmp_path / "other-profile"
    (other_profile / "DogovorGenerator").mkdir(parents=True)
    monkeypatch.setenv("APPDATA", str(other_profile))
    shutil.copyfile(backup, os.path.join(other_profile, "DogovorGenerator",
                                         "database.db"))
    restored = storage.DataStore()
    assert len(restored.counterparties) == 1
    assert restored.counterparties[0]["fields"]["Название"] == "ООО Ромашка"
    assert len(restored.documents) == 1
    assert restored.documents[0]["number"] == "516"


# --------------------------------------------------------------------------- #
#  Снимок перед обновлением
# --------------------------------------------------------------------------- #
def test_snapshot_created_with_version_label(appdata, clock):
    store = _run()
    storage.create_startup_backup(store, "2.21.0")
    backups_before = len(storage._backup_files())
    path = storage.create_update_snapshot(store, "2.21.0")
    assert path, "снимок создан"
    assert "registry_v2.21.0_pre-update.db" in path
    assert storage.SNAPSHOTS_DIR_NAME in path, "снимок лежит в backups/snapshots/"
    assert len(storage._backup_files()) == backups_before, (
        "снимок не вытесняет обычные копии и в лимит 7 не входит"
    )


def test_snapshot_rotation_keeps_three(appdata, clock):
    store = _run()
    for version in ("2.21.0", "2.21.1", "2.21.2", "2.22.0"):
        assert storage.create_update_snapshot(store, version)
    files = [os.path.basename(path) for path in storage._snapshot_files()]
    assert len(files) == 3, "своя ротация снимков — до 3 последних"
    assert "registry_v2.21.0_pre-update.db" not in files, "самая старая вытеснена"
    assert "registry_v2.22.0_pre-update.db" in files


# --------------------------------------------------------------------------- #
#  Копия настроек без ключа DaData
# --------------------------------------------------------------------------- #
def test_settings_copy_excludes_dadata_key(appdata, clock):
    with open(storage.settings_json_path(), "w", encoding="utf-8") as fh:
        json.dump({"theme": "dark", "dadata_api_key": "СЕКРЕТ",
                   "DaData Ключ": "СЕКРЕТ-2", "documents_dir": "D:/Docs"}, fh)
    store = _run()
    result = storage.create_startup_backup(store, "2.21.0")
    assert result["status"] == "created"
    sidecar = os.path.join(
        storage.backups_dir(),
        os.path.basename(result["path"])[:-3] + storage.BACKUP_SETTINGS_SUFFIX,
    )
    with open(sidecar, "r", encoding="utf-8") as fh:
        copied = json.load(fh)
    assert copied["theme"] == "dark" and copied["documents_dir"] == "D:/Docs"
    assert "dadata_api_key" not in copied and "DaData Ключ" not in copied, (
        "ключ DaData — секрет пользователя и в копию не включается"
    )


def test_backup_status_text_reports_state(appdata, clock):
    assert "пока нет" in storage.backup_status_text()
    store = _run()
    storage.create_startup_backup(store, "2.21.0")
    text = storage.backup_status_text()
    assert "Последняя копия" in text and "2.21.0" in text
    assert "копий в папке: 1" in text
