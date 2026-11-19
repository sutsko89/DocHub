"""Миграция имени поля «Полное название компании» (v2.13.0, п. 9.6).

Опечатка «(Покупь)» в ключе поля исправлена на «(Покупатель)». Каноническое
имя живёт в ``document_builder.F.FULL_NAME``; записи реестра со старым
ключом нормализуются при загрузке базы (и при переносе старого JSON, и при
добавлении/обновлении карточки через старых callers). Печатные формы и
предпросмотр читают значения по ключу F.FULL_NAME — без миграции они теряли
бы название контрагента из старых записей.
"""
from __future__ import annotations

import document_builder
import storage
from document_builder import F

OLD_KEY = storage.LEGACY_FULL_NAME_FIELD  # «Полное название компании (Покупь)»


def test_full_name_constant_renamed():
    assert F.FULL_NAME == "Полное название компании (Покупатель)"
    assert "(Покупь)" not in F.FULL_NAME


def test_normalize_renames_party_fields():
    rec = {"fields": {OLD_KEY: "Общество с ограниченной ответственностью Ромашка"}}
    storage._normalize_field_names(rec)
    assert rec["fields"].get(F.FULL_NAME) == \
        "Общество с ограниченной ответственностью Ромашка"
    assert OLD_KEY not in rec["fields"]


def test_normalize_renames_document_data():
    """Данные документа — копия полей карточки: старый ключ там тоже migrируется."""
    rec = {"data": {OLD_KEY: "ООО Ромашка", F.INN: "7701234567"}}
    storage._normalize_field_names(rec)
    assert rec["data"].get(F.FULL_NAME) == "ООО Ромашка"
    assert OLD_KEY not in rec["data"]
    assert rec["data"][F.INN] == "7701234567"  # соседние ключи не тронуты


def test_normalize_keeps_new_value_when_both_keys_present():
    rec = {"fields": {OLD_KEY: "старое значение", F.FULL_NAME: "новое значение"}}
    storage._normalize_field_names(rec)
    assert rec["fields"] == {F.FULL_NAME: "новое значение"}


def test_normalize_is_idempotent():
    rec = {"fields": {OLD_KEY: "ООО Ромашка"}}
    storage._normalize_field_names(rec)
    storage._normalize_field_names(rec)
    assert rec["fields"] == {F.FULL_NAME: "ООО Ромашка"}


def test_normalize_leaves_non_dict_payloads_alone():
    rec = {"fields": None, "data": "не словарь"}
    storage._normalize_field_names(rec)  # не бросает исключений
    assert rec == {"fields": None, "data": "не словарь"}


def test_db_migration_renames_key_on_load(tmp_path, monkeypatch):
    """Запись, сохранённая со старым ключом, при повторном открытии — с новым."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    writer = storage.DataStore()
    rec = writer.add_counterparty("ooo", {OLD_KEY: "ООО Ромашка"}, save=True)
    assert writer.counterparties[0]["fields"].get(F.FULL_NAME) == "ООО Ромашка"

    reader = storage.DataStore()
    loaded = reader.get_counterparty(rec["id"])
    assert loaded is not None
    assert loaded["fields"].get(F.FULL_NAME) == "ООО Ромашка"
    assert OLD_KEY not in loaded["fields"]


def test_db_migration_persists_renamed_key(tmp_path, monkeypatch):
    """После миграции и сохранения старый ключ исчезает и из самой базы."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    store = storage.DataStore()
    store.add_counterparty("ooo", {OLD_KEY: "ООО Ромашка"}, save=True)
    store.save()  # повторная запись уже с нормализованными ключами

    conn = storage._connect_db(storage.db_path())
    try:
        raw = conn.execute(
            "SELECT fields FROM counterparties LIMIT 1").fetchone()[0]
    finally:
        conn.close()
    assert OLD_KEY not in raw
    assert F.FULL_NAME in raw


def test_db_migration_reads_row_written_by_old_version(tmp_path, monkeypatch):
    """Сырая строка со старым ключом (так пишет версия до 2.13.0) — при
    загрузке ключ переименован, и после сохранения база не содержит старого."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    storage.DataStore()  # создаёт схему

    conn = storage._connect_db(storage.db_path())
    try:
        conn.execute(
            "INSERT INTO counterparties (id, type, fields, created_at, "
            "deleted_at, extra_keys) VALUES (?, ?, ?, ?, ?, ?)",
            (
                1, "ooo",
                '{"Полное название компании (Покупь)": "ООО Ромашка", '
                '"Сокращенное название компании": "ООО Ромашка", '
                '"ИНН": "7701234567"}',
                "2026-01-01T00:00:00", None, "{}",
            ),
        )
        conn.commit()
    finally:
        conn.close()

    store = storage.DataStore()
    loaded = store.get_counterparty(1)
    assert loaded is not None
    assert loaded["fields"].get(F.FULL_NAME) == "ООО Ромашка"
    assert OLD_KEY not in loaded["fields"]

    assert store.save()
    conn = storage._connect_db(storage.db_path())
    try:
        raw = conn.execute(
            "SELECT fields FROM counterparties WHERE id = 1").fetchone()[0]
    finally:
        conn.close()
    assert OLD_KEY not in raw and F.FULL_NAME in raw


def test_find_by_name_uses_renamed_key(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    store = storage.DataStore()
    store.add_counterparty("ooo", {
        F.FULL_NAME: "Общество с ограниченной ответственностью Ромашка",
        F.SHORT_NAME: "ООО Ромашка",
    }, save=True)
    found = store.find_by_name("общество с ограниченной ответственностью ромашка")
    assert found is not None
    assert found["fields"][F.SHORT_NAME] == "ООО Ромашка"


def test_display_name_fallback_uses_renamed_key():
    rec = {"type": "ooo", "fields": {F.FULL_NAME: "ООО Ромашка"}}
    assert storage.display_name(rec) == "ООО Ромашка"
    assert storage.display_name({"type": "ooo", "fields": {}}) == "Без названия"


def test_add_counterparty_normalizes_legacy_input(tmp_path, monkeypatch):
    """Старые callers, передающие ключ «(Покупь)», получают канонический ключ."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    store = storage.DataStore()
    rec = store.add_counterparty("ooo", {OLD_KEY: "ООО Ромашка"}, save=False)
    assert rec["fields"].get(F.FULL_NAME) == "ООО Ромашка"
    assert OLD_KEY not in rec["fields"]


def test_update_counterparty_normalizes_legacy_input(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    store = storage.DataStore()
    rec = store.add_counterparty("ooo", {F.SHORT_NAME: "ООО"}, save=False)
    updated = store.update_counterparty(rec["id"], {OLD_KEY: "ООО Ромашка"})
    assert updated is not None
    assert updated["fields"].get(F.FULL_NAME) == "ООО Ромашка"
    assert OLD_KEY not in updated["fields"]
