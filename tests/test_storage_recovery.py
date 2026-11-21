"""Тесты переноса и восстановления реестра (п. 11 этапа 5).

Миграция JSON → SQLite (сохранность данных, корзины, неизвестных ключей
и счётчика id), повреждённая база (.bad-копия + предупреждение), битый
legacy-JSON (файл не трогаем), сбой записи и сверка с папкой (reconcile).
"""
from __future__ import annotations

import json
import os
import sqlite3

import storage
from document_builder import F
from render_helpers import OOO_FIELDS


def _write_json_registry(payload) -> str:
    path = storage.legacy_json_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
    return path


def _valid_registry() -> dict:
    return {
        "schema_version": storage.SCHEMA_VERSION,
        "seq": 2,
        "counterparties": [{
            "id": 1, "type": "ooo",
            "fields": {k: v for k, v in OOO_FIELDS.items()
                       if k not in (F.NUMBER, F.DAY, F.MONTH)},
        }],
        "documents": [],
    }


# ---------- Миграция JSON → SQLite ----------

def test_migration_from_json_preserves_data(appdata):
    registry = _valid_registry()
    registry["documents"].append({
        "id": 2, "kind": storage.KIND_CONTRACT, "party_id": 1,
        "number": "516", "date": "10 января", "data": {"a": "1"},
        "extra": {"restored": True}, "text": "текст",
        "created_at": "2026-01-01T00:00:00",
    })
    registry["counterparties"][0]["archived"] = True  # неизвестный ключ записи
    _write_json_registry(registry)

    store = storage.DataStore()
    assert len(store.counterparties) == 1 and len(store.documents) == 1
    assert store.counterparties[0]["archived"] is True
    assert store.documents[0]["extra"].get("restored") is True
    assert not store.load_warning

    assert os.path.exists(storage.db_path()), "создана база SQLite"
    assert not os.path.exists(storage.legacy_json_path()), "JSON переименован"
    assert os.path.exists(storage.legacy_json_path() + ".migrated")

    reloaded = storage.DataStore()
    assert len(reloaded.counterparties) == 1 and len(reloaded.documents) == 1
    assert reloaded.counterparties[0]["archived"] is True, \
        "неизвестные ключи записи переживают перенос"
    another = reloaded.add_counterparty("ooo", {"ИНН": "2540999999"}, save=False)
    assert another["id"] > 2, "счётчик id продолжается после переноса"


def test_migration_moves_trash_flags(appdata):
    registry = _valid_registry()
    registry["documents"].append({
        "id": 2, "kind": storage.KIND_CONTRACT, "party_id": 1,
        "number": "1", "date": "1", "data": {}, "extra": {}, "text": "",
        "created_at": "2026-01-01T00:00:00",
        "deleted_at": "2026-02-01T00:00:00.123456",
        "deleted_with_party": True,
    })
    _write_json_registry(registry)

    store = storage.DataStore()
    trashed = store.trashed_documents()
    assert len(trashed) == 1 and trashed[0]["deleted_with_party"] is True

    reloaded = storage.DataStore()
    assert len(reloaded.trashed_documents()) == 1, "корзина переживает перенос"
    assert reloaded.restore_document(2) is True
    assert reloaded.active_documents(), "восстановление после переноса работает"


def test_garbage_json_migration_keeps_file(appdata):
    path = storage.legacy_json_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("мусор {{{ это не JSON")

    store = storage.DataStore()
    assert store.counterparties == [] and store.documents == []
    assert store.load_warning, "пользователь должен быть предупреждён"
    assert os.path.exists(path), \
        "повреждённый JSON не трогаем — это единственная копия данных"
    assert not os.path.exists(storage.legacy_json_path() + ".migrated")


def test_truncated_json_migration_starts_empty(appdata):
    text = json.dumps(_valid_registry(), ensure_ascii=False)
    path = _write_json_registry(text[: len(text) // 2])
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning
    assert os.path.exists(path), "обрезанный JSON сохранён для ручного восстановления"


def test_null_json_not_a_dict(appdata):
    path = _write_json_registry("null")
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning
    assert os.path.exists(path)


def test_sections_of_wrong_type_shown_empty(appdata):
    registry = _valid_registry()
    registry["counterparties"] = None
    path = _write_json_registry(registry)
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning
    assert os.path.exists(path)


def test_duplicate_ids_deduplicated_with_seq_raised(appdata):
    registry = _valid_registry()
    registry["counterparties"].append(dict(registry["counterparties"][0]))
    _write_json_registry(registry)
    store = storage.DataStore()
    assert len(store.counterparties) == 1, "дубликат id пропущен при переносе"
    new_party = store.add_counterparty("ooo", {"ИНН": "7709999999"}, save=False)
    assert new_party["id"] > 1, "счётчик id не конфликтует с перенесёнными"


def test_records_with_wrong_type_dropped(appdata):
    registry = _valid_registry()
    registry["counterparties"].append("строка вместо записи")
    _write_json_registry(registry)
    store = storage.DataStore()
    assert len(store.counterparties) == 1  # битая запись отброшена, валидная жива


# ---------- Повреждённая база SQLite ----------

def test_corrupt_sqlite_file_quarantined(appdata):
    path = storage.db_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(b"not a sqlite database at all" * 10)

    store = storage.DataStore()
    assert store.counterparties == [] and store.documents == []
    assert store.load_warning, "пользователь должен быть предупреждён"
    bad_copies = [f for f in os.listdir(os.path.dirname(path))
                  if f.startswith("database.db.bad")]
    assert bad_copies, ".bad-копия повреждённой базы не создана"


# ---------- Схема и режимы базы ----------

def test_schema_indexes_and_wal(appdata):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    conn = sqlite3.connect(storage.db_path())
    try:
        mode = conn.execute("PRAGMA journal_mode").fetchone()[0]
        indexes = {name for (name,) in conn.execute(
            "SELECT name FROM sqlite_master WHERE type = 'index'")}
    finally:
        conn.close()
    assert mode == "wal", "WAL-журнал включён (параллельное чтение)"
    assert {"idx_counterparties_inn", "idx_documents_party",
            "idx_documents_kind"} <= indexes, "индексы по ИНН и документам созданы"


def test_meta_schema_version_written(appdata):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    conn = sqlite3.connect(storage.db_path())
    try:
        meta = dict(conn.execute("SELECT key, value FROM meta"))
    finally:
        conn.close()
    assert int(meta["schema_version"]) == storage.SQLITE_SCHEMA_VERSION
    assert int(meta["seq"]) >= 2, "счётчик id записан в базу"


def test_unknown_record_keys_roundtrip(appdata):
    store = storage.DataStore()
    party = store.add_counterparty("ooo", {"ИНН": "2540123456"})
    party["custom_flag"] = "тест"  # ключ, для которого нет колонки
    # STORE-1A: мутация мимо мутатора — полный путь save_full()
    # (самопроверка перезаписывает весь реестр, включая extra_keys).
    assert store.save_full() is True
    reloaded = storage.DataStore()
    assert reloaded.counterparties[0].get("custom_flag") == "тест"


# ---------- Запись при отказе ОС ----------

def test_save_returns_false_when_os_refuses_write(appdata, monkeypatch):
    """save() возвращает False, когда ОС отказала в записи (аудит 2.0.1).

    Прежний тест имитировал read-only папку через chmod — он не работает на
    Windows (права каталога там не блокируют запись) и под root, поэтому в
    CI приходилось пропускать тест. Контракт «save() честно сообщает об
    отказе и заполняет last_error» проверяется платформонезависимо:
    открытие базы отказывает в доступе.
    """
    store = storage.DataStore()
    # STORE-1A: запись с помеченным dirty-изменением (save=False) — иначе
    # инкрементальный save() без изменений верно вернул бы True no-op.
    store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)

    def denied_connect(path, *args, **kwargs):
        raise PermissionError(13, "Отказано в доступе")

    monkeypatch.setattr(storage, "_connect_db", denied_connect)
    monkeypatch.setattr(storage.time, "sleep", lambda _seconds: None)
    assert store.save() is False
    assert store.last_error  # причина доносится до пользователя
    # Реестр в памяти не «пропал» из-за сбоя записи.
    assert len(store.counterparties) == 1


# ---------- Осиротевшие файлы и сверка с папкой ----------

def test_orphan_docx_restored_into_registry(store, tmp_path):
    from reconcile import find_orphans, restore_orphans
    from services import safe_filename

    party = store.add_counterparty(
        "ooo", {k: v for k, v in OOO_FIELDS.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    docs_root = tmp_path / "docs"
    party_dir = docs_root / safe_filename(storage.display_name(party), "Контрагент")
    party_dir.mkdir(parents=True)

    # Настоящий .docx с правильным именем файла (тип и номер — из имени).
    from services import render_contract, save_docx_atomic
    orphan = party_dir / "Договор поставки №99 ООО Ромашка.docx"
    save_docx_atomic(render_contract(OOO_FIELDS, is_ip=False), str(orphan))

    orphans = find_orphans(str(docs_root), store)
    assert str(orphan) in orphans

    restored = restore_orphans(store, str(docs_root), orphans, safe_filename)
    assert restored == 1
    rec = next(d for d in store.documents if d["number"] == "99")
    assert rec["kind"] == storage.KIND_CONTRACT
    assert rec["party_id"] == party["id"]
    assert (rec.get("extra") or {}).get("restored") is True
    assert not find_orphans(str(docs_root), store)


# --------------------------------------------------------------------------- #
#  D1: восстановление осиротевших файлов поднимает весь реестр
# --------------------------------------------------------------------------- #
def test_restore_fifty_orphans_fully(store, tmp_path):
    from reconcile import find_orphans, restore_orphans
    from services import safe_filename

    party = store.add_counterparty(
        "ooo", {k: v for k, v in OOO_FIELDS.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    docs_root = tmp_path / "docs"
    party_dir = docs_root / safe_filename(storage.display_name(party), "Контрагент")
    party_dir.mkdir(parents=True)
    for number in range(1, 51):
        (party_dir / f"Договор поставки №{number} ООО Ромашка.docx").write_bytes(b"PK\x05\x06")

    orphans = find_orphans(str(docs_root), store)
    assert len(orphans) == 50
    restored = restore_orphans(store, str(docs_root), orphans, safe_filename)
    assert restored == 50, "восстановление не обрывается после первой записи"
    assert not find_orphans(str(docs_root), store)


def test_restore_not_broken_by_stale_last_error(store, tmp_path):
    """Устаревшая last_error прошлых операций не прерывает цикл (D1)."""
    from reconcile import restore_orphans
    from services import safe_filename

    party = store.add_counterparty(
        "ooo", {k: v for k, v in OOO_FIELDS.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    docs_root = tmp_path / "docs"
    party_dir = docs_root / safe_filename(storage.display_name(party), "Контрагент")
    party_dir.mkdir(parents=True)
    paths = []
    for number in (1, 2, 3):
        path = party_dir / f"Договор поставки №{number}.docx"
        path.write_bytes(b"PK\x05\x06")
        paths.append(str(path))

    store.last_error = "устаревшая ошибка прошлой операции"
    restored = restore_orphans(store, str(docs_root), paths, safe_filename)
    assert restored == 3, "проверка после add_document(save=False) больше не ложная"
    assert store.last_error == "", "устаревшая ошибка сброшена"


def test_restore_save_failure_gives_honest_message(store, tmp_path, monkeypatch):
    """Сбой финального save() — честное сообщение, счётчики корректны."""
    from reconcile import restore_orphans
    from services import safe_filename

    party = store.add_counterparty(
        "ooo", {k: v for k, v in OOO_FIELDS.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    docs_root = tmp_path / "docs"
    party_dir = docs_root / safe_filename(storage.display_name(party), "Контрагент")
    party_dir.mkdir(parents=True)
    paths = []
    for number in (1, 2):
        path = party_dir / f"Договор поставки №{number}.docx"
        path.write_bytes(b"PK\x05\x06")
        paths.append(str(path))

    def failing_save():
        store.last_error = "Не удалось записать реестр: диск недоступен"
        return False

    monkeypatch.setattr(store, "save", failing_save)
    restored = restore_orphans(store, str(docs_root), paths, safe_filename)
    assert restored == 2, "записи добавлены в память, счётчик корректен"
    assert "Не удалось записать реестр" in store.last_error, (
        "сбой записи доносится до вызывающего кода"
    )


def test_restore_empty_path_list_unchanged(store):
    from reconcile import restore_orphans

    assert restore_orphans(store, "/nonexistent", []) == 0
    assert store.documents == []
