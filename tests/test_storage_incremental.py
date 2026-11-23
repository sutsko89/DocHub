"""STORE-1A (этап 3) — инкрементальная запись базы (R-01) и счётчики R-05a.

Карточка STORE-1A roadmap rev2: dirty-наборы, точечные
INSERT OR REPLACE / DELETE по изменённым id, обновление seq в meta при
добавлении записи; COUNT(*) в статусной строке с покрывающими индексами.
Схема базы не меняется — строки той же формы, schema_version прежний.

Критерии завершения из карточки (закрываются здесь и в smoke):

  * UI-время точечной save() на крупном реестре — миллисекунды, а не
    секунды (≤ 15 мс @10k подтверждается замером на целевой машине);
  * «500 случайных операций, после каждой дамп базы сравнивается
    с памятью» — test_500_random_ops_match_disk;
  * «база, записанная инкрементально, читается без миграции» —
    схема и форма строк проверяются структурно (test_...schema...);
  * «ошибка записи доходит до last_error и UI» — dirty сохраняется при
    сбое, повторный save() дописывает;
  * полный save_full() остаётся путём отката и самопроверки.
"""
from __future__ import annotations

import random
import sqlite3
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import storage
from document_builder import F

PARTY_COLUMNS = ("id", "type", "fields", "created_at", "deleted_at", "inn", "extra_keys")
DOC_COLUMNS = ("id", "kind", "party_id", "number", "date", "data", "extra", "text",
               "created_at", "deleted_at", "deleted_with_party", "extra_keys")


def _dump_db():
    conn = sqlite3.connect(storage.db_path())
    try:
        parties = {
            row[0]: row for row in conn.execute(
                f"SELECT {', '.join(PARTY_COLUMNS)} FROM counterparties")
        }
        documents = {
            row[0]: row for row in conn.execute(
                f"SELECT {', '.join(DOC_COLUMNS)} FROM documents")
        }
        meta = dict(conn.execute("SELECT key, value FROM meta"))
        indexes = {
            name for (name,) in conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'index'")
        }
    finally:
        conn.close()
    return parties, documents, meta, indexes


def _assert_memory_matches_disk(store: storage.DataStore) -> None:
    """Дамп базы совпадает с памятью: строки, id, seq и rev (STORE-1A).

    meta может быть пуста только до первой фактической записи (save()
    без изменений не открывает базу) — в этом случае seq/rev в базе ещё
    не появились и не сверяются.
    """
    parties, documents, meta, _ = _dump_db()
    assert set(parties) == {_safe_id(rec) for rec in store.counterparties}, \
        "набор id контрагентов на диске не совпадает с памятью"
    assert set(documents) == {_safe_id(rec) for rec in store.documents}, \
        "набор id документов на диске не совпадает с памятью"
    for rec in store.counterparties:
        assert parties[_safe_id(rec)] == store._party_row(rec), \
            f"строка контрагента id={_safe_id(rec)} на диске не совпадает с памятью"
    for rec in store.documents:
        assert documents[_safe_id(rec)] == store._document_row(rec), \
            f"строка документа id={_safe_id(rec)} на диске не совпадает с памятью"
    if meta:
        assert int(meta["seq"]) == store._seq, "seq в meta не соответствует памяти"
        assert int(meta["rev"]) == store._rev, "rev в meta не соответствует памяти"
    assert not store._has_dirty(), "dirty-наборы не очищены после успешной записи"


def _safe_id(rec: dict) -> int:
    return int(rec["id"])


_real_connect = storage._connect_db


# --------------------------------------------------------------------------- #
#  R-01: dirty-наборы и точечная запись
# --------------------------------------------------------------------------- #
def test_save_without_changes_is_noop(store):
    rev_before = store.registry_rev
    calls = {"full": 0, "incremental": 0}
    store._write_all = lambda *a, **k: calls.__setitem__("full", calls["full"] + 1)
    store._write_incremental = lambda *a, **k: calls.__setitem__(
        "incremental", calls["incremental"] + 1)
    assert store.save() is True, "save() без изменений возвращает True"
    assert calls == {"full": 0, "incremental": 0}, \
        "save() без изменений не должен открывать базу"
    assert store.registry_rev == rev_before, "ревизия выросла без записи"


def test_add_marks_dirty_and_writes_single_row(store):
    party = store.add_counterparty("ooo", {"ИНН": "2540123456",
                                           F.SHORT_NAME: "ООО Инкр"}, save=False)
    assert store.save() is True
    _assert_memory_matches_disk(store)
    parties, _documents, meta, _ = _dump_db()
    assert len(parties) == 1
    assert int(meta["seq"]) == store._seq, "seq в meta не обновлён при добавлении"
    assert int(meta["rev"]) == store._rev


def test_edit_writes_only_dirty_row(store):
    for i in range(5):
        store.add_counterparty("ooo", {"ИНН": f"254012345{i}",
                                       F.SHORT_NAME: f"ООО {i}"}, save=False)
    assert store.save() is True
    rev_before = store.registry_rev

    def no_full_write(*_args, **_kwargs):
        raise AssertionError("полная перезапись не должна вызываться: R-01")

    store._write_all = no_full_write
    edited = store.update_counterparty(store.counterparties[2]["id"],
                                       {"ИНН": "2540123452",
                                        F.SHORT_NAME: "ООО переименован"})
    assert edited is not None
    assert store.save() is True
    assert store.registry_rev == rev_before + 1, "ревизия не выросла при записи"
    _assert_memory_matches_disk(store)


def test_phantom_removal_becomes_delete(store, monkeypatch):
    """Сценарий services.save_document (DEF-4): запись добавлена, save() упал,
    вызывающий код снял запись из памяти — повторный save() не воскресает её."""
    rec = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    store.save()
    store.documents.append({"id": 99, "kind": storage.KIND_CONTRACT,
                            "party_id": rec["id"], "number": "1", "date": "",
                            "data": {}, "extra": {}, "text": ""})
    store.mark_dirty(store.documents[0])

    def locked_connect(path, *args, **kwargs):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(storage, "_connect_db", locked_connect)
    monkeypatch.setattr(storage.time, "sleep", lambda _seconds: None)
    assert store.save() is False, "первая запись документа должна упасть"
    # Откат вызывающего кода: запись снята из памяти, dirty-набор сохранился.
    store.documents.remove(store.documents[0])
    monkeypatch.setattr(storage, "_connect_db", _real_connect)
    assert store.save() is True
    _assert_memory_matches_disk(store)
    _parties, documents, _meta, _ = _dump_db()
    assert 99 not in documents, "фантомная запись осталась на диске"


def test_delete_purge_mark_deleted_ids(store):
    party = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    doc = store.add_document(storage.KIND_CONTRACT, party["id"], "516", "01.02.2026",
                             {}, save=False)
    store.save()
    assert store.purge_document(doc["id"]) is True
    _assert_memory_matches_disk(store)
    _parties, documents, _meta, _ = _dump_db()
    assert doc["id"] not in documents


def test_500_random_ops_match_disk(store):
    """Критерий карточки: 500 случайных операций, после каждой дамп == память."""
    rng = random.Random(20261121)
    INNS = iter(rng.randrange(2500000000, 2599999999) for _ in range(100000))
    op_weights = [("add_party", 3), ("add_doc", 3), ("edit_party", 2),
                  ("trash_doc", 2), ("restore_doc", 2), ("trash_party", 1),
                  ("restore_party", 1), ("purge_doc", 1), ("purge_party", 1),
                  ("purge_all", 1)]
    population = [op for op, w in op_weights for _ in range(w)]
    for step in range(500):
        op = rng.choice(population)
        if op == "add_party":
            store.add_counterparty(rng.choice(("ooo", "ip")),
                                   {"ИНН": str(next(INNS)),
                                    F.SHORT_NAME: f"ООО Оп{step}"}, save=False)
        elif op == "add_doc":
            active = store.active_counterparties()
            if not active:
                continue
            party = rng.choice(active)
            store.add_document(rng.choice((storage.KIND_CONTRACT, storage.KIND_SPEC,
                                           storage.KIND_ADDITIONAL)),
                               party["id"], f"{step}", f"01.0{step % 9 + 1}.2026",
                               {"ИНН": str(party["fields"].get("ИНН", ""))},
                               extra={"note": f"op{step}"}, text=f"текст {step}",
                               save=False)
        elif op == "edit_party":
            active = store.active_counterparties()
            if not active:
                continue
            party = rng.choice(active)
            fields = dict(party.get("fields") or {})
            fields[F.SHORT_NAME] = f"ООО Прав{step}"
            assert store.update_counterparty(party["id"], fields) is not None
        elif op == "trash_doc":
            docs = store.active_documents()
            if docs:
                assert store.trash_document(rng.choice(docs)["id"]) is True
        elif op == "restore_doc":
            trashed = store.trashed_documents()
            if trashed:
                assert store.restore_document(rng.choice(trashed)["id"]) is True
        elif op == "trash_party":
            active = store.active_counterparties()
            if active:
                party = rng.choice(active)
                store.trash_counterparty(party["id"])
        elif op == "restore_party":
            trashed = store.trashed_counterparties()
            if trashed:
                store.restore_counterparty(rng.choice(trashed)["id"])
        elif op == "purge_doc":
            trashed = store.trashed_documents()
            if trashed:
                assert store.purge_document(rng.choice(trashed)["id"]) is True
        elif op == "purge_party":
            trashed = store.trashed_counterparties()
            if trashed:
                store.purge_counterparty(rng.choice(trashed)["id"])
        elif op == "purge_all":
            store.purge_all_trash()
        assert store.save() is True, f"save() не удался на шаге {step} ({op})"
        _assert_memory_matches_disk(store)


def test_failed_write_keeps_dirty_and_retries_later(store, monkeypatch):
    store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)

    def locked_connect(path, *args, **kwargs):
        raise sqlite3.OperationalError("database is locked (имитация антивируса)")

    monkeypatch.setattr(storage, "_connect_db", locked_connect)
    monkeypatch.setattr(storage.time, "sleep", lambda _seconds: None)
    assert store.save() is False
    assert store.last_error, "причина сбоя не дошла до last_error"
    assert store._has_dirty(), "dirty-наборы очищены при несостоявшейся записи"

    monkeypatch.setattr(storage, "_connect_db", _real_connect)
    assert store.save() is True, "повторная запись после сбоя не дописала изменения"
    _assert_memory_matches_disk(store)


def test_unknown_keys_roundtrip_incremental(store):
    """Ключ без колонки (extra_keys) доезжает и через точечную запись."""
    rec = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    store.save()
    rec["custom_flag"] = "тест"
    store.mark_dirty(rec)
    assert store.save() is True
    reloaded = storage.DataStore()
    assert reloaded.counterparties[0].get("custom_flag") == "тест"


def test_legacy_field_migration_reaches_disk_incrementally(tmp_path, monkeypatch):
    """Нормализация имён полей при загрузке помечает запись dirty (STORE-1A)."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setattr(storage, "db_path", lambda: tmp_path / "database.db")
    storage.DataStore()
    conn = storage._connect_db(storage.db_path())
    try:
        conn.execute(
            "INSERT INTO counterparties (id, type, fields, created_at, "
            "deleted_at, extra_keys) VALUES (?, ?, ?, ?, ?, ?)",
            (1, "ooo",
             '{"Полное название компании (Покупь)": "ООО Ромашка", '
             '"ИНН": "7701234567"}',
             "2026-01-01T00:00:00", None, "{}"),
        )
        conn.commit()
    finally:
        conn.close()

    store = storage.DataStore()
    assert store._has_dirty(), "миграция при загрузке не пометила запись dirty"
    assert store.save() is True
    raw_conn = storage._connect_db(storage.db_path())
    try:
        raw = raw_conn.execute("SELECT fields FROM counterparties WHERE id = 1").fetchone()[0]
    finally:
        raw_conn.close()
    assert storage.LEGACY_FULL_NAME_FIELD not in raw


# --------------------------------------------------------------------------- #
#  Обратная совместимость схемы («читается кодом 2.20.0 без миграции»)
# --------------------------------------------------------------------------- #
def test_incrementally_written_db_keeps_schema_form(store):
    store.add_counterparty("ooo", {"ИНН": "2540123456", F.SHORT_NAME: "ООО Сх"},
                           save=False)
    doc = store.add_document(storage.KIND_CONTRACT, store.counterparties[0]["id"],
                             "516", "01.02.2026", {"ИНН": "2540123456"},
                             text="текст", save=False)
    store.save()

    _parties, _documents, meta, indexes = _dump_db()
    assert int(meta["schema_version"]) == storage.SQLITE_SCHEMA_VERSION, \
        "schema_version изменился — миграция не требовалась (карточка STORE-1A)"
    assert storage.SQLITE_SCHEMA_VERSION == 1
    assert "idx_documents_status" in indexes, "нет покрывающего индекса R-05a"

    # Форма строк не меняется: те же таблицы и колонки, что в _SCHEMA_SQL.
    conn = sqlite3.connect(storage.db_path())
    try:
        party_columns = [row[1] for row in conn.execute("PRAGMA table_info(counterparties)")]
        document_columns = [row[1] for row in conn.execute("PRAGMA table_info(documents)")]
    finally:
        conn.close()
    assert party_columns == list(PARTY_COLUMNS)
    assert document_columns == list(DOC_COLUMNS)

    # Свежая загрузка читает те же данные без миграций.
    # LAZY-1: тело ленивое — полный текст возвращается через document_body.
    reloaded = storage.DataStore()
    assert len(reloaded.counterparties) == 1 and len(reloaded.documents) == 1
    assert reloaded.documents[0]["id"] == doc["id"]
    assert reloaded.document_body(doc["id"])["text"] == "текст"


# --------------------------------------------------------------------------- #
#  R-05a: счётчики статусной строки COUNT(*) с покрывающими индексами
# --------------------------------------------------------------------------- #
def test_status_counts_matches_memory(store):
    active = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    store.add_counterparty("ip", {"ИНН": "2540123457"}, save=False)
    store.add_document(storage.KIND_CONTRACT, active["id"], "1", "01.01.2026", {}, save=False)
    store.add_document(storage.KIND_SPEC, active["id"], "2", "01.01.2026", {}, save=False)
    unknown = store.add_document(storage.KIND_UNKNOWN, active["id"], "3", "", {}, save=False)
    store.save()
    store.trash_document(unknown["id"])

    counts = store.status_counts()
    assert counts == store._status_counts_from_memory(), \
        "COUNT(*) расходится с подсчётом по памяти"
    assert counts["parties"] == 2
    assert counts["documents"] == 2, "записи неизвестного типа не исключены"
    assert counts["contracts"] == 1
    assert counts["trash"] == 1


def test_status_counts_falls_back_to_memory_when_db_locked(store, monkeypatch):
    store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)

    def denied(_path, *args, **kwargs):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(storage.sqlite3, "connect", denied)
    counts = store.status_counts()
    assert counts == {"parties": 1, "documents": 0, "contracts": 0, "trash": 0}, \
        "статусная строка не должна оставаться пустой при недоступной базе"


def test_covering_index_serves_status_queries(store):
    """План запроса — index-only скан по покрывающему индексу (R-05a)."""
    store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    store.save()
    conn = sqlite3.connect(storage.db_path())
    try:
        plan = conn.execute(
            "EXPLAIN QUERY PLAN SELECT COUNT(*) FROM documents "
            "WHERE deleted_at IS NULL").fetchall()
    finally:
        conn.close()
    detail = " ".join(str(row[-1]) for row in plan).lower()
    assert "idx_documents_status" in detail or "covering index" in detail, \
        f"COUNT(*) не использует покрывающий индекс: {plan}"


# --------------------------------------------------------------------------- #
#  Диагностика: точечная запись на крупном реестре
# --------------------------------------------------------------------------- #
def test_incremental_save_visible_in_perf_log(store, monkeypatch):
    """Критерий карточки: операции реестра видны в perf-канале (EWDOCHUB_PERF_LOG).

    DIAG-1 дублирует каждую applog.timed-точку в perf.log без порога:
    инкрементальная запись идёт через timed("Запись реестра (инкрементально)"),
    поэтому длительность операции измерима рядом с dt-тиками пулера —
    замер по разделу Е покажет, что операции реестра не прерывают тики.
    """
    import os as _os

    import perflog

    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    try:
        party = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
        assert store.save() is True
        perflog.flush()
        text = open(perflog.perf_path(), encoding="utf-8").read()
        assert "duration Запись реестра (инкрементально)" in text, \
            "точечная запись не видна в perf-канале"

        # Вторая фаза: без флага perf-лог не пишется вовсе.
        perflog.reset_for_tests()
        monkeypatch.delenv(perflog.ENV_FLAG)
        perflog.reset_for_tests()
        _os.remove(perflog.perf_path())
        store.update_counterparty(party["id"], {"ИНН": "2540123456"})
        perflog.flush()
        assert not _os.path.exists(perflog.perf_path()), \
            "perf-лог не должен писаться без флага"
    finally:
        perflog.reset_for_tests()


def test_single_edit_save_is_fast_on_2k_documents(store):
    """Точечная save() на реестре с 2000 документов — миллисекунды.

    Целевой порог карточки — ≤ 15 мс UI-времени при 10k на Windows
    (подтверждается замером по разделу Е); здесь — грубая граница для
    регрессии: одна правка не может стоить как полная перезапись.
    """
    party = store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    for i in range(2000):
        store.add_document(
            storage.KIND_SPEC, party["id"], f"{i}", "01.01.2026", {},
            extra={"items": [{"name": f"Позиция {i}", "qty": 1, "unit": "шт.",
                              "price": 10.0}], "total": 10.0},
            text="текст спецификации " * 20, save=False,
        )
    assert store.save() is True  # первичная (крупная) запись — INSERT OR REPLACE всех
    started = time.perf_counter()
    assert store.update_counterparty(party["id"],
                                     {"ИНН": "2540123456",
                                      F.SHORT_NAME: "ООО Переименовано"}) is not None
    elapsed_ms = (time.perf_counter() - started) * 1000
    assert elapsed_ms < 100, (
        f"точечная save() на 2000 документах заняла {elapsed_ms:.1f} мс — "
        "похоже, пишется весь реестр"
    )
    _assert_memory_matches_disk(store)
