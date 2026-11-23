"""Юнит-тесты LAZY-1 (R-09): ленивые тела документов и единый getter.

Контракты карточки (этап 5 roadmap rev2):

* загрузка читает лёгкие колонки — data/extra/text в записях отсутствуют;
* ``document_body`` — единая точка доступа к телу: точечный SELECT из
  базы с собственным соединением, кэш в запись, ``cache=False`` для
  рабочего потока (START-1);
* сохранение (инкрементальное и полное) НЕ теряет тело ленивой записи —
  trash/restore без правки тела не затирают колонки пустыми значениями;
* сроки доп. соглашений живут в лёгкой карте ``expiry_until`` и
  поддерживаются мутаторами — рендер реестра не читает базу;
* критерий карточки: «экспорт после ленивой загрузки возвращает полный
  текст» (services.build_record_docx по ленивой записи).
"""
from __future__ import annotations

import sqlite3

import pytest

import services
import storage
from document_builder import F


def _make_party_doc(store, kind=storage.KIND_ADDITIONAL, text="Полный текст документа",
                    extra=None, data=None):
    party = store.add_counterparty(
        "ooo", {F.INN: "7701234567", F.SHORT_NAME: "ООО Ромашка"}, save=False,
    )
    doc = store.add_document(
        kind, party["id"], "1", "01.01.2026",
        data=data if data is not None else {F.INN: "7701234567"},
        extra=extra if extra is not None else {"valid_until": "31.12.2026"},
        text=text, save=True,
    )
    return party, doc


# --------------------------------------------------------------------------- #
#  Ленивая загрузка
# --------------------------------------------------------------------------- #
def test_load_keeps_body_lazy(store):
    _make_party_doc(store)
    reloaded = storage.DataStore()
    rec = reloaded.documents[0]
    assert "text" not in rec and "data" not in rec and "extra" not in rec


def test_document_body_loads_and_caches(store):
    _make_party_doc(store, text="ТЕЛО ДОКУМЕНТА")
    reloaded = storage.DataStore()
    body = reloaded.document_body(reloaded.documents[0]["id"])
    assert body["text"] == "ТЕЛО ДОКУМЕНТА"
    assert body["extra"]["valid_until"] == "31.12.2026"
    # Кэш в запись: база больше не нужна — после кэша тело берётся из записи.
    rec = reloaded.documents[0]
    assert "text" in rec
    assert reloaded.document_body(rec["id"])["text"] == "ТЕЛО ДОКУМЕНТА"


def test_document_body_cache_false_leaves_record_lazy(store):
    """Режим рабочего потока (START-1): тело возвращается, запись не меняется."""
    _make_party_doc(store)
    reloaded = storage.DataStore()
    body = reloaded.document_body(reloaded.documents[0]["id"], cache=False)
    assert body["text"] == "Полный текст документа"
    assert "text" not in reloaded.documents[0]


def test_document_body_unknown_and_missing(store):
    assert storage.DataStore().document_body(None) == {"text": "", "data": {}, "extra": {}}
    assert storage.DataStore().document_body(9999) == {"text": "", "data": {}, "extra": {}}


def test_corrupt_body_returns_empty_and_keeps_disk_row(store, appdata):
    """Битое тело: пустой результат, запись не кэшируется, диск не перезаписывается."""
    _party, doc = _make_party_doc(store)
    path = storage.db_path()
    conn = sqlite3.connect(path)
    with conn:
        conn.execute("UPDATE documents SET data = '{битый json' WHERE id = ?", (doc["id"],))
    conn.close()
    reloaded = storage.DataStore()
    rec = next(item for item in reloaded.documents if item["id"] == doc["id"])
    body = reloaded.document_body(rec["id"])
    assert body == {"text": "", "data": {}, "extra": {}}
    assert "data" not in rec  # битое тело не кэшируется и не помечается dirty
    assert not reloaded._dirty_documents


# --------------------------------------------------------------------------- #
#  Сохранение не теряет тело ленивой записи
# --------------------------------------------------------------------------- #
def test_incremental_save_after_lazy_preserves_body(store):
    _make_party_doc(store, text="НЕЗАМЕНЯЕМЫЙ ТЕКСТ")
    reloaded = storage.DataStore()
    reloaded.trash_document(reloaded.documents[0]["id"])  # dirty без тела
    assert reloaded.save()
    fresh = storage.DataStore()
    assert fresh.document_body(reloaded.documents[0]["id"])["text"] == "НЕЗАМЕНЯЕМЫЙ ТЕКСТ"


def test_restore_after_lazy_preserves_body(store):
    _make_party_doc(store, text="ТЕКСТ ПОСЛЕ ВОССТАНОВЛЕНИЯ")
    reloaded = storage.DataStore()
    doc_id = reloaded.documents[0]["id"]
    reloaded.trash_document(doc_id)
    reloaded.restore_document(doc_id)
    assert reloaded.save()
    fresh = storage.DataStore()
    assert fresh.document_body(doc_id)["text"] == "ТЕКСТ ПОСЛЕ ВОССТАНОВЛЕНИЯ"
    assert fresh.active_documents()


def test_save_full_after_lazy_preserves_body(store):
    _make_party_doc(store, text="ТЕКСТ ПОЛНОГО ПУТИ",
                    extra={"source_file": "полный.docx"})
    reloaded = storage.DataStore()
    assert reloaded.save_full()
    fresh = storage.DataStore()
    body = fresh.document_body(reloaded.documents[0]["id"])
    assert body["text"] == "ТЕКСТ ПОЛНОГО ПУТИ"
    assert body["extra"]["source_file"] == "полный.docx"


# --------------------------------------------------------------------------- #
#  Лёгкая карта сроков
# --------------------------------------------------------------------------- #
def test_expiry_cache_filled_at_load(store):
    _make_party_doc(store, extra={"valid_until": "31.12.2026"})
    reloaded = storage.DataStore()
    doc_id = reloaded.documents[0]["id"]
    assert reloaded.expiry_until(doc_id) == "31.12.2026"


def test_expiry_cache_ignores_spec_and_trash(store):
    party, _doc = _make_party_doc(store, kind=storage.KIND_SPEC, extra={"total": 1}, text="")
    store.add_document(storage.KIND_ADDITIONAL, party["id"], "2", "01.01.2026",
                       data={}, extra={"valid_until": "01.06.2026"}, text="", save=False)
    trashed = store.add_document(storage.KIND_ADDITIONAL, party["id"], "3", "01.01.2026",
                                 data={}, extra={"valid_until": "01.07.2026"}, text="", save=True)
    store.trash_document(trashed["id"])
    reloaded = storage.DataStore()
    # Спецификации без срока и изъятые записи в карту не попадают.
    values = {
        reloaded.expiry_until(rec["id"])
        for rec in reloaded.active_documents()
        if rec["kind"] == storage.KIND_ADDITIONAL
    }
    assert values == {"01.06.2026"}
    assert reloaded.expiry_until(trashed["id"]) is None


def test_expiry_cache_follows_mutators(store):
    _make_party_doc(store, extra={"valid_until": "01.03.2026"})
    rec = store.documents[0]
    assert store.expiry_until(rec["id"]) == "01.03.2026"  # мутатор обновил карту
    rec["extra"]["valid_until"] = "01.09.2026"
    store.mark_dirty(rec)
    assert store.expiry_until(rec["id"]) == "01.09.2026"


# --------------------------------------------------------------------------- #
#  bulk_source_files и экспорт (критерий карточки)
# --------------------------------------------------------------------------- #
def test_bulk_source_files_only_registered(store):
    _make_party_doc(store, extra={"source_file": "путь/файл.docx"})
    store.add_document(storage.KIND_CONTRACT, store.counterparties[0]["id"],
                       "2", "01.01.2026", data={}, extra={}, text="", save=False)
    result = storage.DataStore().bulk_source_files()
    assert list(result.values()) == ["путь/файл.docx"]


def test_export_after_lazy_load_returns_full_text(store):
    """Критерий карточки LAZY-1: экспорт по ленивой записи возвращает полный текст."""
    _make_party_doc(store, kind=storage.KIND_UNKNOWN, text="ПОЛНЫЙ ТЕКСТ ЭКСПОРТА")
    reloaded = storage.DataStore()
    rec = reloaded.documents[0]
    assert "text" not in rec
    doc = services.build_record_docx(rec, reloaded)
    content = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    assert "ПОЛНЫЙ ТЕКСТ ЭКСПОРТА" in content


def test_export_with_body_in_memory_unchanged(store):
    _make_party_doc(store, kind=storage.KIND_UNKNOWN, text="ТЕКСТ ЧЕРНОВИКА")
    rec = store.documents[0]
    doc = services.build_record_docx(rec, store)
    content = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    assert "ТЕКСТ ЧЕРНОВИКА" in content


# --------------------------------------------------------------------------- #
#  Сроки в agreements по лёгкой карте (без чтения базы)
# --------------------------------------------------------------------------- #
def test_agreements_collect_uses_expiry_cache(store):
    import agreements
    _make_party_doc(store, extra={"valid_until": "01.01.2020"})  # точно просрочено
    reloaded = storage.DataStore()
    rec = reloaded.documents[0]
    assert "extra" not in rec
    buckets = agreements.collect(reloaded)
    assert len(buckets["expired"]) == 1 and not buckets["soon"]
    text = agreements.notice_text(reloaded, docs=list(reloaded.active_documents()))
    assert text and "истекло" in text
