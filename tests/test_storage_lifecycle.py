"""Жизненный цикл реестра: корзина, запись, поиск, вспомогательные функции.

Добивает покрытие storage.py до цели плана (≥ 70 %) и фиксирует поведение
корзины (п. 2.5–2.6) и записи в SQLite (п. 11 этапа 5) на уровне pytest.
"""
from __future__ import annotations

import os

import storage
from document_builder import F
from render_helpers import OOO_FIELDS, IP_FIELDS


def _card(fields):
    return {k: v for k, v in fields.items()
            if k not in (F.NUMBER, F.DAY, F.MONTH)}


# ---------- Корзина (п. 2.5–2.6) ----------

def test_trash_and_restore_counterparty(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    contract = store.add_document(
        storage.KIND_CONTRACT, party["id"], "516", "10 января",
        data=dict(OOO_FIELDS), extra={}, text="",
    )
    assert store.trash_counterparty(party["id"]) == 1
    assert store.find_by_inn(OOO_FIELDS[F.INN]) is None  # скрыт из активных
    assert store.trashed_counterparties() and store.trashed_documents()

    assert store.restore_counterparty(party["id"]) == 1
    assert store.find_by_inn(OOO_FIELDS[F.INN]) is not None
    assert not store.trashed_documents()


def test_trash_document_independently(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    doc = store.add_document(
        storage.KIND_CONTRACT, party["id"], "516", "10 января",
        data=dict(OOO_FIELDS), extra={}, text="",
    )
    assert store.trash_document(doc["id"]) is True
    assert store.active_documents() == []
    assert store.documents_for(party["id"]) == []

    assert store.restore_document(doc["id"]) is True
    assert len(store.active_documents()) == 1


def test_purge_document_removes_permanently(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    doc = store.add_document(
        storage.KIND_CONTRACT, party["id"], "516", "10 января",
        data=dict(OOO_FIELDS), extra={}, text="",
    )
    store.trash_document(doc["id"])
    assert store.purge_document(doc["id"]) is True
    assert store.get_document(doc["id"], include_deleted=True) is None


def test_purge_all_trash(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    doc_a = store.add_document(storage.KIND_CONTRACT, party["id"], "1", "1",
                               data={}, extra={}, text="")
    doc_b = store.add_document(storage.KIND_CONTRACT, party["id"], "2", "1",
                               data={}, extra={}, text="")
    store.trash_document(doc_a["id"])
    store.trash_document(doc_b["id"])
    assert store.purge_all_trash() == 2
    assert store.trash_count() == 0


def test_cleanup_trash_respects_retention(store):
    from datetime import datetime, timedelta
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    old = store.add_document(storage.KIND_CONTRACT, party["id"], "1", "1",
                             data={}, extra={}, text="")
    fresh = store.add_document(storage.KIND_CONTRACT, party["id"], "2", "1",
                               data={}, extra={}, text="")
    store.trash_document(old["id"])
    store.trash_document(fresh["id"])
    stale_date = (datetime.now() - timedelta(days=storage.TRASH_RETENTION_DAYS + 1)
                  ).isoformat(timespec="microseconds")
    old["deleted_at"] = stale_date
    removed = store.cleanup_trash()
    assert removed == 1
    assert store.get_document(old["id"], include_deleted=True) is None
    assert store.get_document(fresh["id"], include_deleted=True) is not None


# ---------- Запись в базу (п. 11 этапа 5) ----------

def test_save_creates_sqlite_without_legacy_json(appdata):
    store = storage.DataStore()
    for number in range(1, 4):
        store.add_counterparty("ooo", {"ИНН": f"25401234567{number}"})
    assert os.path.exists(storage.db_path()), "реестр записан в базу SQLite"
    assert not os.path.exists(storage.legacy_json_path()), \
        "JSON больше не пишется: источник правды — SQLite"


def test_save_returns_true_and_no_error(appdata):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    assert store.save() is True
    assert store.last_error == ""


# ---------- Поиск ----------

def test_find_by_inn_and_name(store):
    store.add_counterparty("ooo", _card(OOO_FIELDS))
    store.add_counterparty("ip", _card(IP_FIELDS))

    by_inn = store.find_by_inn(OOO_FIELDS[F.INN])
    assert by_inn and by_inn["type"] == "ooo"

    by_name = store.find_by_name("ООО «Ромашка»")
    assert by_name is not None
    assert by_name["id"] == by_inn["id"]


def test_find_by_inn_ignores_trashed(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    store.trash_counterparty(party["id"])
    assert store.find_by_inn(OOO_FIELDS[F.INN]) is None


# ---------- Вспомогательные функции ----------

def test_display_names():
    ooo_rec = {"type": "ooo",
               "fields": {F.SHORT_NAME: "ООО «Ромашка»", F.INN: "2540123456"}}
    ip_rec = {"type": "ip", "fields": {F.IP_NAME: "Иванов Иван Иванович",
                                       F.INN: "254012345678"}}
    assert storage.display_name(ooo_rec) == "ООО «Ромашка»"
    assert storage.display_name(ip_rec) == "ИП Иванов Иван Иванович"
    assert "2540123456" in storage.display_with_inn(ooo_rec)


def test_contracts_for(store):
    party = store.add_counterparty("ooo", _card(OOO_FIELDS))
    contract = store.add_document(storage.KIND_CONTRACT, party["id"], "516",
                                  "10 января", data={}, extra={}, text="")
    store.add_document(storage.KIND_SPEC, party["id"], "12", "15 января",
                       data={}, extra={}, text="")
    contracts = store.contracts_for(party["id"])
    assert [c["id"] for c in contracts] == [contract["id"]]


def test_next_id_survives_reload(appdata):
    store = storage.DataStore()
    party = store.add_counterparty("ooo", {"ИНН": "2540123456"})
    assert store.save() is True
    reloaded = storage.DataStore()
    another = reloaded.add_counterparty("ooo", {"ИНН": "2540999999"}, save=False)
    assert another["id"] > party["id"]
