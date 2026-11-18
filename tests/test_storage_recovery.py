"""Тесты восстановления реестра (п. 7.3).

Битый JSON (мусор/обрезка/null/дубли id), восстановление из ``.tmp``,
``.bad``-копия, read-only папка при записи, осиротевшие файлы и восстановление
реестра из папки документов (reconcile).
"""
from __future__ import annotations

import json
import os
import stat
import sys

import pytest

import storage
from document_builder import F
from render_helpers import OOO_FIELDS


def _write_db(text: str) -> str:
    path = storage.db_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)
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


# ---------- Битый JSON ----------

def test_garbage_json_starts_fresh_with_bad_copy(appdata):
    path = _write_db("мусор {{{ это не JSON")
    store = storage.DataStore()
    assert store.counterparties == [] and store.documents == []
    assert store.load_warning, "пользователь должен быть предупреждён"
    bad_copies = [f for f in os.listdir(os.path.dirname(path))
                  if f.startswith("database.json.bad")]
    assert bad_copies, ".bad-копия повреждённого файла не создана"


def test_truncated_json_recovered(appdata):
    text = json.dumps(_valid_registry(), ensure_ascii=False)
    _write_db(text[: len(text) // 2])
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning


def test_null_json_not_a_dict(appdata):
    _write_db("null")
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning


def test_sections_of_wrong_type_shown_empty(appdata):
    registry = _valid_registry()
    registry["counterparties"] = None
    _write_db(json.dumps(registry, ensure_ascii=False))
    store = storage.DataStore()
    assert store.counterparties == [] and store.load_warning


def test_duplicate_ids_do_not_crash_and_seq_is_raised(appdata):
    registry = _valid_registry()
    registry["counterparties"].append(dict(registry["counterparties"][0]))
    _write_db(json.dumps(registry, ensure_ascii=False))
    store = storage.DataStore()
    assert len(store.counterparties) == 2
    new_party = store.add_counterparty("ooo", {"ИНН": "7709999999"}, save=False)
    ids = {rec["id"] for rec in store.counterparties}
    assert new_party["id"] not in ids or new_party["id"] > 1


def test_records_with_wrong_type_dropped(appdata):
    registry = _valid_registry()
    registry["counterparties"].append("строка вместо записи")
    _write_db(json.dumps(registry, ensure_ascii=False))
    store = storage.DataStore()
    assert len(store.counterparties) == 1  # битая запись отброшена, валидная жива


# ---------- Восстановление из .tmp ----------

def test_registry_recovered_from_tmp_file(appdata):
    base = storage.db_path()
    os.makedirs(os.path.dirname(base), exist_ok=True)
    with open(base + ".1234.tmp", "w", encoding="utf-8") as fh:
        json.dump(_valid_registry(), fh, ensure_ascii=False)
    with open(base, "w", encoding="utf-8") as fh:
        fh.write("битые данные")
    store = storage.DataStore()
    assert len(store.counterparties) == 1
    assert store.load_warning  # «восстановлен из временного файла»


# ---------- Запись в read-only папку ----------

@pytest.mark.skipif(not hasattr(os, "geteuid") or os.geteuid() == 0,
                    reason="chmod-блокировка не работает под root / не POSIX")
def test_save_to_readonly_dir_returns_false(appdata):
    store = storage.DataStore()
    store.add_counterparty("ooo", {"ИНН": "2540123456"})
    db_dir = os.path.dirname(storage.db_path())
    os.chmod(db_dir, stat.S_IREAD | stat.S_IEXEC)
    try:
        assert store.save() is False
        assert store.last_error
    finally:
        os.chmod(db_dir, stat.S_IRWXU)


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
