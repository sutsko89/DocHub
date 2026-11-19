"""Проверка диска, а не памяти (п. 7.2): сервисный слой пишет реальные файлы.

После ``services.save_document``:
  * файл .docx существует на диске и открывается как документ;
  * текст печатной формы содержит подстановки;
  * новый ``DataStore()`` (прочитанный с диска) видит запись с теми же данными;
  * ``save()`` возвращает True, а не ложный успех.
"""
from __future__ import annotations

import os

import pytest
from docx import Document as open_docx

import services
import storage
from document_builder import F, docx_to_text, format_amount
from render_helpers import SPEC_ITEMS, spec_total, spec_rec


def _contract_draft(party: dict, fields: dict) -> dict:
    date_text = f"{fields[F.DAY]} {fields[F.MONTH]}"
    return {
        "kind": storage.KIND_CONTRACT,
        "party_id": party["id"],
        "number": fields[F.NUMBER],
        "date": date_text,
        "data": dict(fields),
        "extra": {"day": fields[F.DAY], "month": fields[F.MONTH], "is_ip": False},
        "text": "",
    }


def test_save_contract_disk_roundtrip(store, tmp_path, ooo_fields):
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    draft = _contract_draft(party, ooo_fields)
    result = services.save_document(store, draft, str(tmp_path))

    assert result["ok"] is True and result["error"] is None
    path = result["path"]

    # 1) файл на диске
    assert os.path.exists(path)
    # 2) открывается как .docx и содержит подстановки
    text = docx_to_text(open_docx(path))
    assert "№ 516" in text and ooo_fields[F.INN] in text
    assert "{{" not in text
    # 3) запись есть в реестре (в памяти)
    assert result["record"] in store.documents
    # 4) новый DataStore с диска видит ту же запись
    reloaded = storage.DataStore()
    loaded = next(d for d in reloaded.documents if d["number"] == "516")
    assert loaded["data"][F.INN] == ooo_fields[F.INN]
    assert loaded["extra"]["is_ip"] is False
    # 5) save() возвращает True
    assert store.save() is True


def test_save_spec_disk_roundtrip(store, tmp_path, ooo_fields):
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    contract = store.add_document(
        storage.KIND_CONTRACT, party["id"], "516", "10 января",
        data=dict(ooo_fields), extra={"day": "10", "month": "января",
                                      "is_ip": False}, text="",
    )
    rec = spec_rec(party, contract)
    draft = {
        "kind": storage.KIND_SPEC,
        "party_id": party["id"],
        "number": rec["number"],
        "date": "15 января",
        "data": rec["data"],
        "extra": rec["extra"],
        "text": "предпросмотр спецификации",
    }
    result = services.save_document(store, draft, str(tmp_path))

    assert result["ok"] is True
    path = result["path"]
    assert os.path.exists(path)
    text = docx_to_text(open_docx(path))
    assert format_amount(spec_total()) in text
    assert SPEC_ITEMS[0]["name"] in text
    assert "{{" not in text

    reloaded = storage.DataStore()
    loaded = next(d for d in reloaded.documents if d["kind"] == storage.KIND_SPEC)
    assert loaded["extra"]["contract_number"] == "516"
    assert loaded["extra"]["total"] == pytest.approx(spec_total())


def test_same_document_name_gets_numbered_suffix(store, tmp_path, ooo_fields):
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    first = services.save_document(store, _contract_draft(party, ooo_fields),
                                   str(tmp_path))
    second = services.save_document(store, _contract_draft(party, ooo_fields),
                                    str(tmp_path))
    assert first["ok"] and second["ok"]
    assert first["path"] != second["path"]
    assert "(2)" in os.path.basename(second["path"])


def test_spec_printed_total_matches_printed_lines(store, ooo_fields):
    """П. 9.8: «Итого» печатной формы равен сумме напечатанных строк.

    Копеечный раскол (3 × 3,335 + 1 × 1,115): строки дают 11,13, сырая
    сумма — 11,12. Старый путь писал в реестр сырую сумму и печатал её
    в «Итого» — расхождение с таблицей на копейку. Единый spec_total
    исключает его и исправляет итог старых записей.
    """
    from document_builder import parse_decimal, spec_total as db_spec_total

    party = store.add_counterparty("ooo", ooo_fields)
    items = [
        {"name": "Кабель", "qty": 3.0, "unit": "м", "price": 3.335},
        {"name": "Разъём", "qty": 1.0, "unit": "шт", "price": 1.115},
    ]
    rec = spec_rec(party, None)
    rec["extra"]["items"] = items
    rec["extra"]["total"] = 11.12  # сырая сумма, как её писала версия до 2.14.0
    rec["number"] = "13"

    doc = services.render_spec(rec, store)
    text = docx_to_text(doc)

    expected = format_amount(db_spec_total(items))
    raw = format_amount(parse_decimal("11.120"))
    assert expected == "11,13" and raw == "11,12"
    assert expected in text, "«Итого» — сумма копеечных строк"
    assert raw not in text, "сырая сумма больше не печатается как итог"
    assert "{{" not in text
