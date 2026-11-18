"""Негативные сценарии чистой логики (п. 7.8): валидация и отказы.

Сохранение без контрагента/номера/позиций, сбой транзакции сохранения
с откатом, спецификация без договора-основания — всё без интерфейса.
"""
from __future__ import annotations

import os

import services
import storage
from document_builder import F
from render_helpers import OOO_FIELDS, spec_rec, spec_total


# ---------- Валидация договора ----------

def test_empty_contract_lists_required_fields():
    errors = services.validate_contract_fields({}, is_ip=False)
    for key in (F.NUMBER, F.DAY, F.MONTH, F.INN, F.LEGAL_ADDRESS, F.ACCOUNT):
        assert key in errors, f"поле «{key}» не помечено как обязательное"


def test_ip_contract_requires_ogrnip_and_ip_name():
    minimal = {
        F.NUMBER: "1", F.DAY: "1", F.MONTH: "января",
        F.INN: "254012345678", F.LEGAL_ADDRESS: "г. Владивосток",
        F.ACCOUNT: "40802810500000012345",
        F.CORR_ACCOUNT: "30101810100000000225",
        F.BANK: "Банк", F.BIK: "044525225",
    }
    errors = services.validate_contract_fields(minimal, is_ip=True)
    assert F.OGRNIP in errors
    assert F.IP_NAME in errors


def test_bad_inn_length_flagged():
    data = {k: v for k, v in OOO_FIELDS.items()}
    data[F.INN] = "12345"
    errors = services.validate_contract_fields(data, is_ip=False)
    assert F.INN in errors and "10" in errors[F.INN]


def test_bad_account_length_flagged():
    data = {k: v for k, v in OOO_FIELDS.items()}
    data[F.ACCOUNT] = "4070281"
    errors = services.validate_contract_fields(data, is_ip=False)
    assert F.ACCOUNT in errors and "20" in errors[F.ACCOUNT]


def test_validate_all_fields_returns_list():
    result = services.validate_all_fields({}, is_ip=False)
    assert isinstance(result, list) and result


# ---------- Валидация карточки ----------

def test_party_card_email_optional():
    fields = {k: v for k, v in OOO_FIELDS.items()
              if k not in (F.NUMBER, F.DAY, F.MONTH)}
    fields[F.EMAIL] = ""
    errors = services.validate_party_fields(fields, is_ip=False)
    assert F.EMAIL not in errors  # email необязателен (п. 4.5)


def test_party_card_bad_email_flagged():
    fields = {k: v for k, v in OOO_FIELDS.items()
              if k not in (F.NUMBER, F.DAY, F.MONTH)}
    fields[F.EMAIL] = "не-почта"
    errors = services.validate_party_fields(fields, is_ip=False)
    assert F.EMAIL in errors


# ---------- Сбой транзакции сохранения ----------

def test_save_document_rolls_back_everything(store, tmp_path, monkeypatch,
                                             ooo_fields):
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    draft = {
        "kind": storage.KIND_CONTRACT,
        "party_id": party["id"],
        "number": "RB-1",
        "date": "10 января",
        "data": dict(ooo_fields),
        "extra": {"day": "10", "month": "января", "is_ip": False},
        "text": "",
    }

    real_save = store.save

    def failing_save():
        store.last_error = "имитация: диск недоступен"
        return False

    monkeypatch.setattr(store, "save", failing_save)
    docs_dir = tmp_path / "docs"
    docs_dir.mkdir()
    result = services.save_document(store, draft, str(docs_dir))
    monkeypatch.setattr(store, "save", real_save)

    assert result["ok"] is False
    assert result["record"] is None
    assert "диск" in result["error"]
    # Файл удалён, «осиротевших» документов нет.
    assert os.listdir(str(docs_dir)) == []
    assert not [d for d in store.documents if d["number"] == "RB-1"]


# ---------- Спецификация без договора-основания ----------

def test_spec_without_contract_renders_without_placeholders(store, ooo_fields):
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    rec = spec_rec(party, None)
    rec["extra"]["contract_number"] = ""
    rec["extra"]["total"] = spec_total()

    doc = services.render_spec(rec, store)
    from document_builder import docx_to_text
    text = docx_to_text(doc)
    assert "{{" not in text
    assert rec["number"] in text  # номер спецификации подставлен
