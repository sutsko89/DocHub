"""Общие помощники golden-тестов: образцы данных и сигнатуры печатных форм.

Используются и pytest-набором (test_golden_forms.py), и генератором эталонов
(make_golden.py). GUI не создаётся — только сервисный слой и document_builder.
"""
from __future__ import annotations

import os
import re
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Изолированный профиль пользователя на случай импорта storage без фикстур.
os.environ.setdefault(
    "APPDATA", os.path.join(tempfile.gettempdir(), "dochub_helpers_appdata")
)

import storage  # noqa: E402
import services  # noqa: E402
from document_builder import F, docx_to_text, format_amount, spec_vat_amount_from_items  # noqa: E402


OOO_FIELDS = {
    F.NUMBER: "516", F.DAY: "10", F.MONTH: "января",
    F.FULL_NAME: "Общество с ограниченной ответственностью «Ромашка»",
    F.SHORT_NAME: "ООО «Ромашка»",
    F.DIRECTOR: "Петров Пётр Петрович",
    F.BASIS: "Устав",
    F.INN: "2540123456", F.KPP: "254001001", F.OGRN: "1035001234567",
    F.LEGAL_ADDRESS: "690000, г. Владивосток, ул. Светланская, д. 1",
    F.FACT_ADDRESS: "690000, г. Владивосток, ул. Светланская, д. 1",
    F.EMAIL: "buy@romashka.ru", F.PHONE: "+7 (423) 123-45-67",
    F.ACCOUNT: "40702810123456789012", F.CORR_ACCOUNT: "30101810100000000225",
    F.BANK: "ПАО Приморский банк", F.BIK: "044525225",
}

IP_FIELDS = {
    F.NUMBER: "77", F.DAY: "12", F.MONTH: "февраля",
    F.IP_NAME: "Иванов Иван Иванович",
    F.INN: "254012345678", F.OGRNIP: "304253600012345",
    F.LEGAL_ADDRESS: "690000, г. Владивосток, ул. Пушкинская, д. 2",
    F.FACT_ADDRESS: "",
    F.EMAIL: "ivanov@ip.ru", F.PHONE: "+7 (423) 555-66-77",
    F.ACCOUNT: "40802810500000012345", F.CORR_ACCOUNT: "30101810100000000225",
    F.BANK: "ПАО Приморский банк", F.BIK: "044525225",
}

SPEC_ITEMS = [
    {"name": "Кабель ВВГнг 3х2.5", "qty": 120.0, "unit": "м", "price": 85.5},
    {"name": "Розетка Schneider Electric", "qty": 40.0, "unit": "шт", "price": 210.0},
]


def spec_total() -> float:
    """Итоговая сумма образца спецификации (как её хранит реестр)."""
    from decimal import Decimal
    total = sum(
        (Decimal(str(i["qty"])) * Decimal(str(i["price"])) for i in SPEC_ITEMS),
        Decimal("0"),
    )
    return float(total)


def spec_rec(party_rec: dict, contract_rec: dict | None) -> dict:
    """Запись спецификации для печатной формы."""
    return {
        "kind": storage.KIND_SPEC,
        "party_id": party_rec.get("id"),
        "number": "12",
        "data": dict(party_rec.get("fields") or {}),
        "extra": {
            "contract_id": (contract_rec or {}).get("id"),
            "contract_number": (contract_rec or {}).get("number") or "516",
            "items": [dict(i) for i in SPEC_ITEMS],
            "total": spec_total(),
            "spec_day": "15", "spec_month": "января",
            "prices_until": "", "delivery_term": "10 дней",
            "payment_terms": "отсрочка 30 дней", "delivery_basis": "самовывоз",
            "is_ip": (party_rec or {}).get("type") == "ip",
        },
        "text": "",
    }


def normalize(text: str) -> str:
    """Убирает меняющиеся со временем значения (даты «действуют до», год)."""
    text = re.sub(r"\b\d{2}\.\d{2}\.\d{4}\b", "<ДАТА>", text)
    text = re.sub(r"\b(?:19|20)\d{2} г\.", "<ГОД> г.", text)
    return text


def document_signature(doc) -> dict:
    """Сигнатура печатной формы: текст (с таблицами) + таблицы отдельно."""
    tables = [
        [normalize(cell.text.strip()) for cell in row.cells]
        for table in doc.tables
        for row in table.rows
    ]
    return {"text": normalize(docx_to_text(doc)), "tables": tables}


def render_six_forms() -> dict:
    """Строит шесть печатных форм (ООО/ИП × договор/доп./спецификация).

    Возвращает {имя: сигнатура}. Использует изолированный временный профиль
    и настоящие шаблоны из корня репозитория.
    """
    os.environ["APPDATA"] = os.path.join(
        tempfile.gettempdir(), "dochub_golden_appdata"
    )
    store = storage.DataStore()

    party_ooo = store.add_counterparty("ooo", {
        k: v for k, v in OOO_FIELDS.items()
        if k not in (F.NUMBER, F.DAY, F.MONTH)
    }, save=False)
    party_ip = store.add_counterparty("ip", {
        k: v for k, v in IP_FIELDS.items()
        if k not in (F.NUMBER, F.DAY, F.MONTH)
    }, save=False)

    contract_ooo = store.add_document(
        storage.KIND_CONTRACT, party_ooo["id"], OOO_FIELDS[F.NUMBER],
        f"{OOO_FIELDS[F.DAY]} {OOO_FIELDS[F.MONTH]}",
        data=dict(OOO_FIELDS), extra={"day": OOO_FIELDS[F.DAY],
                                      "month": OOO_FIELDS[F.MONTH], "is_ip": False},
        text="", save=False,
    )
    contract_ip = store.add_document(
        storage.KIND_CONTRACT, party_ip["id"], IP_FIELDS[F.NUMBER],
        f"{IP_FIELDS[F.DAY]} {IP_FIELDS[F.MONTH]}",
        data=dict(IP_FIELDS), extra={"day": IP_FIELDS[F.DAY],
                                     "month": IP_FIELDS[F.MONTH], "is_ip": True},
        text="", save=False,
    )

    contract_data_ooo = dict(OOO_FIELDS)
    contract_data_ip = dict(IP_FIELDS)
    limit_display, term = "1 500 000", "45"

    out = {}
    out["contract_ooo"] = document_signature(
        services.render_contract(contract_data_ooo, is_ip=False))
    out["contract_ip"] = document_signature(
        services.render_contract(contract_data_ip, is_ip=True))
    out["additional_ooo"] = document_signature(services.render_additional(
        contract_data_ooo, False, "2", limit_display, term))
    out["additional_ip"] = document_signature(services.render_additional(
        contract_data_ip, True, "3", limit_display, term))
    out["spec_ooo"] = document_signature(
        services.render_spec(spec_rec(party_ooo, contract_ooo), store))
    out["spec_ip"] = document_signature(
        services.render_spec(spec_rec(party_ip, contract_ip), store))
    return out


def vat_for_items() -> str:
    """НДС «в том числе» по позициям образца (для проверки ключевых сумм)."""
    return format_amount(spec_vat_amount_from_items(SPEC_ITEMS))
