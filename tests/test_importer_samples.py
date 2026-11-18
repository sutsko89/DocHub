"""Импорт на образцах документов (п. 7.4) и матрица ИП (п. 7.5, импорт).

Образцы .docx/.xlsx строятся кодом (openpyxl/python-docx): договор ООО,
договор ИП (матрица ИП), спецификация .xlsx, дубликаты ИНН, неполные
реквизиты, .doc, битый файл.
"""
from __future__ import annotations

import os

from docx import Document

import document_importer as di
import storage
from document_builder import F
from render_helpers import IP_FIELDS, OOO_FIELDS


def _make_contract_docx(path: str, fields: dict, is_ip: bool = False,
                        number: str = "777", with_inn: bool = True) -> None:
    doc = Document()
    doc.add_paragraph(f"ДОГОВОР ПОСТАВКИ № {number}")
    doc.add_paragraph("г. Владивосток")
    doc.add_paragraph("«10» января 2026 г.")
    if is_ip:
        doc.add_paragraph(
            f"Индивидуальный предприниматель {fields[F.IP_NAME]}, "
            "именуемый в дальнейшем «Покупатель»,"
        )
        if with_inn:
            doc.add_paragraph(
                f"ИНН {fields[F.INN]}, ОГРНИП {fields[F.OGRNIP]}"
            )
    else:
        doc.add_paragraph(
            "ООО «Ромашка-Импорт», именуемый в дальнейшем «Покупатель»,"
        )
        if with_inn:
            doc.add_paragraph(
                f"ИНН {fields[F.INN]}, КПП {fields[F.KPP]}, "
                f"ОГРН {fields[F.OGRN]}"
            )
    doc.add_paragraph(f"Адрес: {fields[F.LEGAL_ADDRESS]}")
    doc.add_paragraph(
        f"Р/с {fields[F.ACCOUNT]} в {fields[F.BANK]}, "
        f"БИК {fields[F.BIK]}, к/с {fields[F.CORR_ACCOUNT]}"
    )
    doc.add_paragraph("Генеральный директор Петров Пётр Петрович")
    doc.save(path)


def _make_spec_xlsx(path: str, inn: str = "2540123456") -> None:
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "СПЕЦИФИКАЦИЯ № 12 к Договору поставки № 516"
    ws["A2"] = "от 10 января 2026 г."
    ws["A3"] = "ООО «Ромашка-Импорт»"
    ws["A4"] = f"ИНН {inn}"
    ws["A5"] = "№  Наименование  Кол-во  Цена  Сумма"
    ws["A6"] = "1  Кабель 120 м × 85.50 = 10260.00"
    wb.save(path)


# ---------- Договор ООО ----------

def test_contract_ooo_imported(tmp_path, ooo_fields):
    path = str(tmp_path / "contract_ooo.docx")
    _make_contract_docx(path, ooo_fields, is_ip=False)

    result = di.analyze_file(path)
    assert result["ok"] is True and result["recognized"] is True
    assert result["kind"] == storage.KIND_CONTRACT
    assert result["number"] == "777"
    assert result["date"] == "10.01.2026"
    assert result["party"]["type"] == "ooo"
    assert result["party"]["inn"] == ooo_fields[F.INN]
    assert result["party"]["fields"][F.ACCOUNT] == ooo_fields[F.ACCOUNT]
    assert result["error"] == ""


# ---------- Матрица ИП (импорт) ----------

def test_contract_ip_imported(tmp_path, ip_fields):
    path = str(tmp_path / "contract_ip.docx")
    _make_contract_docx(path, ip_fields, is_ip=True, number="77")

    result = di.analyze_file(path)
    assert result["ok"] is True and result["recognized"] is True
    assert result["party"]["type"] == "ip"
    assert result["party"]["inn"] == ip_fields[F.INN]
    fields = result["party"]["fields"]
    assert fields[F.OGRNIP] == ip_fields[F.OGRNIP]
    assert fields.get(F.KPP, "") == ""  # у ИП нет КПП


def test_ip_contract_without_inn_not_recognized(tmp_path, ip_fields):
    path = str(tmp_path / "contract_ip_no_inn.docx")
    _make_contract_docx(path, ip_fields, is_ip=True, with_inn=False)

    result = di.analyze_file(path)
    assert result["recognized"] is False
    assert result["party"] is None or not result["party"].get("recognized")


# ---------- Спецификация .xlsx ----------

def test_spec_xlsx_imported(tmp_path):
    path = str(tmp_path / "spec.xlsx")
    _make_spec_xlsx(path)

    result = di.analyze_file(path)
    assert result["kind"] == storage.KIND_SPEC
    assert result["number"] == "12"
    assert result["party"] is not None
    assert result["party"]["inn"] == "2540123456"


# ---------- Дубликаты ИНН ----------

def test_duplicate_inn_matches_same_party(store, tmp_path, ooo_fields):
    path_a = str(tmp_path / "doc_a.docx")
    path_b = str(tmp_path / "doc_b.docx")
    _make_contract_docx(path_a, ooo_fields, number="777")
    _make_contract_docx(path_b, ooo_fields, number="778")

    result_a = di.analyze_file(path_a)
    result_b = di.analyze_file(path_b)
    inn = result_a["party"]["inn"]
    assert inn == result_b["party"]["inn"]

    # Первая карточка создаётся, вторая должна быть привязана к ней же,
    # а не плодить дубль (п. 4.1 — решение за пользователем, здесь реестр).
    party = store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )
    existing = store.find_by_inn(inn)
    assert existing is not None
    assert existing["id"] == party["id"]


# ---------- Неполные реквизиты ----------

def test_contract_without_inn_keeps_warning(tmp_path, ooo_fields):
    path = str(tmp_path / "no_inn.docx")
    _make_contract_docx(path, ooo_fields, with_inn=False)

    result = di.analyze_file(path)
    assert result["ok"] is True  # тип и номер опознаны
    assert result["recognized"] is False  # карточку создаёт только человек
    assert result["warning"], "должно быть объяснение, чего не хватило"


# ---------- .doc и битые файлы ----------

def test_doc_extension_rejected(tmp_path):
    path = tmp_path / "old.doc"
    path.write_bytes(b"\xd0\xcf\x11\xe0stub")
    result = di.analyze_file(str(path))
    assert result["ok"] is False
    assert ".doc" in result["error"]


def test_broken_docx_rejected(tmp_path):
    path = tmp_path / "broken.docx"
    path.write_text("это не zip-архив документа", encoding="utf-8")
    result = di.analyze_file(str(path))
    assert result["ok"] is False
    assert result["error"]


def test_empty_file_rejected(tmp_path):
    path = tmp_path / "empty.docx"
    path.write_bytes(b"")
    result = di.analyze_file(str(path))
    assert result["ok"] is False
    assert "пустой" in result["error"]
