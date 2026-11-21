"""Модульные тесты чистого конвейера печатных форм (document_builder.py).

Покрывают расчёт НДС «в том числе» (п. 7.1 — как в 1С), прописью,
разбор чисел, замену плейсхолдеров одним проходом (п. 5.6) и roundtrip
текста документа.
"""
from __future__ import annotations

from decimal import Decimal

import pytest
from docx import Document

from document_builder import (
    F, PH, amount_in_words_ru, amount_to_words_ru, build_replacements,
    digits_only, docx_from_text, docx_to_text, fill_spec_items,
    format_amount, format_date, format_phone, number_to_words_ru,
    parse_decimal, prices_valid_until_default, replace_in_paragraph,
    spec_line_total, spec_total, spec_vat_amount, spec_vat_amount_from_items,
    unfilled_placeholders,
)


# ---------- НДС «в том числе» ----------

def test_vat_included_from_total():
    # 1220 с НДС 22 %: НДС = 1220 * 22 / 122 = 220
    assert spec_vat_amount(1220) == Decimal("220.00")
    assert spec_vat_amount("1 220,00") == Decimal("220.00")


def test_vat_zero_rate():
    assert spec_vat_amount(1220, 0) == Decimal("0.00")


def test_vat_from_items_is_line_by_line():
    items = [{"qty": 3, "price": "100.50"}, {"qty": 2, "price": 50}]
    # позиция 1: 301.50 * 22/122 = 54.37; позиция 2: 100 * 22/122 = 18.03
    assert spec_vat_amount_from_items(items) == Decimal("72.40")


# ---------- Единый путь расчёта итога (п. 9.8) ----------

def test_line_total_quantizes_to_kopecks():
    item = {"name": "Кабель", "qty": 3, "unit": "м", "price": "3,335"}
    assert spec_line_total(item) == Decimal("10.01")  # 10.005 → 10.01 (HALF_UP)


def test_spec_total_is_sum_of_rounded_lines_not_raw_sum():
    """Копеечный раскол: сырая сумма 11,12, сумма строк 11,13 — итог 11,13.

    Печатная форма показывает строки по spec_line_total; «Итого», посчитанный
    по сырым суммам, расходился с таблицей на копейку (п. 9.8).
    """
    items = [
        {"name": "A", "qty": 3, "unit": "м", "price": "3.335"},
        {"name": "B", "qty": 1, "unit": "шт", "price": "1.115"},
    ]
    raw = sum(parse_decimal(i["qty"]) * parse_decimal(i["price"]) for i in items)
    assert raw == Decimal("11.120")
    assert spec_total(items) == Decimal("11.13")
    # и НДС по строкам соответствует тому же итогу
    assert spec_vat_amount_from_items(items) == Decimal("2.01")


def test_spec_total_empty_items_is_zero():
    assert spec_total([]) == Decimal("0")
    assert spec_total(None) == Decimal("0")


def test_spec_total_accepts_floats_and_strings():
    """Позиции из реестра (float) и из таблицы (Decimal/строки) — один результат."""
    floats = [{"qty": 2.0, "price": 100.5}]
    strings = [{"qty": "2", "price": "100,50"}]
    decimals = [{"qty": parse_decimal("2"), "price": parse_decimal("100.5")}]
    assert spec_total(floats) == Decimal("201.00")
    assert spec_total(strings) == Decimal("201.00")
    assert spec_total(decimals) == Decimal("201.00")


# ---------- Числа и суммы ----------

def test_parse_decimal():
    assert parse_decimal("12,5") == Decimal("12.5")
    assert parse_decimal("1 000.5") == Decimal("1000.5")
    assert parse_decimal("1\xa0000") == Decimal("1000")


@pytest.mark.parametrize("bad", ["abc", "", None, "{{total}}"])
def test_parse_decimal_garbage_is_zero(bad):
    assert parse_decimal(bad) == Decimal("0")


def test_format_amount_grouped():
    assert format_amount(1500000) == "1 500 000"
    assert format_amount(1500000.5) == "1 500 000,50"
    assert format_amount("1 500 000,00") == "1 500 000"


def test_number_words():
    assert number_to_words_ru(700) == "семьсот"
    assert number_to_words_ru(1500000) == "один миллион пятьсот тысяч"
    assert number_to_words_ru(21) == "двадцать один"
    assert number_to_words_ru(0) == "ноль"


def test_amount_words():
    assert amount_to_words_ru(150000) == "сто пятьдесят тысяч рублей"
    # для печатной формы — с заглавной и копейками
    assert amount_in_words_ru(150000) == "Сто пятьдесят тысяч рублей 00 копеек"


def test_digits_only_keeps_numbers():
    assert digits_only("Р/с № 40702-810") == "40702810"


def test_format_phone_mask():
    assert format_phone("9123456789") == "+7 (912) 345-67-89"
    assert format_phone("89121234567") == "+7 (912) 123-45-67"
    assert format_phone("") == ""


def test_format_date():
    assert format_date("10") == "10"
    assert format_date("5") == "05"  # ведущий ноль
    assert format_date("abc") == "abc"


# ---------- Замена плейсхолдеров одним проходом ----------

def test_replace_in_paragraph_multiple_placeholders():
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("Договор № {{number}} от {{date}}, ИНН {{INN}}")
    replace_in_paragraph(paragraph, {
        PH.NUMBER: "516", PH.DATE: "10 января", PH.INN: "2540123456",
    })
    assert paragraph.text == "Договор № 516 от 10 января, ИНН 2540123456"


def test_replace_in_paragraph_placeholder_split_across_runs():
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("Договор № ")
    paragraph.add_run("{{num")
    paragraph.add_run("ber}} готов")
    replace_in_paragraph(paragraph, {PH.NUMBER: "7"})
    assert paragraph.text == "Договор № 7 готов"


def test_replace_in_paragraph_without_runs_is_noop():
    doc = Document()
    paragraph = doc.add_paragraph("")
    replace_in_paragraph(paragraph, {PH.NUMBER: "7"})  # не бросает исключений


# ---------- Текст документа туда-обратно ----------

def test_docx_text_roundtrip():
    doc = docx_from_text("Строка 1\nСтрока 2")
    assert docx_to_text(doc) == "Строка 1\nСтрока 2"


# ---------- Словарь замен ----------

def test_build_replacements_ip_has_no_kpp_key():
    data = {
        F.NUMBER: "77", F.DAY: "12", F.MONTH: "февраля",
        F.IP_NAME: "Иванов Иван Иванович", F.INN: "254012345678",
        F.OGRNIP: "304253600012345", F.LEGAL_ADDRESS: "г. Владивосток",
        F.FACT_ADDRESS: "", F.EMAIL: "", F.PHONE: "",
        F.ACCOUNT: "40802810500000012345", F.CORR_ACCOUNT: "30101810100000000225",
        F.BANK: "Банк", F.BIK: "044525225",
    }
    replacements = build_replacements(data, is_ip=True)
    assert replacements[PH.OGRNIP] == "304253600012345"
    assert PH.KPP not in replacements and PH.OGRN not in replacements


def test_build_replacements_ooo_basis_ustav():
    data = {F.NUMBER: "516", F.DAY: "10", F.MONTH: "января", F.BASIS: "Устав"}
    replacements = build_replacements(data, is_ip=False)
    assert replacements[PH.BASIS] == "устава"
    assert replacements[PH.NUMBER] == "516"


# ---------- Даты ----------

def test_prices_valid_until_is_end_of_current_month():
    from datetime import datetime
    import calendar
    today = datetime.now()
    last_day = calendar.monthrange(today.year, today.month)[1]
    expected = f"{last_day:02d}.{today.month:02d}.{today.year} г."
    assert prices_valid_until_default() == expected


# ---------- D5: кросс-run замена позиций и остаточные {{…}} ----------

def _spec_table_doc():
    """Таблица спецификации: заголовок + строка-шаблон с плейсхолдерами."""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    header = table.rows[0].cells
    header[0].paragraphs[0].add_run("№ п/п")
    header[1].paragraphs[0].add_run("Наименование")
    row = table.rows[1]
    row.cells[0].paragraphs[0].add_run("{{item_no}}")
    paragraph = row.cells[1].paragraphs[0]
    paragraph.add_run("{{it")   # токен, разбитый Word'ом на три run
    paragraph.add_run("em_")
    paragraph.add_run("name}}")
    return doc


def test_fill_spec_items_replaces_token_split_across_runs():
    doc = _spec_table_doc()
    fill_spec_items(doc, [{"name": "Болт М8", "qty": "5", "unit": "шт.",
                           "price": "200"}])
    row_text = " | ".join(cell.text for cell in doc.tables[0].rows[1].cells)
    assert row_text == "1 | Болт М8", "разбитый токен заменяется полностью"
    whole = "\n".join(cell.text for row in doc.tables[0].rows for cell in row.cells)
    assert "{{" not in whole


def test_fill_spec_items_multi_items_keeps_numbering():
    doc = _spec_table_doc()
    fill_spec_items(doc, [
        {"name": "Болт", "qty": "1", "unit": "шт.", "price": "10"},
        {"name": "Гайка", "qty": "2", "unit": "шт.", "price": "20"},
    ])
    numbers = [doc.tables[0].rows[i].cells[0].text for i in (1, 2)]
    assert numbers == ["1", "2"], "клонированные строки нумеруются по порядку"


def test_unfilled_placeholders_finds_body_tables_and_header():
    doc = Document()
    doc.add_paragraph("Реквизиты: {{INN}}")
    table = doc.add_table(rows=1, cols=1)
    table.rows[0].cells[0].text = "{{item_1}}"
    doc.sections[0].header.paragraphs[0].add_run("{{PHONE}}")
    assert unfilled_placeholders(doc) == ["{{INN}}", "{{PHONE}}", "{{item_1}}"]


def test_unfilled_placeholders_clean_document_is_empty():
    doc = Document()
    doc.add_paragraph("Договор № 516 готов")
    assert unfilled_placeholders(doc) == []
