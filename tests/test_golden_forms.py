"""Golden-тесты шести печатных форм (п. 7.1 плана).

ООО/ИП × договор/доп. соглашение/спецификация: текст, таблицы и ключевые
суммы сверяются с эталонами из tests/golden/*.json. Отдельно проверяется
отсутствие незаполненных плейсхолдеров {{...}} — это ловит баг 1.1
(пустые «ОГРН»/«КПП» в печатной форме ИП).

Эталоны пересоздаются осознанно: python tests/make_golden.py
(и коммитятся вместе с изменением шаблона или подстановок).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from document_builder import format_amount
from render_helpers import (
    IP_FIELDS, OOO_FIELDS, SPEC_ITEMS, render_six_forms, spec_total,
    vat_for_items,
)

GOLDEN_DIR = Path(__file__).resolve().parent / "golden"
FORMS = (
    "contract_ooo", "contract_ip", "additional_ooo", "additional_ip",
    "spec_ooo", "spec_ip",
)


@pytest.fixture(scope="module")
def signatures() -> dict:
    return render_six_forms()


def _golden(name: str) -> dict:
    path = GOLDEN_DIR / f"{name}.json"
    return json.loads(path.read_text(encoding="utf-8"))


@pytest.mark.parametrize("form", FORMS)
def test_form_matches_golden(signatures, form):
    expected = _golden(form)
    actual = signatures[form]
    assert actual["text"] == expected["text"], (
        f"Текст печатной формы {form} изменился. "
        "Если изменение осознанное — перегенерируйте эталон: python tests/make_golden.py"
    )
    assert actual["tables"] == expected["tables"], (
        f"Таблицы печатной формы {form} изменились"
    )


@pytest.mark.parametrize("form", FORMS)
def test_form_has_no_unfilled_placeholders(signatures, form):
    signature = signatures[form]
    assert "{{" not in signature["text"], f"{form}: остались незаполненные {{{{...}}}}"
    for row in signature["tables"]:
        for cell in row:
            assert "{{" not in cell, f"{form}: в таблице остались {{{{...}}}}"


def test_contract_key_fields_present(signatures):
    text = signatures["contract_ooo"]["text"]
    assert OOO_FIELDS["ИНН"] in text and OOO_FIELDS["ОГРН"] in text
    assert OOO_FIELDS["Расчётный счёт"] in text and OOO_FIELDS["БИК"] in text
    assert "№ 516" in text and "10 января" in text


def test_contract_ip_requisites_present(signatures):
    """Регрессия бага 1.1: печатная форма ИП без пустых плейсхолдеров."""
    text = signatures["contract_ip"]["text"]
    assert IP_FIELDS["ОГРНИП"] in text
    assert IP_FIELDS["ИНН"] in text
    assert "{{KPP}}" not in text and "{{OGRN}}" not in text
    assert "{{OGRNIP}}" not in text


def test_spec_ip_buyer_requisites_no_kpp(signatures):
    """Регрессия бага 1.1: в реквизитах покупателя-ИП нет КПП и пустых подписей."""
    tables = signatures["spec_ip"]["tables"]
    buyer = next(
        (cell for row in tables for cell in row if "Покупател" in cell), ""
    )
    assert buyer, "блок реквизитов покупателя не найден"
    assert "КПП" not in buyer
    assert "ОГРНИП 304253600012345" in buyer
    assert re.search(r"\bОГРН\b(?!ИП)", buyer) is None


def test_spec_key_sums_present(signatures):
    """Ключевые суммы спецификации: итог и НДС «в том числе» по позициям."""
    text = signatures["spec_ooo"]["text"]
    total = format_amount(spec_total())
    vat = vat_for_items()
    assert total in text, f"итог {total} не найден в печатной форме"
    assert vat in text, f"НДС {vat} не найден в печатной форме"
    assert "НДС 22%" in text


def test_spec_items_table_filled(signatures):
    tables = signatures["spec_ooo"]["tables"]
    flat = [cell for row in tables for cell in row]
    assert any(SPEC_ITEMS[0]["name"] in cell for cell in flat), \
        "позиция спецификации не попала в таблицу"
    assert any("120" in cell for cell in flat), "количество не попало в таблицу"
