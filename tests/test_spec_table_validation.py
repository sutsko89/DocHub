"""Валидация ячеек SpecTable (аудит 2.0.1, P2 «мёртвые валидаторы»).

Валидаторы — чистые статические функции: проверяются без GUI. Живой Tk-контекст
(validate="key" на Entry) закрывается smoke-тестом.
"""
from __future__ import annotations

from widgets import SpecTable


def test_digits_validator():
    assert SpecTable._validate_digits("")
    assert SpecTable._validate_digits("12")
    assert SpecTable._validate_digits("12,5")
    assert SpecTable._validate_digits("12.5")
    assert not SpecTable._validate_digits("abc")
    assert not SpecTable._validate_digits("12а")  # кириллица не пролезает
    assert not SpecTable._validate_digits("1 2")


def test_price_validator():
    assert SpecTable._validate_price("1 000,50")   # разряды и копейки
    assert SpecTable._validate_price("150.00")
    assert SpecTable._validate_price("")
    assert not SpecTable._validate_price("150 руб")
    assert not SpecTable._validate_price("1,5шт")


def test_sanitize_qty():
    assert SpecTable._sanitize_value("qty", "abc") == ""      # мусор отсекается
    assert SpecTable._sanitize_value("qty", "12,5") == "12.5"  # нормализация
    assert SpecTable._sanitize_value("qty", "42") == "42"


def test_sanitize_price():
    assert SpecTable._sanitize_value("price", "1 000,50") == "1000.50"
    # «Единицы измерения» из буфера Excel аккуратно отчищаются до числа.
    assert SpecTable._sanitize_value("price", "150 руб") == "150"


def test_sanitize_name_untouched():
    assert SpecTable._sanitize_value("name", "Болт М8 оцинк.") == "Болт М8 оцинк."
    assert SpecTable._sanitize_value("unit", "шт.") == "шт."


def test_sanitize_survives_repeated_separators():
    assert SpecTable._sanitize_value("price", "1.000.000") == "1000000"
    assert SpecTable._sanitize_value("qty", "2.052,59") == "2052.59"


# --------------------------------------------------------------------------- #
#  PERF-1: лимит вставки блока из Excel (D2/R-08)
# --------------------------------------------------------------------------- #
def _matrix(rows: int) -> list[list[str]]:
    return [[f"Позиция {n}", "1", "шт.", "10"] for n in range(1, rows + 1)]


def test_paste_plan_unlimited_block_is_clamped_to_max_rows():
    usable, dropped = SpecTable._paste_plan(_matrix(600), 0, SpecTable.MAX_ROWS)
    assert len(usable) == 500 and dropped == 100


def test_paste_plan_small_block_unchanged():
    usable, dropped = SpecTable._paste_plan(_matrix(10), 0, SpecTable.MAX_ROWS)
    assert len(usable) == 10 and dropped == 0


def test_paste_plan_clamp_respects_start_row():
    # Вставка в непустую таблицу со сдвигом r0: 250 + 300 → в лимит входят 250.
    usable, dropped = SpecTable._paste_plan(_matrix(300), 250, SpecTable.MAX_ROWS)
    assert len(usable) == 250 and dropped == 50


def test_paste_plan_start_row_beyond_limit_keeps_nothing():
    usable, dropped = SpecTable._paste_plan(_matrix(50), 600, SpecTable.MAX_ROWS)
    assert usable == [] and dropped == 50


def test_paste_plan_exactly_to_limit_no_drop():
    usable, dropped = SpecTable._paste_plan(_matrix(500), 0, SpecTable.MAX_ROWS)
    assert len(usable) == 500 and dropped == 0, "вставка ровно до лимита без обрезки"


def test_paste_notice_message_names_both_numbers():
    message = SpecTable.paste_notice_text(500, 100, SpecTable.MAX_ROWS)
    assert "Вставлено 500" in message and "отброшено 100" in message
    assert "500 строк" in message, "назван лимит"
