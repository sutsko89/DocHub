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
