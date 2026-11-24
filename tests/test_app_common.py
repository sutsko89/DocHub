"""app_common — GUI-free константы и настройки приложения (этап M1 миграции).

Точечные тесты нового модуля (план миграции, M1):

* версия — чисто числовая «X.Y.Z» (регулярка EWDochub.spec и
  updater.parse_version не ломаются), флейвор сборки — "tk";
* settings.json: чтение/атомарная запись, честный False при сбое записи;
* печатные настройки спецификации: год «на момент вызова» (аудит 2.0.1),
  кламп spec_vat_rate, приведение типов, мусорные значения → дефолты;
* хелперы каретки и телефона — на duck-typed фейках, без tkinter:
  модуль остаётся GUI-free, контроллеры импортируют его напрямую
  (поздние ``import main`` ликвидированы).

Реэкспорт имён через main.py (контракт smoke_test.py) покрывают
существующие проверки smoke — здесь он не дублируется.
"""
from __future__ import annotations

import datetime as dt
import re
from types import SimpleNamespace

import app_common
from document_builder import SPEC_VAT_RATE


# --------------------------------------------------------------------------- #
#  Версия и флейвор сборки
# --------------------------------------------------------------------------- #
def test_version_is_numeric_and_flavor_is_tk():
    """Версия «X.Y.Z» без суффиксов — контракт EWDochub.spec и И.2 плана."""
    assert re.fullmatch(r"\d+\.\d+\.\d+", app_common.APP_VERSION)
    assert app_common.APP_UI_FLAVOR == "tk"


# --------------------------------------------------------------------------- #
#  settings.json: чтение, атомарная запись, сбой
# --------------------------------------------------------------------------- #
def test_settings_roundtrip(appdata):
    assert app_common.load_settings() == {}
    assert app_common.write_settings({"theme": "dark", "auto_update_check": False})
    assert app_common.load_settings() == {"theme": "dark", "auto_update_check": False}


def test_write_settings_failure_returns_false(appdata):
    """Каталог вместо файла: os.replace падает — честный False без исключения."""
    import os

    os.makedirs(app_common.settings_path())  # каталог вместо файла настроек
    assert app_common.write_settings({"theme": "light"}) is False
    assert app_common.load_settings() == {}, "при сбое настройки читаются пустыми"


# --------------------------------------------------------------------------- #
#  Печатные настройки спецификации
# --------------------------------------------------------------------------- #
def test_get_spec_settings_defaults_and_year_on_call(appdata):
    settings = app_common.get_spec_settings()
    assert settings["spec_year"] == dt.datetime.now().year, \
        "год берётся на момент вызова, а не импорта модуля (аудит 2.0.1)"
    assert settings["spec_vat_rate"] == SPEC_VAT_RATE
    assert settings["spec_supplier"] and settings["spec_city"]
    assert settings["spec_appendix_number"]


def test_get_spec_settings_clamps_vat_and_coerces_types(appdata):
    app_common.write_settings({
        "spec_vat_rate": 150,      # кламп до 100
        "spec_supplier": 12345,    # число → строка
        "spec_year": "мусор",      # битый год — дефолт «на момент вызова»
    })
    settings = app_common.get_spec_settings()
    assert settings["spec_vat_rate"] == 100
    assert settings["spec_supplier"] == "12345"
    assert settings["spec_year"] == dt.datetime.now().year


def test_get_spec_settings_broken_vat_falls_back_to_default(appdata):
    app_common.write_settings({"spec_vat_rate": "abc"})
    assert app_common.get_spec_settings()["spec_vat_rate"] == SPEC_VAT_RATE


# --------------------------------------------------------------------------- #
#  Хелперы каретки и телефона — duck-typed, без tkinter
# --------------------------------------------------------------------------- #
class _FakeEntry:
    """Фейк Tk-Entry: after_idle применяет сразу, icursor записывает вызовы."""

    def __init__(self):
        self.calls: list[int] = []

    def after_idle(self, fn):
        fn()

    def icursor(self, index):
        self.calls.append(index)


class _EntryWithoutCaret:
    """Виджет без доступа к каретке: index() бросает — путь «caret = None»."""

    def __init__(self):
        self.calls: list[int] = []

    def index(self, *_args):
        raise RuntimeError("каретка недоступна")

    def after_idle(self, fn):
        fn()

    def icursor(self, index):
        self.calls.append(index)


class _FakeVar:
    def __init__(self, value):
        self._value = value
        self.set_calls: list[str] = []

    def get(self):
        return self._value

    def set(self, value):
        self.set_calls.append(value)
        self._value = value


def test_restore_caret_after_reformat_shifts_cursor():
    entry = _FakeEntry()
    app_common.restore_caret_after_reformat(entry, 5, 10, 14)
    assert entry.calls == [9], "курсор сдвинут на дельту длины и зажат в границах"


def test_restore_caret_after_reformat_noop_without_entry_or_caret():
    app_common.restore_caret_after_reformat(None, 5, 10, 14)  # не бросает
    entry = _FakeEntry()
    app_common.restore_caret_after_reformat(entry, None, 10, 14)
    assert entry.calls == []


def test_restore_caret_after_reformat_clamps_negative_and_overflow():
    entry = _FakeEntry()
    app_common.restore_caret_after_reformat(entry, 0, 10, 4)   # 0 + 4 - 10 = -6 → 0
    app_common.restore_caret_after_reformat(entry, 20, 10, 4)  # 20 + 4 - 10 → 4
    assert entry.calls == [0, 4]


def test_make_phone_formatter_reformats_8_prefix():
    var = _FakeVar("89121234567")
    entry = _EntryWithoutCaret()  # index("insert") бросает — каретка None
    app_common.make_phone_formatter(var, entry)()
    assert var.set_calls == ["+7 (912) 123-45-67"], "8 → +7, переформатирование"


def test_make_phone_formatter_noop_when_already_formatted():
    var = _FakeVar("+7 (912) 345-67-89")
    app_common.make_phone_formatter(var)()  # entry=None — не нужен
    assert var.set_calls == [], "переформатирование «в пустоту» не выполняется"


def test_make_phone_formatter_uses_literal_insert_index():
    """entry.index вызывается с литералом «insert» (== tk.INSERT): модуль
    GUI-free, поведение то же."""
    seen: list[str] = []

    class _SpyEntry:
        def index(self, name):
            seen.append(name)
            raise RuntimeError("каретка недоступна")

        def after_idle(self, fn):
            fn()

        def icursor(self, index):
            pass

    var = _FakeVar("9123456789")
    app_common.make_phone_formatter(var, _SpyEntry())()
    assert seen == ["insert"]
    assert var.set_calls == ["+7 (912) 345-67-89"]
