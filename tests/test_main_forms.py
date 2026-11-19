"""Юнит-тесты сборщиков данных форм main.py (аудит 2.0.1, P1 «старт покрытия»).

ContractApp не создаётся целиком — методы вызываются с фейковыми виджетами
через ``ContractApp.method(fake_app, ...)``. Проверяются контракты из аудита:
  * _cp_collect возвращает None при сбое чтения поля (не пустую строку);
  * _form_signature деградирует до None (страница «грязная», а не «чистая»);
  * _cp_save при сбое сборки показывает ошибку и не пишет реестр.
"""
from __future__ import annotations

import logging

import pytest

import main
from document_builder import F


class FakeVar:
    """Минимальный аналог tk.StringVar для подписей форм."""

    def __init__(self, value: str = ""):
        self._value = value

    def get(self) -> str:
        return self._value


class FakeEntry:
    """Минимальный аналог поля ввода."""

    def __init__(self, value: str = ""):
        self._value = value

    def get(self) -> str:
        return self._value


class BrokenEntry:
    """Поле, чтение которого падает (уничтоженный виджет, сбой Tk)."""

    def get(self):
        raise RuntimeError("widget destroyed")


def make_app(entries: dict) -> "main.ContractApp":
    """«Пустой» ContractApp без __init__: только то, что нужно сборщикам.

    Подкласс нужен потому, что у ContractApp._spec_rows — property без
    сеттера (она читает живой виджет SpecTable); в тестах это обычный атрибут.
    """
    app = object.__new__(_FormHarness)
    app._cp_entries = entries
    app._cp_type_var = FakeVar("Юр. лицо")
    app._log = logging.getLogger("test-main-forms")
    app._form_baseline = {}
    app.root = None  # алерты замокаются; настоящий root не нужен
    return app


class _FormHarness(main.ContractApp):
    """Класс-подложка: свойства для живого UI заменены обычными атрибутами."""

    _spec_rows: list = []  # перекрывает property экземпляра — тест задаёт сам


# --------------------------------------------------------------------------- #
#  _cp_collect
# --------------------------------------------------------------------------- #
def test_collect_reads_values_and_strips():
    app = make_app({"ooo": {F.INN: FakeEntry(" 7701234567 "), F.KPP: FakeEntry("770101001")}})
    data = main.ContractApp._cp_collect(app, "ooo")
    assert data == {F.INN: "7701234567", F.KPP: "770101001"}


def test_collect_failure_returns_none_not_empty_string():
    app = make_app({"ooo": {F.INN: FakeEntry("7701234567"), F.KPP: BrokenEntry()}})
    assert main.ContractApp._cp_collect(app, "ooo") is None


def test_collect_ip_copies_legal_address_to_fact():
    app = make_app({"ip": {F.LEGAL_ADDRESS: FakeEntry("г. Москва")}})
    data = main.ContractApp._cp_collect(app, "ip")
    assert data[F.FACT_ADDRESS] == "г. Москва"


# --------------------------------------------------------------------------- #
#  _form_signature
# --------------------------------------------------------------------------- #
def test_signature_counterparty_is_sorted_tuple():
    app = make_app({"ooo": {F.INN: FakeEntry("7701234567"), F.KPP: FakeEntry("x")}})
    signature = main.ContractApp._form_signature(app, "counterparty")
    assert signature == tuple(sorted({F.INN: "7701234567", F.KPP: "x"}.items()))


def test_signature_counterparty_degrades_to_none():
    app = make_app({"ooo": {F.INN: BrokenEntry()}})
    assert main.ContractApp._form_signature(app, "counterparty") is None


def test_signature_contract_reads_vars():
    app = make_app({})
    app.contract_number_var = FakeVar("5")
    app.contract_day_var = FakeVar("12")
    app.contract_month_var = FakeVar("января")
    assert main.ContractApp._form_signature(app, "contract") == ("5", "12", "января")


def test_signature_spec_includes_items():
    app = make_app({})
    app._spec_rows = [{
        "name": FakeVar("Болт"), "qty": FakeVar("5"),
        "unit": FakeVar("шт."), "price": FakeVar("200"),
    }]
    app.spec_number_var = FakeVar("1")
    app.spec_day_var = FakeVar("01")
    app.spec_month_var = FakeVar("января")
    app.spec_delivery_term_var = FakeVar("10 дней")
    app.spec_payment_terms_var = FakeVar("предоплата")
    app.spec_delivery_basis_var = FakeVar("самовывоз")
    app.spec_prices_until_var = FakeVar("31.12.2026")
    signature = main.ContractApp._form_signature(app, "spec")
    assert signature[7] == (("Болт", "5", "шт.", "200"),)


def test_signature_missing_attribute_degrades_to_none():
    app = make_app({})  # contract-поля не заданы → AttributeError
    assert main.ContractApp._form_signature(app, "contract") is None


# --------------------------------------------------------------------------- #
#  Защита несохранённых данных (None ≠ baseline)
# --------------------------------------------------------------------------- #
def test_form_becomes_dirty_after_widget_failure():
    entries = {"ooo": {F.INN: FakeEntry("7701234567")}}
    app = make_app(entries)
    app._form_baseline["counterparty"] = main.ContractApp._form_signature(app, "counterparty")
    assert not main.ContractApp._is_form_dirty(app, "counterparty")

    entries["ooo"][F.INN] = BrokenEntry()  # виджет сломался посреди сессии
    assert main.ContractApp._form_signature(app, "counterparty") is None
    # None не равен baseline → страница «грязная», данные защищены.
    assert main.ContractApp._is_form_dirty(app, "counterparty")


# --------------------------------------------------------------------------- #
#  _cp_save: сбой сборки — ошибка пользователю, ничего не пишется
# --------------------------------------------------------------------------- #
def test_cp_save_reports_collect_failure_and_saves_nothing(monkeypatch):
    alerts: list[tuple] = []
    monkeypatch.setattr(main.ui, "alert", lambda *args, **kwargs: alerts.append(args))
    app = make_app({"ooo": {F.INN: BrokenEntry()}})
    main.ContractApp._cp_save(app)
    assert alerts, "пользователю показано сообщение о сбое"
    assert "Не удалось прочитать поля формы" in alerts[0][1]
    assert not hasattr(app, "store"), "до хранилища дело не дошло"


def test_cp_save_with_valid_data_passes_validation(monkeypatch):
    alerts: list[tuple] = []
    monkeypatch.setattr(main.ui, "alert", lambda *args, **kwargs: alerts.append(args))
    app = make_app({"ooo": {
        F.INN: FakeEntry("7701234567"),
        F.KPP: FakeEntry("770101001"),
        F.OGRN: FakeEntry("1037739010221"),
        F.LEGAL_ADDRESS: FakeEntry("г. Москва"),
        F.ACCOUNT: FakeEntry("40702810123456789012"),
        F.CORR_ACCOUNT: FakeEntry("30101810100000000225"),
        F.BANK: FakeEntry("Банк"),
        F.BIK: FakeEntry("044525225"),
        F.FULL_NAME: FakeEntry("ООО"),
        F.SHORT_NAME: FakeEntry("ООО"),
        F.DIRECTOR: FakeEntry("Иванов И.И."),
    }})

    class _StoreStop(Exception):
        pass

    class _FakeStore:
        last_error = ""

        def __init__(self):
            self.added = None

        def find_by_inn(self, inn):
            return None

        def add_counterparty(self, mode, data):
            self.added = (mode, dict(data))
            raise _StoreStop()  # останавливаем поток до записи на диск

    app.store = _FakeStore()
    with pytest.raises(_StoreStop):
        main.ContractApp._cp_save(app)
    assert not alerts, "валидация валидных данных не показывает ошибок"
    assert app.store.added[0] == "ooo"
    assert app.store.added[1][F.INN] == "7701234567"
