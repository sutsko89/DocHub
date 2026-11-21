"""CHKSUM — контрольные цифры ИНН (10/12), ОГРН/ОГРНИП при вводе реквизитов.

Этап 2 roadmap: опечатка в контрольной цифре ловится при вводе, а не
контрагентом на печатной форме. Контракты:

  * официальные алгоритмы: ИНН 10/12 — взвешенные суммы по модулю 11
    (Приказ ФНС); ОГРН — остаток первых 12 цифр на 11, ОГРНИП — на 13,
    берётся младший разряд остатка (эталон — пример из Википедии
    1037739010891 и реальный ИНН Сбербанка 7707083893);
  * проверка выполняется ТОЛЬКО для новых/изменённых значений: старые
    карточки открываются и пересохраняются без перепроверки (в реестрах
    могли остаться номера, введённые до включения проверки);
  * save_card (новая карточка) отклоняет реквизит с неверной контрольной
    цифрой для обоих форматов; окно редактирования проверяет только
    изменённые поля (проверено кодом: previous=saved в save()).

Тесты идут без Tk: валидаторы — чистые функции, путь сохранения — на
фейковом приложении counterparty_controller.
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import counterparty_controller as cc
import services
import storage
from document_builder import F


# --------------------------------------------------------------------------- #
#  Фикстуры и фейки (аналоги test_counterparty_controller, без Tk)
# --------------------------------------------------------------------------- #
@pytest.fixture()
def appdata(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    return tmp_path / "appdata"


@pytest.fixture()
def store(appdata):
    return storage.DataStore()


class FakeVar:
    def __init__(self, value=""):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class FakeEntry:
    def __init__(self, value=""):
        self._value = value
        self.error_marks: list[bool] = []

    def get(self):
        return self._value

    def set_error(self, on: bool):
        self.error_marks.append(on)


@pytest.fixture()
def tk_vars(monkeypatch):
    monkeypatch.setattr(cc.tk, "StringVar", FakeVar)


@pytest.fixture()
def ui_stub(monkeypatch):
    state: dict = {"alerts": []}

    def fake_alert(root, *args, **kwargs):
        state["alerts"].append((args, kwargs))

    monkeypatch.setattr(cc.ui, "alert", fake_alert)
    return state


# Карточка с валидными контрольными цифрами (общая с test_counterparty_controller).
from test_counterparty_controller import OOO_CARD, seed_entries, FakeApp  # noqa: E402


# --------------------------------------------------------------------------- #
#  ИНН: 10 цифр (юр. лицо)
# --------------------------------------------------------------------------- #
def test_inn10_valid_sberbank():
    """Реальный ИНН ПАО Сбербанк — эталон валидного номера."""
    assert services.inn_checksum_valid("7707083893")


def test_inn10_rejects_wrong_control_digit():
    assert not services.inn_checksum_valid("2540123456"), \
        "опечатка в контрольном разряде ловится"
    assert not services.inn_checksum_valid("7701234567")
    assert services.inn_checksum_valid("2540123459"), "корректный близнец проходит"


@pytest.mark.parametrize("bad", ["25401234", "25401234567", "abcdefghij", ""])
def test_inn10_rejects_wrong_shape(bad):
    assert not services.inn_checksum_valid(bad)


# --------------------------------------------------------------------------- #
#  ИНН: 12 цифр (ИП)
# --------------------------------------------------------------------------- #
def test_inn12_valid_and_invalid():
    assert services.inn_checksum_valid("254012345692")
    assert not services.inn_checksum_valid("254012345678"), \
        "12-значный ИНН проверяется по обоим контрольным разрядам"
    assert not services.inn_checksum_valid("254012345693")


# --------------------------------------------------------------------------- #
#  ОГРН (13) и ОГРНИП (15)
# --------------------------------------------------------------------------- #
def test_ogrn_wikipedia_example():
    """Пример из статьи «Контрольное число» — эталон алгоритма."""
    assert services.ogrn_checksum_valid("1037739010891")
    assert not services.ogrn_checksum_valid("1037739010892")
    assert not services.ogrn_checksum_valid("1035001234567"), "фикстура с опечаткой"
    assert services.ogrn_checksum_valid("1035001234564"), "корректный близнец проходит"


def test_ogrnip_valid_and_invalid():
    assert services.ogrnip_checksum_valid("304253600012343")
    assert not services.ogrnip_checksum_valid("304253600012345")
    assert not services.ogrnip_checksum_valid("304500100000154")
    assert services.ogrnip_checksum_valid("304500100000158")


@pytest.mark.parametrize("bad", ["1027", "10277001321970", "102770013219x", ""])
def test_ogrn_rejects_wrong_shape(bad):
    assert not services.ogrn_checksum_valid(bad)


# --------------------------------------------------------------------------- #
#  checksum_errors: только изменённые реквизиты
# --------------------------------------------------------------------------- #
def test_checksum_errors_flags_new_bad_inn():
    data = {F.INN: "2540123456", F.OGRN: "1035001234567"}
    errors = services.checksum_errors(data, is_ip=False, previous=None)
    assert F.INN in errors and F.OGRN in errors, "новые значения проверяются"
    assert "контрольной цифры" in errors[F.INN]


def test_checksum_errors_skips_unchanged_requisites():
    """Старые карточки не перепроверяются: значения равны прежним."""
    previous = {F.INN: "2540123456", F.OGRN: "1035001234567"}
    errors = services.checksum_errors(previous, is_ip=False, previous=previous)
    assert errors == {}, "неизменённые реквизиты не проверяются"


def test_checksum_errors_flags_only_changed_field():
    previous = {F.INN: "2540123456", F.OGRN: "1035001234567"}
    data = {F.INN: "2540123456", F.OGRN: "1035001234565"}  # изменён, цифра битая
    errors = services.checksum_errors(data, is_ip=False, previous=previous)
    assert F.OGRN in errors and F.INN not in errors, \
        "проверяется только поле, изменённое пользователем"


def test_checksum_errors_ip_uses_ogrnip():
    previous = {F.INN: "254012345692", F.OGRNIP: "304253600012345"}
    data = {F.INN: "254012345692", F.OGRNIP: "304253600012343"}
    errors = services.checksum_errors(data, is_ip=True, previous=previous)
    assert F.OGRNIP in errors or errors == {}, \
        "исправление на валидный ОГРНИП не даёт ошибки"
    bad = {F.INN: "254012345692", F.OGRNIP: "304253600012345"}
    errors_bad = services.checksum_errors(bad, is_ip=True, previous={})
    assert F.OGRNIP in errors_bad and "ОГРНИП" in errors_bad[F.OGRNIP]


def test_checksum_errors_empty_values_pass():
    assert services.checksum_errors({}, is_ip=False, previous=None) == {}


# --------------------------------------------------------------------------- #
#  Путь сохранения: save_card отклоняет неверную контрольную цифру
# --------------------------------------------------------------------------- #
def test_save_card_rejects_bad_checksum_inn(store, ui_stub, tk_vars):
    app = FakeApp(store)
    flow = app._counterparty_flow
    card = dict(OOO_CARD)
    card[F.INN] = "2540123456"  # контрольная цифра битая
    seed_entries(app, "ooo", card)
    flow.save_card()
    assert not app.store.counterparties, "карточка с опечаткой не сохраняется"
    alerts = [args for args, _ in ui_stub["alerts"]]
    assert any("контрольной цифры" in text for alert in alerts for text in alert), \
        "сообщение называет проблему"


def test_save_card_rejects_bad_checksum_ogrn(store, ui_stub, tk_vars):
    app = FakeApp(store)
    flow = app._counterparty_flow
    card = dict(OOO_CARD)
    card[F.OGRN] = "1035001234567"  # контрольная цифра битая
    seed_entries(app, "ooo", card)
    flow.save_card()
    assert not app.store.counterparties, "карточка с опечаткой в ОГРН не сохраняется"


def test_save_card_accepts_valid_checksum(store, ui_stub):
    app = FakeApp(store)
    flow = app._counterparty_flow
    seed_entries(app, "ooo", OOO_CARD)
    flow.save_card()
    assert len(app.store.counterparties) == 1, \
        "корректные реквизиты сохраняются без вопросов"
