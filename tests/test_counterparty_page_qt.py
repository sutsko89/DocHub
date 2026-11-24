r"""Тесты M4 — страница «Добавить контрагента» и окно редактирования на Qt.

Проверяют критерии этапа M4 (план миграции) на offscreen-платформе
(fixture ``qapp`` pytest-qt, APPDATA изолируется ``appdata``; модальные
диалоги подменяются на уровне qt_dialogs):

* страница заменяет плейсхолдер; состав полей ООО/ИП — Tk-паритет
  («не трогаем»: состав полей F, правила обязательности); у ИП
  фактический адрес повторяет юридический;
* чтение полей: сбой не превращается в пустую строку (D4) —
  обязательное блокирует сохранение с явным сообщением, необязательное —
  явный выбор «Сохранить с прежними значениями» (DEF-1);
* валидация через services.validate_party_fields + checksum_errors
  (CHKSUM: новая карточка — все реквизиты новые; окно правки — только
  изменённые, previous=saved);
* дубликат ИНН — выбор из трёх исходов (п. 4.1); при правке — ошибка поля;
* сохранение: файл-«фантом» при сбое записи не остаётся (п. 9.2);
* «Заполнить по ИНН» (U8): guard busy, проверка ИНН по режиму, без ключа
  — понятная ошибка, подстановка в поля из очереди в главном потоке,
  тихий отказ при закрытом окне, статус запроса;
* окно правки: prefill, закрытие с изменениями — подтверждение (DEF-1),
  без изменений — сразу, Escape — тот же guard;
* защита страницы п. 4.4: dirty при правке, Ctrl+S, подтверждение при
  уходе, сбой подписи = «грязная»;
* кросс-совместимость: база после Qt-добавления читается свежим
  DataStore, schema_version прежняя.

W-1/W-2 и Tk-часть тестов этот файл не трогает.
"""

from __future__ import annotations

import logging
import os
import sqlite3
import sys
from pathlib import Path

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import counterparty_page_qt
import dadata_client
import dadata_qt
import main_qt
import qt_dialogs
import services
import storage
from document_builder import F
from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtGui import QValidator
from PySide6.QtWidgets import QDialog, QLineEdit

REPO_ROOT = Path(__file__).resolve().parent.parent

# CHKSUM (этап 2): реквизиты с валидными контрольными цифрами (те же
# значения, что в test_counterparty_controller — эталон Tk-поведения).
OOO_CARD = {
    F.INN: "2540123459",
    F.KPP: "254001001",
    F.OGRN: "1035001234564",
    F.FULL_NAME: "Общество с ограниченной ответственностью «Ромашка»",
    F.SHORT_NAME: "ООО «Ромашка»",
    F.DIRECTOR: "Петров Пётр Петрович",
    F.BASIS: "Устав",
    F.LEGAL_ADDRESS: "690000, г. Владивосток, ул. Светланская, д. 1",
    F.FACT_ADDRESS: "690000, г. Владивосток, ул. Светланская, д. 1",
    F.EMAIL: "buy@romashka.ru",
    F.PHONE: "+7 (423) 123-45-67",
    F.ACCOUNT: "40702810123456789012",
    F.CORR_ACCOUNT: "30101810100000000225",
    F.BANK: "ПАО Приморский банк",
    F.BIK: "044525225",
}
IP_CARD = {
    F.INN: "254012345692",
    F.OGRNIP: "304253600012343",
    F.IP_NAME: "Иванов Иван Иванович",
    F.LEGAL_ADDRESS: "690000, г. Владивосток, ул. Светланская, д. 2",
    F.EMAIL: "ip@romashka.ru",
    F.PHONE: "+7 (423) 123-45-68",
    F.ACCOUNT: "40702810123456789012",
    F.CORR_ACCOUNT: "30101810100000000225",
    F.BANK: "ПАО Приморский банк",
    F.BIK: "044525225",
}

DADATA_OOO_INFO = {
    "full_name": "ООО «Ландыш» (DaData)",
    "short_name": "ООО «Ландыш»",
    "director": "Сидоров Сидор Сидорович",
    "legal_address": "г. Москва, ул. Тверская, д. 1",
    "inn": "7701234567",
    "kpp": "770101001",
    "ogrn": "1037739010221",
    "bank_name": "ПАО Сбербанк",
    "bik": "044525225",
    "correspondent_account": "30101810100000000225",
}


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_dialogs(monkeypatch):
    """Подменяет модальные диалоги: подтверждения True, выбор и алерты пишутся."""
    captured = {
        "confirms": [],
        "alerts": [],
        "choice": {"value": "cancel", "calls": []},
    }

    def fake_confirm(*args, **kwargs):
        captured["confirms"].append(args)
        return True

    def fake_choice(*args, **kwargs):
        captured["choice"]["calls"].append(args)
        return captured["choice"]["value"]

    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(
        qt_dialogs, "alert", lambda *a, **k: captured["alerts"].append(a)
    )
    monkeypatch.setattr(qt_dialogs, "choice", fake_choice)
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs):
    """Каркас с живыми страницами CP1/M4; таймеры и окно убираются после.

    Окно закрывается и удаляется с явным flush DeferredDelete: обычный
    processEvents не сбрасывает отложенные удаления Qt — без flush каждый
    каркас оставлял живые виджеты, и время setup росло линейно по файлу
    (0,12 → 4,5 с; с flush — плоско, ~0,03 с). Те же две строки в
    докстринге фикстуры M3-тестов не меняются (не входят в этап M4).
    """
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl._stop_timers()
    if ctrl.spec_flow.preview_timer is not None:
        ctrl.spec_flow.preview_timer.stop()
    ctrl.window.close()
    ctrl.window.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    qapp.processEvents()


def fill_card(flow, mode: str, data: dict) -> None:
    """Заполняет поля карточки значениями словаря (программная правка)."""
    for key, value in data.items():
        adapter = flow.entries[mode].get(key)
        if adapter is not None:
            adapter.set_text(str(value))


def save_stored_party(controller, mode: str = "ooo", fields: dict | None = None):
    """Добавляет контрагента напрямую в store (без UI) — эталонный rec."""
    data = dict(OOO_CARD if mode == "ooo" else IP_CARD)
    if fields:
        data.update(fields)
    if mode == "ip":
        data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")
    return controller.store.add_counterparty(mode, data)


# --------------------------------------------------------------------------- #
#  Страница: установка, состав полей, тип карточки
# --------------------------------------------------------------------------- #
def test_counterparty_page_replaces_placeholder(controller):
    flow = controller.counterparty_flow
    assert controller.pages["counterparty"] is flow.page
    assert controller.stack.indexOf(flow.page) >= 0
    assert flow.status_label is not None


def test_page_fields_composition_tk_parity(controller):
    """Состав полей обеих карточек — Tk-паритет («не трогаем» план M4)."""
    flow = controller.counterparty_flow
    ooo = set(flow.entries["ooo"])
    ip = set(flow.entries["ip"])
    # ООО: КПП/ОГРН/руководитель/основание; ИП: ОГРНИП/ФИО.
    assert {F.KPP, F.OGRN, F.FULL_NAME, F.SHORT_NAME, F.DIRECTOR, F.BASIS} <= ooo
    assert {F.OGRNIP, F.IP_NAME} <= ip
    assert F.OGRNIP not in ooo and F.FULL_NAME not in ip
    # У ИП фактический адрес повторяет юридический — отдельного поля нет.
    assert F.FACT_ADDRESS not in ip
    assert F.FACT_ADDRESS in ooo
    # Общие банковские поля в обоих режимах.
    assert {F.INN, F.LEGAL_ADDRESS, F.ACCOUNT, F.CORR_ACCOUNT, F.BANK, F.BIK} <= (ooo & ip)


def test_type_switch_raises_stack(controller):
    flow = controller.counterparty_flow
    assert flow.current_mode() == "ooo"
    flow.on_type_change("ИП")
    assert flow.current_mode() == "ip"
    assert flow.stack.currentIndex() == 1
    flow.on_type_change("Юр. лицо")
    assert flow.current_mode() == "ooo"
    assert flow.stack.currentIndex() == 0


def test_format_hints_are_placeholders(controller):
    """Подсказки формата — подложкой пустого поля (U4), не в значении."""
    flow = controller.counterparty_flow
    inn_edit = flow.entries["ooo"][F.INN].edit
    assert inn_edit.placeholderText() == "10 или 12 цифр"
    assert inn_edit.text() == ""


def test_digits_validators_reject_letters(controller):
    flow = controller.counterparty_flow
    inn_edit = flow.entries["ooo"][F.INN].edit
    assert inn_edit.maxLength() == 12
    state, _text, _pos = inn_edit.validator().validate("abc", 0)
    assert state == QValidator.State.Invalid
    account_edit = flow.entries["ooo"][F.ACCOUNT].edit
    state, _text, _pos = account_edit.validator().validate("407028101234567890123", 0)
    assert state == QValidator.State.Invalid  # длиннее 20


def test_digits_paste_sanitized(controller, monkeypatch):
    """Вставка «р/с 40702810…» кладёт только цифры (аудит 2.0.1, Tk-паритет)."""
    flow = controller.counterparty_flow
    edit = flow.entries["ooo"][F.ACCOUNT].edit
    from PySide6.QtCore import QMimeData

    mime = QMimeData()
    mime.setText("р/с 40702810123456789012")
    edit.insertFromMimeData(mime)
    assert edit.text() == "40702810123456789012"


def test_phone_mask_applies_on_input(controller):
    flow = controller.counterparty_flow
    edit = flow.entries["ooo"][F.PHONE].edit
    edit.setText("9123456789")
    assert edit.text() == "+7 (912) 345-67-89"
    edit.setText("")
    assert edit.text() == ""
    edit.setText("89123456789")
    assert edit.text().startswith("+7 (912)")


def test_phone_paste_sanitizer():
    # Пробелы допустимы (Tk-паритет: ch.isdigit() or ch in "+-()/. ").
    assert counterparty_page_qt.sanitize_phone_paste("тел: +7 (912) 345") == " +7 (912) 345"
    assert counterparty_page_qt.sanitize_phone_paste("abc") == ""


# --------------------------------------------------------------------------- #
#  Сбор данных и сохранение карточки
# --------------------------------------------------------------------------- #
def test_collect_strips_values(controller):
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", {F.INN: "  2540123459  "})
    data = flow.collect("ooo")
    assert data[F.INN] == "2540123459"


def test_collect_ip_copies_legal_to_fact(controller):
    flow = controller.counterparty_flow
    fill_card(flow, "ip", {F.LEGAL_ADDRESS: "г. Владивосток, ул. Алеутская, 3"})
    data = flow.collect("ip")
    assert data[F.FACT_ADDRESS] == "г. Владивосток, ул. Алеутская, 3"


def test_collect_read_failure_returns_none(controller):
    """Сбой чтения поля — None, а не пустая строка (D4); тип — в журнал."""
    flow = controller.counterparty_flow

    class Broken:
        def get(self):
            raise RuntimeError("C++ object deleted")

    logged: list = []
    original_log = flow.log
    monkey = pytest.MonkeyPatch()
    monkey.setattr(
        type(original_log), "warning",
        lambda self, msg, *a: logged.append(str(msg) % a if a else str(msg)),
    )
    original = flow.entries["ooo"][F.EMAIL]
    flow.entries["ooo"][F.EMAIL] = Broken()
    try:
        data = flow.collect("ooo")
    finally:
        flow.entries["ooo"][F.EMAIL] = original
        monkey.undo()
    assert data is None
    assert any("поле карточки" in message for message in logged)


def test_save_card_empty_shows_validation_errors(controller, stub_dialogs):
    flow = controller.counterparty_flow
    flow.save_card()
    assert any(
        "Ошибки в данных" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )
    # Подсветка: у обязательных полей — «Заполните поле».
    assert flow.error_labels["ooo"][F.INN].text() == "Заполните поле"
    assert controller.store.active_counterparties() == []


def test_save_card_checksum_error_blocks_store(controller, stub_dialogs):
    """CHKSUM: опечатка в контрольной цифре ловится при вводе."""
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", {**OOO_CARD, F.INN: "2540123456"})  # контрольная битая
    flow.save_card()
    assert any(
        "Ошибки в данных" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )
    assert "контрольной цифры" in flow.error_labels["ooo"][F.INN].text()
    assert controller.store.active_counterparties() == []


def test_save_card_success_sets_current_party(controller, stub_dialogs):
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", OOO_CARD)
    flow.save_card()
    parties = controller.store.active_counterparties()
    assert len(parties) == 1
    assert parties[0]["type"] == "ooo"
    assert controller.current_party_id == parties[0]["id"]
    assert "добавлен в реестр" in controller.toast_label.text()
    # Tk-паритет: после сохранения форма «чистая».
    assert flow.is_form_dirty() is False
    # Уход к использованию контрагента (в Tk — «Договор», M6; ближайшая
    # живая страница — «Спецификация»; отклонение зафиксировано в заметках).
    assert controller.current_page == "spec"


def test_save_card_duplicate_inn_cancel(controller, stub_dialogs):
    save_stored_party(controller)
    controller.select_page("counterparty")
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", OOO_CARD)
    stub_dialogs["choice"]["value"] = "cancel"
    flow.save_card()
    assert len(controller.store.active_counterparties()) == 1
    assert controller.current_page == "counterparty"


def test_save_card_duplicate_inn_create_anyway(controller, stub_dialogs):
    save_stored_party(controller)
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", OOO_CARD)
    stub_dialogs["choice"]["value"] = "create"
    flow.save_card()
    assert len(controller.store.active_counterparties()) == 2


def test_save_card_duplicate_inn_open_existing(controller, stub_dialogs):
    existing = save_stored_party(controller)
    controller.select_page("counterparty")
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", OOO_CARD)
    stub_dialogs["choice"]["value"] = "open"
    flow.save_card()
    assert len(controller.store.active_counterparties()) == 1
    assert controller.current_party_id == existing["id"]
    assert controller.current_page == "spec"


def test_save_card_store_error_removes_phantom(controller, stub_dialogs):
    """Сбой записи реестра: алерт честный, «фантома» в памяти нет (п. 9.2)."""
    flow = controller.counterparty_flow
    fill_card(flow, "ooo", OOO_CARD)
    store = controller.store
    original_save = storage.DataStore.save

    def failing_save(self):
        self.last_error = "Не удалось записать базу данных на диск"
        return False

    monkey = pytest.MonkeyPatch()
    monkey.setattr(storage.DataStore, "save", failing_save)
    try:
        flow.save_card()
    finally:
        monkey.undo()
    assert any(
        "Не удалось сохранить контрагента" == str(alert[1])
        for alert in stub_dialogs["alerts"]
    )
    assert len(store.counterparties) == 0


def test_save_card_collect_failure_alerts(controller, stub_dialogs):
    flow = controller.counterparty_flow

    class Broken:
        def get(self):
            raise RuntimeError("C++ object deleted")

    original = flow.entries["ooo"][F.INN]
    flow.entries["ooo"][F.INN] = Broken()
    try:
        flow.save_card()
    finally:
        flow.entries["ooo"][F.INN] = original
    assert any(
        "Не удалось прочитать поля формы" == str(alert[1])
        for alert in stub_dialogs["alerts"]
    )
    assert controller.store.active_counterparties() == []


# --------------------------------------------------------------------------- #
#  Защита страницы п. 4.4: dirty, Ctrl+S, уход
# --------------------------------------------------------------------------- #
def test_page_dirty_on_edit(controller):
    flow = controller.counterparty_flow
    assert flow.is_form_dirty() is False
    fill_card(flow, "ooo", {F.INN: "2540123459"})
    assert flow.is_form_dirty() is True
    flow.reset_form_baseline()
    assert flow.is_form_dirty() is False


def test_page_dirty_on_type_switch(controller):
    """Переключение типа меняет подпись активного режима — как в Tk."""
    flow = controller.counterparty_flow
    flow.on_type_change("ИП")
    assert flow.is_form_dirty() is True
    flow.reset_form_baseline()
    assert flow.is_form_dirty() is False


def test_page_dirty_read_failure_is_dirty(controller):
    """Сбой подписи — None — не совпадает ни с каким baseline («грязная»)."""
    flow = controller.counterparty_flow
    original_collect = flow.collect
    flow.collect = lambda mode: None  # noqa: B010 — подмена прогона
    try:
        assert flow.is_form_dirty() is True
    finally:
        flow.collect = original_collect


def test_ctrl_s_saves_card(controller, stub_dialogs):
    controller.select_page("counterparty")
    fill_card(controller.counterparty_flow, "ooo", OOO_CARD)
    controller.save_current_page()
    assert len(controller.store.active_counterparties()) == 1


def test_leave_dirty_page_asks_confirm(controller, stub_dialogs):
    controller.select_page("counterparty")
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "2540123459"})
    controller.select_page("spec")
    assert any(
        "Несохранённые изменения" == str(args[1])
        for args in stub_dialogs["confirms"]
    )
    # Подтверждение (stub True) — раздел сменился.
    assert controller.current_page == "spec"


def test_leave_dirty_page_cancel_keeps_page(controller, stub_dialogs, monkeypatch):
    controller.select_page("counterparty")
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "2540123459"})
    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: False)
    controller.select_page("spec")
    assert controller.current_page == "counterparty"


def test_leave_clean_page_no_confirm(controller, stub_dialogs):
    controller.select_page("counterparty")
    controller.select_page("spec")
    assert stub_dialogs["confirms"] == []


# --------------------------------------------------------------------------- #
#  Окно редактирования (DEF-1, Escape, сохранение)
# --------------------------------------------------------------------------- #
def test_edit_dialog_prefills_fields(controller):
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    assert dialog.windowTitle().startswith("Редактирование контрагента")
    assert dialog.entries[F.SHORT_NAME].get() == OOO_CARD[F.SHORT_NAME]
    assert dialog.entries[F.BASIS].get() == "Устав"
    assert dialog.has_changes() is False
    dialog._really_close()


def test_edit_dialog_clean_close_no_confirm(controller, stub_dialogs):
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.request_close()
    assert not dialog.isVisible()
    assert stub_dialogs["confirms"] == []


def test_edit_dialog_dirty_close_asks_confirm(controller, stub_dialogs):
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.show()
    dialog.entries[F.DIRECTOR].set_text("Другой Директор")
    assert dialog.has_changes() is True
    dialog.request_close()
    assert any(
        "Несохранённые изменения" == str(args[1]) for args in stub_dialogs["confirms"]
    )
    # stub confirm=True («Закрыть без сохранения») — окно закрылось,
    # карточка не изменилась.
    assert not dialog.isVisible()
    assert controller.store.get_counterparty(rec["id"])["fields"][F.DIRECTOR] \
        == OOO_CARD[F.DIRECTOR]


def test_edit_dialog_escape_triggers_same_guard(controller, stub_dialogs, monkeypatch):
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.show()
    dialog.entries[F.DIRECTOR].set_text("Другой Директор")

    def recording_false(*args, **kwargs):
        stub_dialogs["confirms"].append(args)
        return False

    monkeypatch.setattr(qt_dialogs, "confirm", recording_false)
    dialog.reject()  # Escape/крестик → guard → отказ → окно осталось
    assert dialog.isVisible()
    assert any(
        "Несохранённые изменения" == str(args[1]) for args in stub_dialogs["confirms"]
    )


def test_edit_dialog_save_updates_card(controller, stub_dialogs):
    rec = save_stored_party(controller)
    controller.set_current_party(rec)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.entries[F.DIRECTOR].set_text("Смирнов Смирн Смирнович")
    dialog.save()
    updated = controller.store.get_counterparty(rec["id"])
    assert updated["fields"][F.DIRECTOR] == "Смирнов Смирн Смирнович"
    assert "обновлена" in controller.toast_label.text()
    assert "Ранее созданные документы не изменялись" in controller.toast_label.text()
    # Документы не трогаются (их здесь нет), карточка — единственное изменение.
    assert not dialog.isVisible()


def test_edit_dialog_def1_blocking_read_failure(controller, stub_dialogs):
    """DEF-1: сбой чтения ОБЯЗАТЕЛЬНОГО поля блокирует сохранение."""
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)

    class Broken:
        def get(self):
            raise RuntimeError("C++ object deleted")

    original = dialog.entries[F.INN]
    dialog.entries[F.INN] = Broken()
    try:
        dialog.save()
    finally:
        dialog.entries[F.INN] = original
    assert any(
        "Не удалось прочитать поля карточки" == str(alert[1])
        for alert in stub_dialogs["alerts"]
    )
    assert controller.store.get_counterparty(rec["id"])["fields"] == dict(OOO_CARD)


def test_edit_dialog_def1_optional_failure_keeps_previous(controller, stub_dialogs):
    """DEF-1: сбой необязательного поля — явный выбор «с прежними значениями»."""
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.entries[F.DIRECTOR].set_text("Смирнов Смирн Смирнович")

    class Broken:
        def get(self):
            raise RuntimeError("C++ object deleted")

    original = dialog.entries[F.EMAIL]
    dialog.entries[F.EMAIL] = Broken()
    try:
        dialog.save()
    finally:
        dialog.entries[F.EMAIL] = original
    # Явный выбор был (confirm → True), необязательный email сохранён прежним.
    updated = controller.store.get_counterparty(rec["id"])
    assert updated["fields"][F.EMAIL] == OOO_CARD[F.EMAIL]
    assert updated["fields"][F.DIRECTOR] == "Смирнов Смирн Смирнович"


def test_edit_dialog_checksum_skips_unchanged_requisites(controller, stub_dialogs):
    """CHKSUM при правке: прежние реквизиты не перепроверяются (previous=saved)."""
    rec = save_stored_party(
        controller, fields={F.INN: "2540123456", F.OGRN: "1035001234567"}
    )
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.entries[F.DIRECTOR].set_text("Смирнов Смирн Смирнович")
    dialog.save()
    updated = controller.store.get_counterparty(rec["id"])
    assert updated["fields"][F.INN] == "2540123456"  # прежний ИНН не тронут
    assert updated["fields"][F.DIRECTOR] == "Смирнов Смирн Смирнович"


def test_edit_dialog_duplicate_inn_errors(controller, stub_dialogs):
    first = save_stored_party(controller)
    second = save_stored_party(
        controller,
        fields={F.INN: "7701234560", F.SHORT_NAME: "ООО «Ландыш»"},
    )
    assert second["fields"][F.INN] == "7701234560"
    assert services.inn_checksum_valid("7701234560")
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, first)
    dialog.show()
    dialog.entries[F.INN].set_text("7701234560")  # ИНН второй карточки
    dialog.save()
    assert "уже есть в реестре" in dialog.error_labels[F.INN].text()
    assert controller.store.get_counterparty(first["id"])["fields"][F.INN] \
        == OOO_CARD[F.INN]


def test_edit_dialog_store_error_alerts(controller, stub_dialogs):
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.show()
    dialog.entries[F.DIRECTOR].set_text("Смирнов Смирн Смирнович")

    def failing_save(self):
        self.last_error = "Не удалось записать базу данных на диск"
        return False

    monkey = pytest.MonkeyPatch()
    monkey.setattr(storage.DataStore, "save", failing_save)
    try:
        dialog.save()
    finally:
        monkey.undo()
    assert any(
        "Не удалось сохранить изменения" == str(alert[1])
        for alert in stub_dialogs["alerts"]
    )
    assert dialog.isVisible()


def test_edit_dialog_from_spec_bar_button(controller, stub_dialogs):
    """Точка входа M4: кнопка на панели «Документ для» открывает окно правки."""
    rec = save_stored_party(controller)
    controller.set_current_party(rec)
    controller.spec_flow.edit_current_party()
    dialogs = controller.window.findChildren(QDialog)
    assert dialogs, "окно правки открыто"
    dialogs[0]._really_close()
    # Без выбранного контрагента — честный алерт, окна нет.
    controller.current_party_id = None
    controller.spec_flow._sync_current_party()
    controller.spec_flow.edit_current_party()
    assert any(
        "Контрагент не выбран" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )


# --------------------------------------------------------------------------- #
#  «Заполнить по ИНН» (U8): guard, валидация, подстановка, тихий отказ
# --------------------------------------------------------------------------- #
class RecordedThread:
    def __init__(self, target, args, daemon):
        self.target = target
        self.args = args
        self.daemon = daemon
        self.started = False

    def start(self):
        self.started = True


class RecordingFactory:
    def __init__(self):
        self.created: list = []

    def __call__(self, target=None, args=(), daemon=None, **_kw):
        thread = RecordedThread(target, args, daemon)
        self.created.append(thread)
        return thread


def test_dadata_guard_busy(controller, stub_dialogs):
    """Guard повторных кликов: при активном запросе новый воркер не стартует."""
    flow = controller.dadata_flow
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "7701234567"})
    flow.busy = True  # активный запрос
    factory = RecordingFactory()
    flow.fill_from_inn(thread_factory=factory)
    assert factory.created == []


def test_dadata_inn_length_alerts_by_mode(controller, stub_dialogs):
    flow = controller.dadata_flow
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "123"})
    flow.fill_from_inn()
    assert any(
        "Проверьте ИНН" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )
    assert flow.busy is False


def test_dadata_ip_mode_expects_12_digits(controller, stub_dialogs):
    flow = controller.dadata_flow
    controller.counterparty_flow.on_type_change("ИП")
    fill_card(controller.counterparty_flow, "ip", {F.INN: "7701234567"})  # 10, а нужно 12
    flow.fill_from_inn()
    assert any(
        "ИНН должен содержать 12 цифр" in str(alert[2])
        for alert in stub_dialogs["alerts"]
    )


def test_dadata_no_key_alerts(controller, stub_dialogs, monkeypatch):
    flow = controller.dadata_flow
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "7701234567"})
    monkeypatch.setattr(dadata_client, "get_api_key", lambda: "")
    flow.fill_from_inn()
    assert any(
        "Нет API-ключа" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )
    assert flow.busy is False


def test_dadata_worker_starts_and_status_busy(controller, monkeypatch):
    flow = controller.dadata_flow
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "7701234567"})
    monkeypatch.setattr(dadata_client, "get_api_key", lambda: "test-key")
    factory = RecordingFactory()
    flow.fill_from_inn(thread_factory=factory)
    assert len(factory.created) == 1
    assert factory.created[0].started is True
    args = factory.created[0].args
    assert args[0] == "7701234567" and args[1] == "test-key" and args[2] == "ooo"
    assert "Загружаю реквизиты" in controller.counterparty_flow.status_label.text()
    # Poll-таймеры опроса зарегистрированы (контракт раздела Д: QTimer, не QThread).
    assert "dadata_result" in controller._timers
    assert "dadata_busy" in controller._timers
    flow.abort()


def test_dadata_result_applied_to_fields(controller):
    """Результат из очереди подставляется в главном потоке (result_tick)."""
    flow = controller.dadata_flow
    fill_card(controller.counterparty_flow, "ooo", {F.INN: "7701234567"})
    flow.busy = True
    queue = __import__("queue").Queue()
    flow.result_queue = queue
    flow._target = controller.counterparty_flow.dadata_target("ooo")
    queue.put(("ok", "ooo", DADATA_OOO_INFO))
    flow.result_tick()
    entries = controller.counterparty_flow.entries["ooo"]
    assert entries[F.SHORT_NAME].get() == "ООО «Ландыш»"
    assert entries[F.LEGAL_ADDRESS].get() == "г. Москва, ул. Тверская, д. 1"
    assert entries[F.BIK].get() == "044525225"
    assert entries[F.FACT_ADDRESS].get() == "г. Москва, ул. Тверская, д. 1"  # был пуст
    assert "Реквизиты заполнены" in controller.counterparty_flow.status_label.text()
    assert flow.busy is False


def test_dadata_result_error_alerts(controller, stub_dialogs):
    flow = controller.dadata_flow
    flow.busy = True
    queue = __import__("queue").Queue()
    flow.result_queue = queue
    flow._target = controller.counterparty_flow.dadata_target("ooo")
    queue.put(("error", "Организация с таким ИНН не найдена."))
    flow.result_tick()
    assert any(
        "Не удалось заполнить реквизиты" == str(alert[1])
        for alert in stub_dialogs["alerts"]
    )
    assert flow.busy is False


def test_dadata_party_loaded_no_data_alerts(controller, stub_dialogs):
    flow = controller.dadata_flow
    flow._target = controller.counterparty_flow.dadata_target("ooo")
    flow.party_loaded("ooo", {key: "" for key in DADATA_OOO_INFO})
    assert any(
        "Данных нет" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )


def test_dadata_quiet_failure_when_edit_dialog_closed(controller, stub_dialogs, qapp):
    """U8: окно правки закрылось, пока запрос летел — тихий отказ, без алерта."""
    rec = save_stored_party(controller)
    dialog = counterparty_page_qt.EditCounterpartyDialog(controller, rec)
    dialog.show()
    target = dialog.dadata_target()
    dialog.deleteLater()
    dialog._closing = True
    dialog.reject()
    qapp.processEvents()
    controller.dadata_flow._target = target
    controller.dadata_flow.party_loaded("ooo", DADATA_OOO_INFO)
    assert stub_dialogs["alerts"] == []
    assert controller.dadata_flow.busy is False


# --------------------------------------------------------------------------- #
#  Кросс-совместимость базы (M3+)
# --------------------------------------------------------------------------- #
def test_cross_base_after_qt_add(controller):
    fill_card(controller.counterparty_flow, "ooo", OOO_CARD)
    controller.counterparty_flow.save_card()
    fresh = storage.DataStore()
    parties = fresh.active_counterparties()
    assert len(parties) == 1
    assert storage.display_name(parties[0]) == "ООО «Ромашка»"
    with sqlite3.connect(controller.store.db_path() if hasattr(
        controller.store, "db_path"
    ) else storage.db_path()) as conn:
        meta = dict(conn.execute("SELECT key, value FROM meta"))
    assert int(meta.get("schema_version", "0")) == storage.SQLITE_SCHEMA_VERSION
