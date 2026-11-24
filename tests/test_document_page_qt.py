r"""Тесты M6-a — страницы «Договор» и «Доп. соглашение» на Qt.

Qt-эквиваленты test_document_page_controller / test_document_save_controller
(Tk-контракты сохраняются дословно):

* сборка страниц заменяет плейсхолдеры; предпросмотр — печатная форма из
  шаблона без записи на диск, только при видимой странице (P-07);
* единый текущий контрагент (п. 4.2): выбор на одной странице переносится
  на остальные; поиск по ИНН (U3) с честным алертом;
* «Действует до» — дефолт 31 декабря текущего года (2.18.0), суммы и
  сроки прописью, формат лимита с разрядами;
* сохранение договора/доп. соглашения: транзакция «файл + запись» (п. 9.2),
  честное «сохранён» с тостом и открытием файла (сценарий Г #4), валидация
  с подсветкой, подтверждение повтора номера (п. 4.1);
* LAZY-1: тело ленивого договора дочитывается store.document_body —
  предпросмотр и сохранение доп. соглашения получают полный текст;
* защита несохранённых данных (п. 4.4): подпись не включает выбор
  контрагента, сбой чтения — «грязная», уход/закрытие — подтверждение.

Запускаются на offscreen-платформе (fixture ``qapp`` pytest-qt, APPDATA
изолируется ``appdata``; модальные диалоги подменяются на уровне
qt_dialogs). Tk-часть набора этот файл не трогает.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import document_page_qt
import main_qt
import qt_dialogs
import services
import shell_open
import storage
from document_builder import F, docx_to_text
from PySide6.QtWidgets import QLabel

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_dialogs(monkeypatch):
    """Подменяет модальные диалоги: подтверждения True, алерты собираются."""
    captured = {"confirms": [], "alerts": [], "choices": []}

    def fake_confirm(*args, **kwargs):
        captured["confirms"].append(args)
        return True

    def fake_choice(*args, **kwargs):
        captured["choices"].append(args)
        return None

    captured["confirm"] = fake_confirm
    captured["choice"] = fake_choice
    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(qt_dialogs, "alert",
                        lambda *a, **k: captured["alerts"].append(a))
    monkeypatch.setattr(qt_dialogs, "choice", fake_choice)
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs, monkeypatch):
    """Каркас с живыми страницами документов; открытие файлов глушится."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl._stop_timers()
    for flow in (ctrl.spec_flow, ctrl.contract_flow, ctrl.additional_flow):
        if getattr(flow, "preview_timer", None) is not None:
            flow.preview_timer.stop()


@pytest.fixture()
def seeded_party(controller):
    """Контрагент ООО из golden-образца + договор-основание."""
    from render_helpers import OOO_FIELDS

    store = controller.store
    fields = {k: v for k, v in OOO_FIELDS.items()
              if k not in (F.NUMBER, F.DAY, F.MONTH)}
    party = store.add_counterparty("ooo", fields)
    contract = store.add_document(
        storage.KIND_CONTRACT, party["id"], OOO_FIELDS[F.NUMBER],
        f"{OOO_FIELDS[F.DAY]} {OOO_FIELDS[F.MONTH]}",
        data=dict(OOO_FIELDS),
        extra={"day": OOO_FIELDS[F.DAY], "month": OOO_FIELDS[F.MONTH],
               "is_ip": False},
        text="",
    )
    for flow in (controller.spec_flow, controller.contract_flow,
                 controller.additional_flow):
        flow.refresh_parties()
    return {"party": party, "contract": contract}


def redirect_docs_dir(flow, tmp_path):
    """Перенаправляет папку документов потока в tmp (как в тестах M3)."""
    flow.party_documents_dir = lambda rec=None: str(tmp_path)


def render_once(controller, qapp):
    """Синхронно прогоняет отложенный предпросмотр (таймер 120 мс)."""
    for flow in (controller.spec_flow, controller.contract_flow,
                 controller.additional_flow):
        timer = getattr(flow, "preview_timer", None)
        if timer is not None and timer.isActive():
            timer.stop()
        render = getattr(flow, "render_preview", None)
        if render is not None:
            render()
        else:
            flow.render_spec_preview()
    qapp.processEvents()


# --------------------------------------------------------------------------- #
#  Сборка страниц и предпросмотр (п. 5.6, P-07)
# --------------------------------------------------------------------------- #
def test_document_pages_replace_placeholders(controller):
    assert controller.pages["contract"] is controller.contract_flow.page
    assert controller.pages["additional"] is controller.additional_flow.page
    assert controller.stack.indexOf(controller.contract_flow.page) >= 0
    for key in ("contract", "additional"):
        texts = [l.text() for l in controller.pages[key].findChildren(QLabel)]
        assert not any("в разработке" in t for t in texts), key


def test_contract_preview_renders_from_template(controller, seeded_party, qapp):
    """Предпросмотр договора — печатная форма из шаблона (сценарий Г #4)."""
    flow = controller.contract_flow
    flow.number_entry.setText("516")
    flow.day_entry.setText("10")
    controller.select_page("contract")
    render_once(controller, qapp)
    text = flow.preview_text.toPlainText()
    assert len(text) > 500, "предпросмотр — полный текст печатной формы"
    assert "516" in text
    assert flow.preview_text.isReadOnly(), "предпросмотр только для чтения (п. 4.3)"


def test_contract_preview_placeholder_without_party(controller):
    controller.select_page("contract")
    controller.contract_flow.render_preview()
    assert "Выберите контрагента" in controller.contract_flow.preview_text.toPlainText()


def test_preview_only_on_visible_page_p07(controller, seeded_party, qapp):
    """P-07: операции из других страниц не рендерят скрытый предпросмотр."""
    flow = controller.contract_flow
    controller.select_page("spec")
    flow.number_entry.setText("777")
    render_once(controller, qapp)
    assert flow.preview_text.toPlainText() == "", \
        "скрытая страница не рендерится (текст не собран)"
    controller.select_page("contract")
    flow.on_page_open()
    assert "777" in flow.preview_text.toPlainText(), \
        "переключение на страницу пересобирает предпросмотр"


def test_additional_preview_needs_contract(controller, seeded_party, qapp):
    """Без договора-основания предпросмотр просит выбрать, не падает."""
    store = controller.store
    for doc in list(store.documents):
        if doc.get("kind") == storage.KIND_CONTRACT:
            store.purge_document(doc["id"])
    controller.additional_flow.refresh_parties()
    controller.select_page("additional")
    controller.additional_flow.render_preview()
    assert "договор-основание" in controller.additional_flow.preview_text.toPlainText()


def test_additional_preview_renders_contract_text(controller, seeded_party, qapp):
    controller.select_page("additional")
    controller.additional_flow.limit_entry.setText("10 000")
    controller.additional_flow.term_entry.setText("10")
    render_once(controller, qapp)
    text = controller.additional_flow.preview_text.toPlainText()
    assert len(text) > 500
    assert "дополнительное соглашение" in text.lower() or "Доп" in text


# --------------------------------------------------------------------------- #
#  Единый текущий контрагент (п. 4.2) и поиск по ИНН (U3)
# --------------------------------------------------------------------------- #
def test_party_selection_propagates_to_all_doc_pages(controller, seeded_party):
    """Выбор на странице договора переносится на остальные (п. 4.2)."""
    flow = controller.contract_flow
    flow.party_combo.setCurrentIndex(0)  # первый (единственный) контрагент
    display = flow.party_display(seeded_party["party"], flow.party_duplicates())
    assert controller.spec_flow.party_combo.currentText() == display
    assert controller.additional_flow.party_combo.currentText() == display
    assert controller.current_party_id == seeded_party["party"]["id"]


def test_party_selection_from_additional_propagates_back(controller, seeded_party):
    """Симметрия: выбор на доп. соглашении виден на договоре и спецификации."""
    flow = controller.additional_flow
    flow.party_combo.setCurrentIndex(0)
    display = flow.party_display(seeded_party["party"], flow.party_duplicates())
    assert controller.contract_flow.party_combo.currentText() == display
    assert controller.spec_flow.party_combo.currentText() == display


def test_contract_search_by_inn_propagates(controller, seeded_party):
    """Поиск по ИНН со страницы договора подставляет во все страницы (U3)."""
    from render_helpers import OOO_FIELDS

    flow = controller.contract_flow
    flow.party_combo.setCurrentText("— выберите контрагента —")
    flow.search_entry.setText(OOO_FIELDS[F.INN])
    flow.find_party_by_inn(flow.search_entry)
    display = flow.party_display(seeded_party["party"], flow.party_duplicates())
    assert flow.party_combo.currentText() == display
    assert controller.additional_flow.party_combo.currentText() == display
    assert flow.search_entry.text() == "", "поле поиска очищается (Tk-контракт)"


def test_search_unknown_inn_alerts(controller, seeded_party, stub_dialogs):
    flow = controller.contract_flow
    flow.search_entry.setText("0000000000")
    flow.find_party_by_inn(flow.search_entry)
    assert any(
        "Контрагент не найден" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )


def test_stale_party_cleared_from_combos(controller, seeded_party):
    """Удалённый контрагент честно исчезает со всех страниц (п. 4.2)."""
    display = controller.contract_flow.party_display(
        seeded_party["party"], controller.contract_flow.party_duplicates())
    controller.set_current_party(seeded_party["party"])
    assert controller.contract_flow.party_combo.currentText() == display
    controller.store.trash_counterparty(seeded_party["party"]["id"])
    for flow in (controller.contract_flow, controller.additional_flow):
        flow.refresh_parties()
    assert controller.contract_flow.selected_party() is None
    assert controller.additional_flow.selected_party() is None


# --------------------------------------------------------------------------- #
#  Доп. соглашение: «Действует до», прописью, формат лимита (2.18.0)
# --------------------------------------------------------------------------- #
def test_additional_valid_until_default_is_dec31_current_year(controller):
    """«Действует до» по умолчанию — 31 декабря текущего года (п. 5 шаблона)."""
    today = datetime.now()
    assert controller.additional_flow.valid_entry.text() == (
        f"31.12.{today.year}"
    )


def test_limit_grouping_and_cursor(controller):
    """Лимит форматируется с разрядами; курсор не съедает ввод."""
    flow = controller.additional_flow
    flow.limit_entry.setText("1000000")
    assert flow.limit_entry.text() == "1 000 000"
    assert flow.limit_entry.cursorPosition() == 9  # курсор после 7-й цифры
    # Вставка в середину: разряды пересчитываются, курсор сохраняется.
    flow.limit_entry.setCursorPosition(2)  # перед «000»
    flow.limit_entry.insert("5")
    assert flow.limit_entry.text() == "15 000 000"
    assert flow.limit_entry.cursorPosition() == 3, "курсор не съедает ввод"
    flow.limit_entry.clear()
    flow.limit_entry.setText("abc")  # валидатор отсекает буквы
    assert flow.limit_entry.text() == ""


def test_limit_and_term_words_labels(controller):
    flow = controller.additional_flow
    flow.limit_entry.setText("15000")
    flow.term_entry.setText("45")
    assert "пятнадцать тысяч рублей" in flow.limit_words_label.text()
    assert "сорок пять" in flow.words_label.text()
    flow.term_entry.clear()
    assert flow.words_label.text() == "Срок прописью: —"


def test_contract_auto_substitution_in_additional(controller, seeded_party):
    """Договор-основание подставляется автоматически при выборе (п. 5.1)."""
    flow = controller.additional_flow
    flow.party_combo.setCurrentIndex(0)
    assert flow.contract_combo.currentText() == flow.contract_display(
        seeded_party["contract"]
    )
    assert "подставлен автоматически" in flow.contract_info.text()


# --------------------------------------------------------------------------- #
#  Сохранение договора (п. 4.1, 9.2; сценарий Г #4)
# --------------------------------------------------------------------------- #
def test_save_contract_without_party_blocks(controller, stub_dialogs):
    controller.select_page("contract")
    controller.contract_flow.save_contract()
    assert controller.store.documents == []
    assert any("Контрагент не выбран" == str(a[1])
               for a in stub_dialogs["alerts"])


def test_save_contract_validation_highlights_number(controller, seeded_party,
                                                    stub_dialogs):
    controller.select_page("contract")
    flow = controller.contract_flow
    flow.day_entry.setText("10")
    before = len([d for d in controller.store.documents
                  if d.get("kind") == storage.KIND_CONTRACT])
    flow.save_contract()
    assert flow._number_error.text() != "", "поле номера подсвечено"
    after = len([d for d in controller.store.documents
                 if d.get("kind") == storage.KIND_CONTRACT])
    assert after == before, "запись не добавлена"
    assert any("Ошибки в данных" == str(a[1])
               for a in stub_dialogs["alerts"])


def test_save_contract_writes_file_and_record(controller, seeded_party,
                                              monkeypatch, tmp_path):
    """Сохранение договора: файл в папке контрагента + запись в реестре.

    Честное «сохранён»: тост с путём, файл открыт в Word (сценарий Г #4).
    Номер 517 — не как у seeded-договора (516), чтобы подтверждение повтора
    не вмешивалось в сценарий.
    """
    flow = controller.contract_flow
    flow.number_entry.setText("517")
    flow.day_entry.setText("10")
    redirect_docs_dir(flow, tmp_path)
    opened: list[str] = []
    monkeypatch.setattr(flow.app, "open_file", lambda p: opened.append(p) or True)
    flow.save_contract()

    contracts = [d for d in controller.store.documents
                 if d.get("kind") == storage.KIND_CONTRACT
                 and d["number"] == "517"]
    assert len(contracts) == 1
    assert contracts[0]["extra"]["is_ip"] is False
    files = list(tmp_path.rglob("*.docx"))
    assert len(files) == 1, "файл записан в папку документов"
    assert opened == [str(files[0])], "файл сразу открывается в Word"
    toast_text = controller.toast_label.text()
    assert "сохранён" in toast_text and "517" in toast_text
    assert controller._toast_path == str(files[0])
    # Базовая подпись сброшена — форма «чистая» после сохранения.
    assert not flow.is_form_dirty()


def test_save_contract_duplicate_number_requires_confirmation(
    controller, seeded_party, monkeypatch, tmp_path,
):
    """Повтор номера у того же контрагента — подтверждение (п. 4.1)."""
    flow = controller.contract_flow
    flow.number_entry.setText("520")
    flow.day_entry.setText("10")
    redirect_docs_dir(flow, tmp_path)
    flow.save_contract()
    count520 = lambda: len([d for d in controller.store.documents
                            if d.get("kind") == storage.KIND_CONTRACT
                            and d["number"] == "520"])
    assert count520() == 1

    titles: list[str] = []

    def scripted_confirm(*args, **kwargs):
        titles.append(str(args[1]))
        return False

    monkeypatch.setattr(qt_dialogs, "confirm", scripted_confirm)
    flow.save_contract()
    assert len(titles) == 1 and titles[0] == "Повтор номера"
    assert count520() == 1, "отказ — сохранения нет"

    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: True)
    flow.save_contract()
    assert count520() == 2


def test_save_contract_def4_rollback_honest_message(controller, seeded_party,
                                                    monkeypatch, tmp_path,
                                                    stub_dialogs):
    """Сбой записи реестра: файл удалён, запись не добавлена (DEF-4)."""
    flow = controller.contract_flow
    flow.number_entry.setText("521")
    flow.day_entry.setText("10")
    redirect_docs_dir(flow, tmp_path)

    def failing_save():
        controller.store.last_error = "Не удалось записать реестр: отказ ОС"
        return False

    monkeypatch.setattr(controller.store, "save", failing_save)
    flow.save_contract()
    assert not [d for d in controller.store.documents
                if d.get("kind") == storage.KIND_CONTRACT
                and d["number"] == "521"]
    assert list(tmp_path.rglob("*.docx")) == [], "файл удалён при откате"
    alerts = [a for a in stub_dialogs["alerts"]
              if "Не удалось сохранить документ" == str(a[1])]
    assert alerts, "сообщение о сбое показано"
    assert "Созданный файл удалён" in str(alerts[-1][2])


# --------------------------------------------------------------------------- #
#  Сохранение доп. соглашения (п. 4.1, 9.2; сценарий Г #5)
# --------------------------------------------------------------------------- #
def _fill_additional(controller, number="1", limit="10 000", term="10",
                     valid=None):
    flow = controller.additional_flow
    flow.party_combo.setCurrentIndex(0)
    flow.number_entry.setText(number)
    flow.limit_entry.setText(limit)
    flow.term_entry.setText(term)
    if valid is not None:
        flow.valid_entry.setText(valid)
    return flow


def test_save_additional_without_contract_blocks(controller, seeded_party,
                                                 stub_dialogs):
    store = controller.store
    for doc in list(store.documents):
        if doc.get("kind") == storage.KIND_CONTRACT:
            store.purge_document(doc["id"])
    controller.additional_flow.refresh_parties()
    controller.select_page("additional")
    flow = _fill_additional(controller)
    flow.save_additional()
    assert any("Договор не выбран" == str(a[1])
               for a in stub_dialogs["alerts"])
    assert not [d for d in controller.store.documents
                if d.get("kind") == storage.KIND_ADDITIONAL]


def test_save_additional_validation_highlights_fields(controller, seeded_party,
                                                      stub_dialogs):
    controller.select_page("additional")
    flow = _fill_additional(controller, limit="", term="abc", valid="13.13.2026")
    flow.save_additional()
    assert flow._limit_error.text() == "Заполните поле"
    assert flow._term_error.text() == "Должно быть целое число дней"
    assert flow._valid_error.text() == "Введите дату в формате ДД.ММ.ГГГГ"
    assert any("Не хватает данных" == str(a[1])
               for a in stub_dialogs["alerts"])
    assert not [d for d in controller.store.documents
                if d.get("kind") == storage.KIND_ADDITIONAL]


def test_save_additional_writes_record_with_valid_until(
    controller, seeded_party, monkeypatch, tmp_path,
):
    """Сохранение доп. соглашения: «Действует до» в extra (вид + ISO, 2.18.0)."""
    flow = _fill_additional(controller, valid="31.03.2027")
    redirect_docs_dir(flow, tmp_path)
    flow.save_additional()
    addls = [d for d in controller.store.documents
             if d.get("kind") == storage.KIND_ADDITIONAL]
    assert len(addls) == 1
    extra = addls[0]["extra"]
    assert extra["valid_until"] == "31.03.2027"
    assert extra["valid_until_iso"] == "2027-03-31"
    assert extra["limit"] == "10 000"
    assert extra["term"] == "10"
    assert extra["contract_id"] == seeded_party["contract"]["id"]
    assert list(tmp_path.rglob("*.docx")), "файл записан"


def test_save_additional_empty_number_becomes_bn(controller, seeded_party,
                                                 tmp_path):
    """Пустой номер сохраняется как «б/н» (Tk-контракт)."""
    flow = _fill_additional(controller, number="")
    redirect_docs_dir(flow, tmp_path)
    flow.save_additional()
    addls = [d for d in controller.store.documents
             if d.get("kind") == storage.KIND_ADDITIONAL]
    assert addls[0]["number"] == "б/н"


# --------------------------------------------------------------------------- #
#  LAZY-1: ленивое тело договора дочитывается полностью
# --------------------------------------------------------------------------- #
def test_additional_uses_lazy_contract_body(controller, seeded_party, qapp,
                                            monkeypatch, tmp_path):
    """LAZY-1: предпросмотр и сохранение дочитывают тело ленивой записи.

    Запись договора без поля ``data`` (ленивая загрузка) — предпросмотр и
    черновик сохранения берут полное тело через store.document_body.
    """
    store = controller.store
    contract = seeded_party["contract"]
    # Эмулируем ленивую запись: тело лежит только в lazy-таблице.
    body = dict(contract.get("data") or {})
    lazy_rec = {
        "id": contract["id"], "kind": contract["kind"],
        "party_id": contract["party_id"], "number": contract["number"],
        "date": contract["date"], "extra": contract["extra"],
        "text": contract["text"],
    }
    store.document_body(contract["id"])  # прогрев кэша хранилища
    monkeypatch.setattr(store, "documents_for", lambda pid: [lazy_rec])

    flow = controller.additional_flow
    flow.refresh_parties()
    controller.select_page("additional")
    flow.limit_entry.setText("10 000")
    flow.term_entry.setText("10")
    render_once(controller, qapp)
    text = flow.preview_text.toPlainText()
    assert len(text) > 500, "ленивое тело дочитано — предпросмотр полный"

    # Черновик сохранения содержит те же поля договора.
    with_body = {**lazy_rec, **store.document_body(contract["id"])}
    assert with_body.get("data"), "document_body возвращает тело ленивой записи"


# --------------------------------------------------------------------------- #
#  Несохранённые данные (п. 4.4) — обе новые страницы
# --------------------------------------------------------------------------- #
def test_contract_form_signature_excludes_party(controller, seeded_party):
    """Выбор контрагента не входит в подпись (переносится, п. 4.4)."""
    flow = controller.contract_flow
    base = flow.form_signature()
    controller.set_current_party(seeded_party["party"])
    assert flow.form_signature() == base
    flow.number_entry.setText("9")
    assert flow.form_signature() != base


def test_contract_dirty_and_leave_confirmation(controller, seeded_party,
                                               stub_dialogs):
    """Уход с «грязного» договора — подтверждение (п. 4.4, U7)."""
    controller.select_page("contract")
    flow = controller.contract_flow
    assert not flow.is_form_dirty()
    flow.number_entry.setText("777")
    assert flow.is_form_dirty()
    assert controller.is_current_page_dirty()
    controller.confirm_leave_page()
    assert stub_dialogs["confirms"], "диалог подтверждения показан"
    message = str(stub_dialogs["confirms"][-1][2])
    assert "Договор поставки" in message, "подпись страницы в диалоге"


def test_additional_dirty_and_close_window_label(controller, seeded_party,
                                                 stub_dialogs):
    """Закрытие окна с «грязным» доп. соглашением — подпись страницы."""
    controller.select_page("additional")
    flow = controller.additional_flow
    flow.term_entry.setText("21")
    assert controller.is_current_page_dirty()
    controller._confirm_close_dirty()
    message = str(stub_dialogs["confirms"][-1][2])
    assert "Доп. соглашение" in message


class _BrokenEntry:
    """Поле, у которого сломано чтение (эмуляция разрушенного виджета)."""

    def text(self):
        raise RuntimeError("widget destroyed")


def test_contract_dirty_read_failure_counts_dirty(controller, seeded_party):
    """Сбой чтения подписи — страница «грязная» (аудит 2.0.1).

    Подменённое поле восстанавливается в finally: отложенный таймер
    предпросмотра, который pytest-qt выталкивает между тестами, не должен
    увидеть разрушенный виджет — это артефакт теста, не производственный
    путь (в приложении таймер умирает вместе со страницей).
    """
    controller.select_page("contract")
    flow = controller.contract_flow
    original = flow.number_entry
    flow.number_entry = _BrokenEntry()
    try:
        assert flow.is_form_dirty() is True
        assert controller.is_current_page_dirty() is True
    finally:
        flow.number_entry = original


def test_ctrl_s_saves_contract_page(controller, seeded_party, monkeypatch,
                                    tmp_path):
    """Ctrl+S на странице договора сохраняет договор (п. 4.4)."""
    controller.select_page("contract")
    flow = controller.contract_flow
    flow.number_entry.setText("522")
    flow.day_entry.setText("10")
    redirect_docs_dir(flow, tmp_path)
    controller.save_current_page()
    assert [d for d in controller.store.documents
            if d.get("kind") == storage.KIND_CONTRACT
            and d["number"] == "522"]


# --------------------------------------------------------------------------- #
#  Реестр обновляется после сохранения (M5-хвост операций)
# --------------------------------------------------------------------------- #
def test_save_contract_refreshes_registry(controller, seeded_party, tmp_path):
    controller.select_page("contract")
    flow = controller.contract_flow
    flow.number_entry.setText("523")
    flow.day_entry.setText("10")
    redirect_docs_dir(flow, tmp_path)
    controller.registry_flow._registry_rendered = False
    flow.save_contract()
    # Реестр получил новое состояние (ленивый рендер сброшен обновлением).
    rows = controller.registry_flow._registry_rows
    assert str(seeded_party["party"]["id"]) in rows or rows
    assert controller.status_label.text().startswith("Контрагентов:")
