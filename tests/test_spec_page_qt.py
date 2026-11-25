r"""Тесты CP1 — страница «Спецификация» на Qt (этап M3 миграции Tk → PySide6).

Проверяют критерии контрольной точки CP1 (раздел Е плана миграции) на
offscreen-платформе (fixture ``qapp`` pytest-qt, APPDATA изолируется
``appdata``; модальные диалоги подменяются на уровне qt_dialogs):

* PERF-1: вставка блока из Excel — лимит 500 с учётом стартовой ячейки,
  «вставлено 500, отброшено 100», ровно до лимита — без сообщения
  (Qt-обвязка над тем же spec_table_core.paste_plan, что и Tk);
* U12: правка ячейки меняет сумму строки и итог; итог строки под
  таблицей, предпросмотра и печатной формы — один spec_total;
* undo (снапшот + подтверждение деструктивных действий);
* предпросмотр — печатная форма из шаблона без записи на диск,
  сигнатура совпадает с golden-эталоном спецификации (текст не изменился);
* сохранение через services.save_document: файл в папке контрагента и
  запись в реестре; DEF-4 — честное сообщение при сбое записи, без
  файла-сироты; DEF-2 — документ с незаполненным {{…}} не сохраняется,
  сообщение называет токен;
* Ctrl+S; уход со страницы и закрытие окна с «грязной» формой —
  подтверждение, без изменений — без диалога (п. 4.4);
* кросс-совместимость: база после Qt-сохранения читается свежим
  DataStore, schema_version прежняя (заготовка отката на Tk-сборку).

W-1/W-2 и Tk-часть тестов этот файл не трогает.
"""

from __future__ import annotations

import json
import os
import sqlite3
import sys
from decimal import Decimal
from pathlib import Path

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import document_builder
import main_qt
import qt_dialogs
import services
import shell_open
import spec_table_core
import storage
from document_builder import F, docx_to_text
from render_helpers import OOO_FIELDS, SPEC_ITEMS, document_signature
from PySide6.QtGui import QValidator
from PySide6.QtWidgets import QLabel

REPO_ROOT = Path(__file__).resolve().parent.parent
GOLDEN_DIR = Path(__file__).resolve().parent / "golden"

ITEM_ROWS = ("Кабель ВВГнг 3х2.5", "120", "м", "85,5")
ITEM_TOTAL_TEXT = "Позиций: 10 · Итого: 10260,00"
ITEM_TOTAL_TEXT2 = "Позиций: 10 · Итого: 18660,00"
DUPLICATE_CONFIRM_CALLS: list = []  # вызовы подтверждения повтора (п. 4.1)


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_dialogs(monkeypatch):
    """Подменяет модальные диалоги: подтверждения True, алерты собираются.

    Контроллеры резолвят qt_dialogs.alert/confirm на каждый вызов —
    подмена видна и через AppController.confirm, и через виджеты.
    """
    captured = {"confirms": [], "alerts": []}

    def fake_confirm(*args, **kwargs):
        captured["confirms"].append(args)
        return True

    captured["confirm"] = fake_confirm
    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(
        qt_dialogs, "alert", lambda *a, **k: captured["alerts"].append(a)
    )
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs):
    """Каркас с живой страницей спецификации; таймеры останавливаются после."""
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl._stop_timers()
    if ctrl.spec_flow.preview_timer is not None:
        ctrl.spec_flow.preview_timer.stop()


@pytest.fixture()
def seeded_party(controller):
    """Контрагент ООО и договор-основание — данные golden-образца (render_helpers)."""
    store = controller.store
    fields = {k: v for k, v in OOO_FIELDS.items()
              if k not in (F.NUMBER, F.DAY, F.MONTH)}
    party = store.add_counterparty("ooo", fields)
    contract = store.add_document(
        storage.KIND_CONTRACT, party["id"], OOO_FIELDS[F.NUMBER],
        f"{OOO_FIELDS[F.DAY]} {OOO_FIELDS[F.MONTH]}",
        # data — полные поля, как в golden-образце (дата договора печатается
        # из data контракта: render_spec → PH.DATE/PH.MONTH).
        data=dict(OOO_FIELDS),
        extra={"day": OOO_FIELDS[F.DAY], "month": OOO_FIELDS[F.MONTH],
               "is_ip": False},
        text="",
    )
    controller.spec_flow.refresh_parties()
    return {"party": party, "contract": contract}


def fill_number(controller: main_qt.AppController, number: str = "12") -> None:
    controller.spec_flow.number_entry.setText(number)


def fill_item(controller: main_qt.AppController, row: int = 0,
              values=ITEM_ROWS) -> None:
    name, qty, unit, price = values
    table = controller.spec_flow.table
    table.item(row, 1).setText(name)
    table.item(row, 2).setText(qty)
    table.item(row, 3).setText(unit)
    table.item(row, 4).setText(price)


def fill_golden_items(controller) -> None:
    """Строки golden-образца SPEC_ITEMS в таблицу страницы."""
    flow = controller.spec_flow
    flow.table.set_row_count(max(spec_table_core.MIN_ROWS, len(SPEC_ITEMS)))
    for row, item in enumerate(SPEC_ITEMS):
        flow.table.item(row, 1).setText(item["name"])
        flow.table.item(row, 2).setText(str(item["qty"]))
        flow.table.item(row, 3).setText(item["unit"])
        flow.table.item(row, 4).setText(str(item["price"]).replace(".", ","))


# --------------------------------------------------------------------------- #
#  Страница: установка, контекст (п. 4.2)
# --------------------------------------------------------------------------- #
def test_spec_page_replaces_placeholder(controller):
    flow = controller.spec_flow
    assert controller.pages["spec"] is flow.page
    assert controller.stack.currentWidget() is flow.page
    assert controller.stack.indexOf(flow.page) >= 0


def test_placeholder_pages_remaining(controller):
    """Живы «Спецификация» (M3), «Добавить контрагента» (M4), «Реестр» (M5),
    «Договор» и «Доп. соглашение» (M6); остальные сообщают этап
    (M2-контракт, продлевается честно на каждый этап)."""
    for key, _title, stage in main_qt.NAV_PAGES:
        page = controller.pages[key]
        texts = [label.text() for label in page.findChildren(QLabel)]
        if key in ("spec", "counterparty", "registry", "contract", "additional",
                   "settings"):
            # живые страницы M3/M4/M5/M6
            assert not any("в разработке" in text for text in texts), key
        else:
            assert any("в разработке" in text for text in texts), key
            assert any(stage in text for text in texts), key


def test_party_selection_sets_contract_and_bar(controller, seeded_party):
    flow = controller.spec_flow
    assert flow.party_combo.currentText() == flow.party_display(
        seeded_party["party"], flow.party_duplicates()
    )
    assert flow.contract_combo.currentText() == flow.contract_display(
        seeded_party["contract"]
    )
    assert "подставлен автоматически" in flow.contract_info.text()
    assert seeded_party["party"]["id"] == controller.current_party_id
    assert flow.current_party_label.text().startswith("Документ для:")


def test_party_search_by_inn(controller, seeded_party):
    flow = controller.spec_flow
    flow.party_combo.setCurrentText("— выберите контрагента —")
    flow.search_entry.setText(OOO_FIELDS[F.INN])
    flow.find_party_by_inn()
    assert flow.party_combo.currentText() == flow.party_display(
        seeded_party["party"], flow.party_duplicates()
    )
    assert flow.search_entry.text() == "", "поле поиска очищается (Tk-контракт)"


def test_party_search_unknown_inn_alerts(controller, seeded_party, stub_dialogs):
    flow = controller.spec_flow
    flow.search_entry.setText("0000000000")
    flow.find_party_by_inn()
    assert any(
        "Контрагент не найден" == str(alert[1]) for alert in stub_dialogs["alerts"]
    )


def test_search_entry_validator_digits_only(controller):
    """В поле поиска проходят только цифры (Tk: _digits_vcmd_for(12))."""
    validator = controller.spec_flow.search_entry.validator()
    assert validator is not None
    assert validator.validate("7701234567", 0)[0] == QValidator.State.Acceptable
    assert validator.validate("12a", 0)[0] == QValidator.State.Invalid


# --------------------------------------------------------------------------- #
#  PERF-1: вставка блока из Excel (критерий CP1 № 2)
# --------------------------------------------------------------------------- #
def _tsv(rows: int) -> str:
    return "\n".join(f"Позиция {n}\t1\tшт\t10" for n in range(1, rows + 1))


def test_paste_600_rows_clamps_to_500_with_notice(controller, qapp, seeded_party):
    flow = controller.spec_flow
    notices: list[str] = []
    flow.table.notice.connect(notices.append)
    qapp.clipboard().setText(_tsv(600))
    flow.table.paste_clipboard()
    assert flow.table.rowCount() == spec_table_core.MAX_ROWS == 500
    assert len(notices) == 1
    assert "Вставлено 500" in notices[0] and "отброшено 100" in notices[0]
    assert flow.totals_label.text() == "Позиций: 500 · Итого: 5000,00"


def test_paste_exact_limit_no_notice(controller, qapp, seeded_party):
    flow = controller.spec_flow
    notices: list[str] = []
    flow.table.notice.connect(notices.append)
    qapp.clipboard().setText(_tsv(500))
    flow.table.paste_clipboard()
    assert flow.table.rowCount() == 500
    assert notices == [], "вставка ровно до лимита — без сообщения"


def test_paste_respects_start_cell(controller, qapp, seeded_party):
    """Вставка в непустую таблицу со сдвигом: 250 + 300 → 250 строк (PERF-1)."""
    flow = controller.spec_flow
    fill_item(controller, row=0, values=("Начальная позиция", "1", "шт", "1"))
    qapp.clipboard().setText(_tsv(300))
    flow.table.paste_block_at(250, 1)  # Ctrl+V в ячейку «Наименование» 251-й строки
    assert flow.table.rowCount() == spec_table_core.MAX_ROWS
    assert flow.table._text(250, 1) == "Позиция 1"
    assert flow.table._text(499, 1) == "Позиция 250"


def test_paste_sanitizes_garbage_in_numeric_cells(controller, qapp, seeded_party):
    """Мусор в «Кол-во»/«Цена» отсекается (аудит 2.0.1, P2) — Qt-обвязка."""
    flow = controller.spec_flow
    qapp.clipboard().setText("Болт М8\tabc\tшт\t150 руб")
    flow.table.paste_clipboard()
    values = flow.table.row_values()[0]
    assert values["qty"] == ""
    assert values["price"] == "150"


def test_paste_notice_shown_as_toast(controller, qapp, seeded_party):
    """PERF-1 сообщение — немодальным тостом каркаса (Tk: on_notice=app._notify)."""
    flow = controller.spec_flow
    qapp.clipboard().setText(_tsv(600))
    flow.table.paste_clipboard()
    assert "Вставлено 500" in controller.toast_label.text()


# --------------------------------------------------------------------------- #
#  U12: правка меняет сумму строки и итог; один spec_total (критерий № 3)
# --------------------------------------------------------------------------- #
def test_cell_edit_updates_row_and_total(controller, seeded_party):
    flow = controller.spec_flow
    fill_item(controller)
    assert document_builder.spec_total(flow.table.get_items()) == Decimal("10260.00")
    assert flow.totals_label.text() == ITEM_TOTAL_TEXT
    fill_item(controller, row=1, values=("Розетка", "40", "шт", "210"))
    assert flow.totals_label.text() == "Позиций: 10 · Итого: 18660,00"
    assert document_builder.spec_total(
        flow.table.get_items()
    ) == Decimal("18660.00")


def test_status_total_matches_preview_and_record(controller, seeded_party):
    """«Один spec_total»: строка итогов = предпросмотр = печатная форма."""
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    flow.flush_spec_preview()
    total = document_builder.spec_total(flow.table.get_items())
    formatted = document_builder.format_amount(total)  # «10 260»
    # Строка итогов под таблицей (формат Tk: две копейки, без разрядов):
    assert flow.totals_label.text() == ITEM_TOTAL_TEXT
    # Предпросмотр — печатная форма из шаблона, сумма та же:
    preview = flow.preview_text.toPlainText()
    assert formatted in preview
    # Запись реестра (extra.total) — то же число:
    rec = flow.build_document_record()
    assert rec["extra"]["total"] == float(total)
    # Печатная форма из записи — тот же итог (путь сохранения):
    doc = services.build_record_docx(rec, controller.store)
    assert formatted in docx_to_text(doc)


# --------------------------------------------------------------------------- #
#  Undo и подтверждения (критерий CP1 № 4)
# --------------------------------------------------------------------------- #
def test_undo_restores_deleted_row(controller, seeded_party):
    flow = controller.spec_flow
    fill_item(controller)
    flow.table.setCurrentCell(0, 1)
    flow.table._delete_row(0)
    assert flow.table._text(0, 1) == "", "строка удалена"
    assert flow.table.undo() is True
    assert flow.table._text(0, 1) == ITEM_ROWS[0], "undo вернул данные"
    assert flow.totals_label.text() == ITEM_TOTAL_TEXT


def test_destructive_declined_keeps_row(controller, seeded_party, monkeypatch):
    """Отказ в подтверждении оставляет строку (U7-семантика каркаса)."""
    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: False)
    flow = controller.spec_flow
    fill_item(controller)
    flow.table.setCurrentCell(0, 1)
    flow.table._delete_row(0)
    assert flow.table._text(0, 1) == ITEM_ROWS[0]
    assert flow.table.undo() is False, "undo не сработал — снапшот не пушнулся"


def test_filled_row_delete_requires_confirmation(controller, seeded_party,
                                                 stub_dialogs):
    flow = controller.spec_flow
    fill_item(controller)
    flow.table.setCurrentCell(0, 1)
    flow.table._delete_row(0)
    assert any("Удалить строку с данными?" in str(args)
               for args in stub_dialogs["confirms"])


# --------------------------------------------------------------------------- #
#  Предпросмотр и golden-эталон (критерий CP1 № 5)
# --------------------------------------------------------------------------- #
def test_preview_built_from_template_without_disk_write(controller, seeded_party,
                                                        appdata):
    """Предпросмотр = печатная форма из шаблона; на диске ничего нет."""
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    profile_dir = Path(appdata)
    before = {str(p) for p in profile_dir.rglob("*") if p.suffix in (".docx", ".tmp")}
    flow.flush_spec_preview()
    text = flow.preview_text.toPlainText()
    after = {str(p) for p in profile_dir.rglob("*") if p.suffix in (".docx", ".tmp")}
    assert "СПЕЦИФИКАЦИЯ" in text
    assert ITEM_ROWS[0] in text
    assert before == after, "предпросмотр не пишет на диск"


def test_preview_matches_golden_spec_ooo(controller, seeded_party):
    """Печатная форма из сборки страницы совпадает с golden spec_ooo.json."""
    flow = controller.spec_flow
    # Параметры golden-образца (render_helpers.spec_rec): день/месяц/поля.
    flow.day_entry.setText("15")
    flow.month_combo.setCurrentText("января")
    fill_number(controller, "12")
    flow.prices_entry.setText("")
    flow.delivery_term_entry.setText("10 дней")
    flow.payment_terms_entry.setText("отсрочка 30 дней")
    flow.delivery_basis_entry.setText("самовывоз")
    fill_golden_items(controller)
    rec = flow.build_document_record()
    doc = services.build_record_docx(rec, controller.store)
    signature = document_signature(doc)
    golden = json.loads((GOLDEN_DIR / "spec_ooo.json").read_text(encoding="utf-8"))
    assert signature["text"] == golden["text"], (
        "текст печатной формы спецификации изменился — предпросмотр CP1 "
        "собирает форму, отличную от эталона"
    )
    assert signature["tables"] == golden["tables"]


def test_preview_def2_shows_unfilled_tokens(controller, seeded_party, monkeypatch):
    """Предпросмотр показывает остаточные {{…}} до попытки сохранения (DEF-2).

    Подмена build_record_docx — как в Tk-тестах test_services_disk: сценарий
    «незаполненный {{item_*}}» на реальном шаблоне достигается только
    пустой таблицей позиций, которую save_spec блокирует раньше.
    """
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    doc = document_builder.docx_from_text("Спецификация {{item_1}} и {{INN}}")
    monkeypatch.setattr(services, "build_record_docx", lambda draft, store_: doc)
    flow.flush_spec_preview()
    text = flow.preview_text.toPlainText()
    assert "остались незаполненные поля" in text
    assert "{{item_1}}" in text


# --------------------------------------------------------------------------- #
#  Сохранение (критерии CP1 № 6, 7)
# --------------------------------------------------------------------------- #
def test_save_creates_file_in_party_folder_and_registry_record(
    controller, seeded_party, monkeypatch, tmp_path, stub_dialogs
):
    """Файл в папке контрагента + запись в реестре + тост и статус."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    flow.day_entry.setText("10")
    flow.month_combo.setCurrentText("января")
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    flow.save_spec()

    files = list(tmp_path.rglob("*.docx"))
    assert len(files) == 1, "файл записан в индивидуальную папку контрагента"
    record = controller.store.documents[-1]
    assert record["kind"] == storage.KIND_SPEC
    assert record["number"] == "12"
    assert record["date"] == "10 января"
    assert record["extra"]["total"] == 10260.0
    assert len(record["extra"]["items"]) == 1
    assert controller.store.get_document(record["id"]) is not None
    # Текст записи — предпросмотр печатной формы (Tk-контракт draft["text"]).
    assert "СПЕЦИФИКАЦИЯ" in record["text"]
    # После сохранения форма «чистая» (baseline пересчитан).
    assert flow.is_form_dirty() is False
    assert "Спецификация № 12" in controller.toast_label.text()
    # В реестре договор-основание + спецификация:
    assert "Документов: 2" in controller.status_label.text()


def _specs_of(controller) -> list:
    """Спецификации реестра (договор-основание — тоже документ, не мешает)."""
    return [d for d in controller.store.documents
            if d["kind"] == storage.KIND_SPEC]


def test_save_without_number_blocks_and_alerts(controller, seeded_party,
                                               stub_dialogs):
    flow = controller.spec_flow
    fill_item(controller)
    flow.save_spec()
    assert flow.number_error.text() == "Заполните поле"
    assert _specs_of(controller) == []
    assert any("Не хватает данных" == str(a[1]) for a in stub_dialogs["alerts"])


def test_save_without_items_blocks(controller, seeded_party, stub_dialogs):
    flow = controller.spec_flow
    fill_number(controller, "12")
    flow.save_spec()
    assert _specs_of(controller) == []
    assert any("Нет позиций" == str(a[1]) for a in stub_dialogs["alerts"])


def test_save_without_party_blocks(qapp, appdata, stub_dialogs):
    """Без контрагента сохранение честно отклоняется (Tk-контракт)."""
    ctrl = main_qt.AppController(qapp)
    try:
        ctrl.spec_flow.save_spec()
        assert ctrl.store.documents == []
        assert any(
            "Контрагент не выбран" == str(a[1]) for a in stub_dialogs["alerts"]
        )
    finally:
        ctrl._stop_timers()


def test_duplicate_number_requires_confirmation(controller, seeded_party,
                                                monkeypatch, tmp_path):
    """Повтор номера у того же контрагента — подтверждение (п. 4.1)."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    flow.save_spec()  # первое сохранение — дубля нет, подтверждение не нужно
    assert len(_specs_of(controller)) == 1
    assert len(DUPLICATE_CONFIRM_CALLS) == 0

    titles: list[str] = []

    def scripted_confirm(*args, **kwargs):
        titles.append(str(args[1]))
        return False

    monkeypatch.setattr(qt_dialogs, "confirm", scripted_confirm)
    flow.save_spec()  # отказ — сохранения нет
    assert len(_specs_of(controller)) == 1
    assert titles == ["Повтор номера"], "подтверждение повтора показано"

    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: True)
    flow.save_spec()  # подтверждение — создан второй документ
    assert len(_specs_of(controller)) == 2


# --------------------------------------------------------------------------- #
#  DEF-4: транзакция сохранения в Qt-обвязке (критерий CP1 № 6)
# --------------------------------------------------------------------------- #
def test_def4_rollback_honest_message_and_no_orphan(controller, seeded_party,
                                                    monkeypatch, tmp_path,
                                                    stub_dialogs):
    """Сбой записи реестра: файл удалён, запись не добавлена, алерт честный."""
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")

    def failing_save():
        controller.store.last_error = "Не удалось записать реестр: отказ ОС"
        return False

    monkeypatch.setattr(controller.store, "save", failing_save)
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    flow.save_spec()

    assert _specs_of(controller) == [], "фантом записи не остаётся"
    assert list(tmp_path.rglob("*.docx")) == [], "файл удалён при откате"
    assert list(tmp_path.rglob("*.docx.tmp")) == []
    alerts = [a for a in stub_dialogs["alerts"]
              if "Не удалось сохранить документ" == str(a[1])]
    assert alerts, "сообщение о сбое показано"
    assert "Созданный файл удалён" in str(alerts[-1][2])


# --------------------------------------------------------------------------- #
#  DEF-2: предохранитель остаточных {{…}} в Qt-обвязке (критерий CP1 № 7)
# --------------------------------------------------------------------------- #
def test_def2_unfilled_token_blocks_save_and_names_token(
    controller, seeded_party, monkeypatch, tmp_path, stub_dialogs,
):
    """Документ с незаполненным {{item_*}} не сохраняется; токен назван (DEF-2).

    Qt-обвязка: сохранение идёт через services.save_document — предохранитель
    (общий с Tk) отклоняет запись; подмена build_record_docx — как в
    Tk-тестах test_services_disk, поскольку на реальном шаблоне сценарий
    «{{item_*}} остались» достигается только пустой таблицей, которую
    save_spec блокирует раньше («Нет позиций»).
    """
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    doc = document_builder.docx_from_text("Спецификация {{item_1}} и {{INN}}")
    monkeypatch.setattr(services, "build_record_docx", lambda draft, store_: doc)
    flow.save_spec()

    assert _specs_of(controller) == [], "запись не добавлялась"
    assert list(tmp_path.rglob("*.docx")) == [], "файл не записывается"
    alerts = [a for a in stub_dialogs["alerts"]
              if "остались незаполненные поля" in str(a[2])]
    assert alerts, "предохранитель DEF-2 сработал"
    assert "{{item_1}}" in str(alerts[-1][2]), "сообщение называет токен"


# --------------------------------------------------------------------------- #
#  Ctrl+S и защита несохранённых данных (критерии CP1 № 8, 9)
# --------------------------------------------------------------------------- #
def test_ctrl_s_saves_current_page(controller, seeded_party, monkeypatch,
                                   tmp_path):
    """Ctrl+S сохраняет спецификацию (контракт Tk _on_save_shortcut)."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    controller._on_save_shortcut()
    assert len(_specs_of(controller)) == 1


def test_leave_dirty_page_requires_confirmation(controller, seeded_party,
                                                monkeypatch):
    """Уход с изменениями — подтверждение; отказ оставляет страницу."""
    flow = controller.spec_flow
    fill_number(controller, "12")
    assert flow.is_form_dirty() is True
    answers = {"ok": False}
    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: answers["ok"])
    controller.select_page("registry")
    assert controller.current_page == "spec", "отказ — остались на странице"
    answers["ok"] = True
    controller.select_page("registry")
    assert controller.current_page == "registry"


def test_leave_clean_page_no_confirmation(controller, seeded_party, stub_dialogs):
    """Уход без изменений — без диалога (критерий CP1 № 8)."""
    controller.select_page("registry")
    assert controller.current_page == "registry"
    assert stub_dialogs["confirms"] == []


def test_close_dirty_window_requires_confirmation(controller, seeded_party,
                                                  qapp, monkeypatch):
    """Закрытие с «грязной» формой — подтверждение; отказ не закрывает."""
    flow = controller.spec_flow
    fill_number(controller, "12")
    answers = {"ok": False}
    monkeypatch.setattr(qt_dialogs, "confirm", lambda *a, **k: answers["ok"])
    controller.window.close()
    qapp.processEvents()
    assert controller.window.isVisible(), "отказ — окно открыто"
    answers["ok"] = True
    controller.window.close()
    qapp.processEvents()
    assert not controller.window.isVisible(), "подтверждение — окно закрыто"


def test_close_clean_window_accepted_without_dialog(controller, seeded_party,
                                                    qapp, stub_dialogs):
    controller.window.close()
    qapp.processEvents()
    assert not controller.window.isVisible()
    assert stub_dialogs["confirms"] == []


def test_dirty_on_signature_failure_treated_as_dirty(controller, seeded_party,
                                                     monkeypatch):
    """Сбой чтения подписи — страница «грязная» (аудит 2.0.1)."""
    flow = controller.spec_flow
    flow.reset_form_baseline()
    assert flow.is_form_dirty() is False

    def broken_snapshot():
        raise RuntimeError("сбой чтения формы")

    monkeypatch.setattr(type(flow.table), "snapshot", lambda self: broken_snapshot())
    assert flow.is_form_dirty() is True
    # Защита каркаса (is_current_page_dirty) тоже честно считает «грязной»:
    assert controller.is_current_page_dirty() is True


# --------------------------------------------------------------------------- #
#  Кросс-совместимость базы (критерий CP1 № 11 — заготовка отката)
# --------------------------------------------------------------------------- #
def test_qt_saved_record_readable_by_fresh_store(controller, seeded_party,
                                                 monkeypatch, tmp_path):
    """Написано Qt → читается свежим DataStore; schema_version прежняя."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    flow = controller.spec_flow
    fill_item(controller)
    fill_number(controller, "12")
    monkeypatch.setattr(flow, "documents_dir", lambda: str(tmp_path))
    flow.save_spec()

    with sqlite3.connect(storage.db_path()) as conn:
        meta = dict(conn.execute("SELECT key, value FROM meta"))
    assert int(meta.get("schema_version", "0")) == storage.SQLITE_SCHEMA_VERSION

    fresh = storage.DataStore()
    record = fresh.documents[-1]
    assert record["kind"] == storage.KIND_SPEC
    assert record["number"] == "12"
    body = fresh.document_body(record["id"])
    assert body["extra"]["total"] == 10260.0
    assert body["extra"]["items"][0]["name"] == ITEM_ROWS[0]
    # Файл открывается python-docx (Tk 2.28.2 печатает ту же форму — путь
    # общий build_record_docx; живой прогон «Qt-файл → Tk-сборка» — runbook).
    from docx import Document

    docx_files = list(tmp_path.rglob("*.docx"))
    assert len(docx_files) == 1
    assert "СПЕЦИФИКАЦИЯ" in docx_to_text(Document(str(docx_files[0])))


# --------------------------------------------------------------------------- #
#  Оболочка: shell_open (точка M3) и честный отказ без файла
# --------------------------------------------------------------------------- #
def test_open_file_missing_file_alerts_and_returns_false(controller, stub_dialogs):
    assert controller.open_file("C:/no/such/file.docx") is False
    assert any(
        "Файл не найден" == str(a[1]) for a in stub_dialogs["alerts"]
    )


def test_shell_open_system_missing_file_returns_false(appdata):
    assert shell_open.open_file_system("C:/no/such/file.docx") is False
    assert shell_open.show_in_file_manager("C:/no/such/file.docx") is False


def test_documents_known_folder_non_windows_fallback(appdata):
    """На не-Windows — fallback «~/Documents» (контракт Tk-версии)."""
    if sys.platform == "win32":
        pytest.skip("не-Windows fallback на Windows не вызывается")
    assert shell_open.documents_known_folder().endswith("Documents")
