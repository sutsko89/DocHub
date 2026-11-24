r"""Тесты M5 — страница «Реестр» на Qt (TREE-1/UX-REG).

Проверяют критерии этапа M5 (план миграции) на offscreen-платформе
(fixture ``qapp`` pytest-qt, APPDATA изолируется ``appdata``; модальные
диалоги подменяются на уровне qt_dialogs):

* страница заменяет плейсхолдер; модель строк из RowSpec/DiffPlan,
  колонки «Контрагент · Тип · ИНН · ID · Адрес», точки типов (V4);
* поиск по названию и ИНН по всей базе, сортировка (U2) с сохранением
  в settings.json, отложенный поиск 200 мс;
* порция R-05b: «Показать ещё» с сохранением фильтра/сортировки/
  выделения; запись № 10 000 достижима, подпись «показаны все N»,
  поиск находит № 5 000 независимо от порции (реальная база 10 001);
* диф (R-04): неизменённые строки не пересоздаются, dataChanged —
  только у изменившихся, порядок после сдвига совпадает с моделью;
* «добавить 3, удалить 1, восстановить 1» — отображение совпадает
  с данными, секция «Корзина» и кнопка очистки;
* drill «Документы — <имя>»: колонки/заголовок/подстрока, возврат
  «← Назад»/Escape, удалённый контрагент — тихий возврат;
* панель действий (V3): подписи по типу записи, маршрутизация к
  потокам, Delete не касается записей корзины;
* клавиатура (U13): стрелки/Enter/Delete/Ctrl+F без конфликтов с
  полями ввода, click-toggle повторным кликом, двойной клик открывает;
* теги сроков: expired/soon красят строку красным/амберным по теме;
* кросс-совместимость: база после операций Qt-реестра читается свежим
  DataStore, schema_version прежняя.

W-1/W-2 и Tk-часть тестов этот файл не трогает.
"""

from __future__ import annotations

import os
import json
import sqlite3
import sys

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import app_common
import main_qt
import qt_dialogs
import qt_theme
import registry_core as rcore
import storage
from document_builder import F
from PySide6.QtCore import QCoreApplication, QEvent, Qt
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QDialog


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
    """Каркас с живыми страницами CP1/M4/M5; окно и таймеры убираются после.

    Окно закрывается и удаляется с явным flush DeferredDelete: обычный
    processEvents не сбрасывает отложенные удаления Qt — без flush каждый
    каркас оставлял живые виджеты, и время setup росло линейно по файлу
    (наблюдение M4, docs/qt-migration-stage-notes.md).
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


def add_party(controller, name: str, inn: str = "7701234560", ptype: str = "ooo"):
    """Добавляет контрагента напрямую в store (без UI) — эталонный rec."""
    fields = {F.INN: inn, "Сокращенное название компании": name}
    if ptype == "ooo":
        fields[F.FULL_NAME] = name
    else:
        fields[F.IP_NAME] = name
    return controller.store.add_counterparty(ptype, fields)


# --------------------------------------------------------------------------- #
#  GUI-free ядро (registry_core) — диф-план и значения строк
# --------------------------------------------------------------------------- #
def _row(iid: str, name: str = "ООО", *, tags=(), dot=None):
    return rcore.RowSpec(iid=iid, values=(name, "Тип", "ИНН", iid, "Адрес"),
                         tags=tags, dot=dot)


def test_plan_display_count_within_limit():
    assert rcore.plan_display_count(10_000, rcore.DISPLAY_SOFT_LIMIT) == (10_000, False)


def test_plan_display_count_over_limit():
    shown, has_more = rcore.plan_display_count(10_001, rcore.DISPLAY_SOFT_LIMIT)
    assert (shown, has_more) == (rcore.DISPLAY_SOFT_LIMIT, True)


def test_plan_display_count_extended_portion():
    shown, has_more = rcore.plan_display_count(50_000, rcore.DISPLAY_SOFT_LIMIT * 2)
    assert (shown, has_more) == (20_000, True)


def test_plan_diff_no_change():
    plan = rcore.plan_diff(["a", "b"], [_row("a"), _row("b")])
    assert plan.to_delete == [] and plan.ops == []


def test_plan_diff_add_edit_delete_keeps_prefix():
    plan = rcore.plan_diff(
        ["a", "b", "c"],
        [_row("a", "A2"), _row("c", "3"), _row("d", "4")],
    )
    assert plan.to_delete == ["b"]
    assert (2, "d", "insert") in plan.ops
    assert all(iid not in {"a", "c"} for _pos, iid, _op in plan.ops), \
        "неизменённые строки плана не трогаются (критерий 5.1)"


def test_plan_diff_move_single_op():
    plan = rcore.plan_diff(["a", "b", "c"], [_row("c"), _row("a"), _row("b")])
    assert plan.to_delete == []
    assert plan.ops == [(0, "c", "move")], "минимальная операция: сдвиг «c» в начало"


def test_plan_diff_append_many_is_insert_only():
    current = [f"p{i}" for i in range(1000)]
    desired = [_row(f"p{i}") for i in range(1000)] + \
              [_row(f"n{i}") for i in range(200)]
    plan = rcore.plan_diff(current, desired)
    assert len(plan.ops) == 200
    assert all(op[2] == "insert" for op in plan.ops)


def test_columns_constants_tk_parity():
    """Колонки реестра — Tk-паритет (U1): имена, заголовки, ширины."""
    assert [col[0] for col in rcore.PARTY_COLUMNS] == \
        ["name", "kind", "inn", "id", "address"]
    assert [col[1] for col in rcore.PARTY_COLUMNS] == \
        ["Контрагент", "Тип", "ИНН", "ID", "Адрес"]
    assert [col[1] for col in rcore.DOC_COLUMNS] == ["Документ", "Дата", "Срок"]


def test_party_row_values_columns():
    rec = {"id": "p-1", "type": "ooo",
           "fields": {F.INN: "7701234567", F.SHORT_NAME: "ООО Ромашка"}}

    values = rcore.party_row_values(rec)

    assert values[0] == "ООО Ромашка"
    assert values[1] == "Юр. лицо"
    assert values[2] == "7701234567"
    assert values[3] == "p-1"


def test_party_row_values_missing_inn_and_long_address():
    rec = {"id": 1, "type": "ooo",
           "fields": {F.INN: "", F.SHORT_NAME: "ООО",
                      F.LEGAL_ADDRESS: "г. Владивосток, " + "длинная улица, " * 8}}

    values = rcore.party_row_values(rec)

    assert values[2] == "—"
    assert values[4].endswith("…"), "длинный адрес усечён"
    assert rcore.party_row_dot({"type": "ip"}) == rcore.DOT_IP
    assert rcore.party_row_dot({"type": "ooo"}) == rcore.DOT_OOO


def test_doc_row_values_expiry_without_separator():
    rec = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "3",
           "date": "01.10.2026", "extra": {"valid_until_iso": "2026-01-01"}}

    _doc, _date, expiry = rcore.doc_row_values(rec)

    assert expiry == "истекло 01.01.2026", "разделитель уходит — срок в своей колонке"


def test_doc_row_tags_marks_expired_and_soon():
    expired = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "1",
               "extra": {"valid_until_iso": "2026-01-01"}}
    soon = {"id": 2, "kind": storage.KIND_ADDITIONAL, "number": "2",
            "extra": {"valid_until_iso": "2026-12-15"}}
    contract = {"id": 3, "kind": storage.KIND_CONTRACT, "number": "5", "extra": {}}

    assert rcore.doc_row_tags(expired) == ("expired",)
    assert rcore.doc_row_tags(soon) == ("soon",)
    assert rcore.doc_row_tags(contract) == ()


def test_action_labels_for_kinds():
    assert rcore.action_labels_for("party") == ("Редактировать", "Удалить")
    assert rcore.action_labels_for("doc") == ("Открыть Word", "Удалить")
    assert rcore.action_labels_for("trash_party") == ("Восстановить", "Удалить навсегда")
    assert rcore.action_labels_for("trash_doc") == ("Восстановить", "Удалить навсегда")
    assert rcore.action_labels_for("unknown") == ("Показать в папке", "Удалить")
    assert rcore.action_labels_for("none") is None


# --------------------------------------------------------------------------- #
#  Страница: установка, поиск, сортировка
# --------------------------------------------------------------------------- #
def test_registry_page_replaces_placeholder(controller):
    flow = controller.registry_flow
    assert controller.pages["registry"] is flow.page
    assert controller.stack.indexOf(flow.page) >= 0
    # Ленивый рендер: страница построена, список ещё не рисован.
    assert flow._registry_rendered is False


def test_parties_empty_query_returns_all(controller):
    add_party(controller, "ООО Ромашка", "7701234567")
    add_party(controller, "ООО Стрелка", "7709990000")

    assert len(controller.registry_flow.parties()) == 2


def test_parties_filters_by_name_substring(controller):
    add_party(controller, "ООО Ромашка", "7701234567")
    add_party(controller, "ООО Стрелка", "7709990000")
    controller.registry_flow.search_edit.setText("ромашка")

    names = [storage.display_name(rec) for rec in controller.registry_flow.parties()]
    assert names == ["ООО Ромашка"]


def test_parties_filters_by_inn_digits(controller):
    party = add_party(controller, "ООО Ромашка", "7701234567")
    add_party(controller, "ООО Стрелка", "7709990000")
    controller.registry_flow.search_edit.setText("01234567")  # подстрока ИНН

    found = controller.registry_flow.parties()
    assert [rec["id"] for rec in found] == [party["id"]]


def test_parties_no_match_returns_empty(controller):
    add_party(controller, "ООО Ромашка", "7701234567")
    controller.registry_flow.search_edit.setText("несуществующий")

    assert controller.registry_flow.parties() == []


def test_sorted_parties_by_added_keeps_store_order(controller):
    add_party(controller, "ООО Ромашка", "7701234567")
    add_party(controller, "ООО Стрелка", "7709990000")
    controller.registry_flow._registry_sort = "added"

    names = [storage.display_name(rec)
             for rec in controller.registry_flow.sorted_parties()]
    assert names == ["ООО Ромашка", "ООО Стрелка"]


def test_sorted_parties_by_name(controller):
    add_party(controller, "ООО Стрелка", "7709990000")
    add_party(controller, "ООО Ромашка", "7701234567")
    controller.registry_flow._registry_sort = "name"

    names = [storage.display_name(rec)
             for rec in controller.registry_flow.sorted_parties()]
    assert names == ["ООО Ромашка", "ООО Стрелка"], \
        "алфавитный порядок не зависит от добавления"


def test_sorted_parties_by_inn_empty_last(controller):
    add_party(controller, "ООО С ИНН", "7709990000")
    add_party(controller, "ООО Без ИНН", "")
    controller.registry_flow._registry_sort = "inn"

    names = [storage.display_name(rec)
             for rec in controller.registry_flow.sorted_parties()]
    assert names == ["ООО С ИНН", "ООО Без ИНН"], "пустой ИНН — в конец списка"


def test_on_sort_changed_persists_setting(controller):
    flow = controller.registry_flow

    flow.on_sort_changed("По ИНН")

    assert flow._registry_sort == "inn"
    assert controller.settings["registry_sort"] == "inn"
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        assert json.load(fh).get("registry_sort") == "inn", \
            "settings.json — тот же файл, что у Tk-сборки"


def test_on_sort_changed_same_value_is_noop(controller, monkeypatch):
    flow = controller.registry_flow
    saved = []
    monkeypatch.setattr(controller, "save_setting",
                        lambda key, value: saved.append((key, value)))

    flow.on_sort_changed("По добавлению")

    assert saved == [], "перестановка и запись не выполняются"


def test_on_docs_type_changed(controller):
    flow = controller.registry_flow

    flow.on_docs_type_changed("Договоры")

    assert flow._docs_type_filter == storage.KIND_CONTRACT


# --------------------------------------------------------------------------- #
#  Модель и перерисовка: полный рендер, диф, порции (R-04, R-05b)
# --------------------------------------------------------------------------- #
def test_refresh_builds_rows_and_caption(controller):
    recs = [add_party(controller, f"ООО Партия-{i}", f"770123456{i}")
            for i in range(3)]
    flow = controller.registry_flow
    flow.refresh()

    model = flow.table.model
    assert model.order() == tuple(f"p:{rec['id']}" for rec in recs)
    assert flow._registry_caption.text() == "Контрагенты (3)"
    values = model.rows[0].values
    assert values[0] == "ООО Партия-0" and values[2] == "7701234560"


def test_search_filter_refreshes_model(controller):
    add_party(controller, "ООО Ромашка", "7701234567")
    add_party(controller, "ООО Стрелка", "7709990000")
    flow = controller.registry_flow
    flow.refresh()

    flow.search_edit.setText("стрелка")
    flow.apply_filter()  # отложенный таймер не ждём — Enter-семантика

    assert flow.table.model.order() == ("p:2",)
    assert "Найдено по запросу «стрелка»: 1" in flow._registry_caption.text()


def test_show_more_extends_portion_and_caption(controller, monkeypatch):
    """Порция R-05b: «Показать ещё», подпись, достижимая последняя запись."""
    for i in range(3):
        add_party(controller, f"ООО Порция-{i}", f"770123456{i}")
    flow = controller.registry_flow
    monkeypatch.setattr(rcore, "DISPLAY_SOFT_LIMIT", 2)
    flow.reset_filter()  # полная популяция пересчитывает порцию (R-05b)

    model = flow.table.model
    assert model.order() == ("p:1", "p:2", "show-more")
    assert "показаны 2 из 3" in flow._registry_caption.text()

    flow.show_more()
    assert model.order() == ("p:1", "p:2", "p:3"), "диф добавил оставшуюся строку"
    assert "показаны все 3" in flow._registry_caption.text(), \
        "подпись «показаны все N» после исчерпания порции"


def test_refresh_full_resets_portion(controller, monkeypatch):
    add_party(controller, "ООО Один", "7701234560")
    add_party(controller, "ООО Два", "7701234561")
    add_party(controller, "ООО Три", "7701234562")
    flow = controller.registry_flow
    monkeypatch.setattr(rcore, "DISPLAY_SOFT_LIMIT", 2)
    flow.reset_filter()
    flow.show_more()
    assert flow._registry_shown == 4

    flow.on_sort_changed("По алфавиту")  # смена сортировки — полная популяция

    assert flow._registry_shown == 2, "порция сброшена при смене вида (R-05b)"


def test_show_more_row_reports_activation(controller, monkeypatch):
    add_party(controller, "ООО Один", "7701234560")
    add_party(controller, "ООО Два", "7701234561")
    add_party(controller, "ООО Три", "7701234562")
    flow = controller.registry_flow
    monkeypatch.setattr(rcore, "DISPLAY_SOFT_LIMIT", 2)
    flow.reset_filter()

    flow.on_tree_activate("show-more")

    assert flow._registry_shown == 4, "активация строки «Показать ещё» (R-05b)"


def test_diff_appends_without_recreating_prefix(controller, monkeypatch):
    """«Показать ещё» — диф: неизменённые строки не пересоздаются (5.1)."""
    for i in range(5):
        add_party(controller, f"ООО Диф-{i}", f"770123456{i}")
    flow = controller.registry_flow
    monkeypatch.setattr(rcore, "DISPLAY_SOFT_LIMIT", 3)
    flow.reset_filter()
    flow.table.set_selection("p:1")
    qapp = controller.app
    qapp.processEvents()
    assert flow.table.selected_iid() == "p:1"

    flow.show_more()
    qapp.processEvents()

    model = flow.table.model
    assert model.order() == ("p:1", "p:2", "p:3", "p:4", "p:5")
    assert flow.table.selected_iid() == "p:1", \
        "выделение пережило «Показать ещё» (контракт порции)"


def test_apply_diff_updates_values_without_recreating_rows(controller):
    rec = add_party(controller, "ООО Изменяемая", "7701234560")
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model

    rec["fields"][F.SHORT_NAME] = "ООО Изменённая"
    flow.refresh()  # диф: сигнатура не менялась

    assert model.order() == ("p:1",)
    assert model.rows[0].values[0] == "ООО Изменённая"


def test_data_changed_only_for_updated_rows(controller):
    rec = add_party(controller, "ООО До правки", "7701234560")
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model
    changes: list = []
    model.dataChanged.connect(lambda tl, br, *_a: changes.append((tl.row(), br.row())))

    rec["fields"][F.SHORT_NAME] = "ООО После правки"
    flow.refresh()

    assert len(changes) == 1 and changes[0] == (0, 0), \
        "dataChanged — только у изменившейся строки (R-04)"


def test_apply_diff_moves_to_desired_order(controller):
    """Сдвиг (R-04): итоговый порядок совпадает с желаемой моделью."""
    ids = [add_party(controller, f"ООО Порядок-{i}", f"770123456{i}")["id"]
           for i in range(3)]
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model

    rows = [
        _row(f"p:{ids[2]}", "ООО Порядок-2"),
        _row(f"p:{ids[0]}", "ООО Порядок-0"),
        _row(f"p:{ids[1]}", "ООО Порядок-1"),
    ]
    model.apply_diff(rows)
    controller.app.processEvents()

    assert model.order() == (f"p:{ids[2]}", f"p:{ids[0]}", f"p:{ids[1]}")


def test_row_decoration_dot_for_party_type(controller):
    ooo = add_party(controller, "ООО Точка", "7701234560")
    ip = add_party(controller, "ИП Точка", "254012345692", ptype="ip")
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model

    from PySide6.QtGui import QColor
    d_ooo = model.data(model.index(0, 0), Qt.ItemDataRole.DecorationRole)
    d_ip = model.data(model.index(1, 0), Qt.ItemDataRole.DecorationRole)
    assert d_ooo is not None and not d_ooo.isNull()
    assert d_ip is not None and not d_ip.isNull()
    assert model.rows[0].dot == rcore.DOT_OOO
    assert model.rows[1].dot == rcore.DOT_IP


# --------------------------------------------------------------------------- #
#  Сценарий «добавить 3, удалить 1, восстановить 1» + секция «Корзина»
# --------------------------------------------------------------------------- #
def test_add3_delete1_restore1_display_matches_data(controller):
    ids = [add_party(controller, f"ООО Сценарий-{i}", f"770123456{i}")["id"]
           for i in range(3)]
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model
    assert model.order() == tuple(f"p:{i}" for i in ids), "добавление 3 — видно"

    rec = controller.store.get_counterparty(ids[1])
    controller.lifecycle_flow.delete_counterparty(rec)

    assert model.order() == (f"p:{ids[0]}", f"p:{ids[2]}", "sec:trash", f"tp:{ids[1]}")
    trash_row = model.row(f"tp:{ids[1]}")
    assert trash_row is not None and trash_row.tags == ("muted",)
    assert flow._registry_empty_trash_btn.isHidden() is False, \
        "кнопка «Очистить корзину» появилась вместе с секцией"

    controller.lifecycle_flow.restore_counterparty(rec)

    assert model.order() == tuple(f"p:{i}" for i in ids), "восстановление вернуло запись"
    assert flow._registry_empty_trash_btn.isHidden() is True, \
        "пустая корзина: кнопка очистки скрыта (как у 2.18.0)"


def test_trash_doc_row_lists_party_name(controller):
    pid = add_party(controller, "ООО С документом", "7701234560")["id"]
    doc = controller.store.add_document(storage.KIND_CONTRACT, pid, "5",
                                        "01.10.2026", data={}, extra={}, text="")
    controller.lifecycle_flow.delete_document(doc)
    flow = controller.registry_flow
    flow.refresh()
    model = flow.table.model

    row = model.row(f"td:{doc['id']}")
    assert row is not None
    assert "ООО С документом" in row.values[0], "строка корзины называет контрагента"


# --------------------------------------------------------------------------- #
#  Drill «Документы — <имя>» (2.19.0)
# --------------------------------------------------------------------------- #
def test_drill_opens_and_back(controller):
    pid = add_party(controller, "ООО Провал", "7701234560")["id"]
    controller.store.add_document(storage.KIND_CONTRACT, pid, "5", "01.10.2026",
                                  data={}, extra={}, text="")
    flow = controller.registry_flow
    flow.refresh()

    flow.open_party_documents(controller.store.get_counterparty(pid))

    assert flow.is_drill_mode
    assert flow._registry_title_label.text() == "Документы — ООО Провал"
    assert [col[1] for col in flow.table.model.columns] == ["Документ", "Дата", "Срок"]
    assert "Документы — ООО Провал (1)" in flow._registry_caption.text()
    assert flow._registry_back_btn.isVisibleTo(flow.page)
    assert flow._registry_search_bar.isVisibleTo(flow.page) is False
    assert flow._registry_filter_bar.isVisibleTo(flow.page)

    flow.back_to_parties()

    assert not flow.is_drill_mode
    assert flow._registry_title_label.text() == "Реестр"
    assert [col[1] for col in flow.table.model.columns][0] == "Контрагент"
    assert flow._docs_selected_id is None


def test_on_page_open_resets_drill(controller):
    pid = add_party(controller, "ООО Уход", "7701234560")["id"]
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))
    assert flow.is_drill_mode

    flow.on_page_open()

    assert not flow.is_drill_mode, \
        "возврат на страницу реестра показывает список контрагентов"


def test_drill_party_deleted_returns_to_parties(controller):
    pid = add_party(controller, "ООО Пропащий", "7701234560")["id"]
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))
    controller.lifecycle_flow.delete_counterparty(controller.store.get_counterparty(pid))

    flow.refresh()

    assert not flow.is_drill_mode, \
        "мягкое удаление drill-контрагента — тихий возврат к списку (2.19.0)"


def test_drill_type_filter(controller):
    pid = add_party(controller, "ООО Фильтр", "7701234560")["id"]
    contract = controller.store.add_document(
        storage.KIND_CONTRACT, pid, "5", "01.10.2026", data={}, extra={}, text="")
    spec = controller.store.add_document(
        storage.KIND_SPEC, pid, "2", "02.10.2026", data={}, extra={}, text="")
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))

    flow.on_docs_type_changed("Спецификации")

    assert flow.table.model.order() == (f"d:{spec['id']}",)
    assert "спецификации" in flow._registry_caption.text()


# --------------------------------------------------------------------------- #
#  Панель действий (V3) и маршрутизация
# --------------------------------------------------------------------------- #
def test_action_panel_labels_per_kind(controller):
    pid = add_party(controller, "ООО Панель", "7701234560")["id"]
    doc = controller.store.add_document(storage.KIND_CONTRACT, pid, "5",
                                        "01.10.2026", data={}, extra={}, text="")
    flow = controller.registry_flow
    flow.refresh()

    flow.on_tree_select(f"p:{pid}")
    assert flow._registry_action_btn1.text() == "Редактировать"
    assert flow._registry_action_btn2.text() == "Удалить"

    flow.open_party_documents(controller.store.get_counterparty(pid))
    flow.on_tree_select(f"d:{doc['id']}")
    assert flow._registry_action_btn1.text() == "Открыть Word"
    assert flow._registry_action_btn2.text() == "Удалить"

    controller.lifecycle_flow.delete_document(doc)
    flow.back_to_parties()  # секция «Корзина» — в виде контрагентов
    flow.on_tree_select(f"td:{doc['id']}")
    assert flow._registry_action_btn1.text() == "Восстановить"
    assert flow._registry_action_btn2.text() == "Удалить навсегда"

    flow.on_tree_select(None, user=True)
    assert flow._registry_actions_hint.isVisibleTo(flow.page)
    assert flow._registry_action_btn1.isHidden()


def test_run_action_edit_opens_dialog(controller):
    pid = add_party(controller, "ООО Редакт", "7701234560")["id"]
    rec = controller.store.get_counterparty(pid)
    flow = controller.registry_flow
    flow.refresh()
    flow.on_tree_select(f"p:{pid}")

    flow.run_action("party", rec)

    dialogs = controller.window.findChildren(QDialog)
    assert len(dialogs) == 1, "«Редактировать» открывает окно правки (M4)"
    dialogs[0].deleteLater()


def test_run_action_delete_and_restore_routing(controller):
    pid = add_party(controller, "ООО Маршрут", "7701234560")["id"]
    rec = controller.store.get_counterparty(pid)
    flow = controller.registry_flow
    flow.refresh()
    flow.on_tree_select(f"p:{pid}")

    flow.run_action("party", rec, second=True)  # «Удалить» → корзина
    trashed = controller.store.trashed_counterparties()
    assert [r["id"] for r in trashed] == [pid], "мягкое удаление — запись в корзине"

    flow.refresh()
    flow.on_tree_select(f"tp:{pid}")
    flow.run_action("trash_party", trashed[0])  # «Восстановить»
    assert controller.store.get_counterparty(pid) is not None


def test_delete_key_prefers_document(controller):
    pid = add_party(controller, "ООО Дели", "7701234560")["id"]
    doc = controller.store.add_document(storage.KIND_CONTRACT, pid, "5",
                                        "01.10.2026", data={}, extra={}, text="")
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))
    flow._docs_selected_id = doc["id"]

    assert flow.delete_selected() is True
    assert doc["id"] in [d["id"] for d in controller.store.trashed_documents()]


def test_delete_key_ignores_trash_selection(controller):
    pid = add_party(controller, "ООО Корз", "7701234560")["id"]
    controller.lifecycle_flow.delete_counterparty(controller.store.get_counterparty(pid))
    flow = controller.registry_flow
    flow.refresh()
    flow.on_tree_select(f"tp:{pid}")
    assert flow._registry_selected_extra is not None

    assert flow.delete_selected() is False, \
        "Delete не удаляет навсегда запись корзины — только через панель"


def test_delete_key_in_drill_ignores_party(controller):
    pid = add_party(controller, "ООО В провале", "7701234560")["id"]
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))
    flow._docs_selected_id = None

    assert flow.delete_selected() is False, \
        "внутри документов Delete не удаляет контрагента"


def test_delete_selected_without_selection_is_noop(controller):
    assert controller.registry_flow.delete_selected() is False


def test_select_other_party_does_not_break_drill(controller):
    pid_a = add_party(controller, "ООО А-дрилл", "7701234560")["id"]
    pid_b = add_party(controller, "ООО Б-дрилл", "7701234561")["id"]
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid_a))

    flow.select_party(controller.store.get_counterparty(pid_b))

    assert flow._registry_drill_party_id == pid_a, \
        "выделение другого контрагента не меняет текущий вид"


# --------------------------------------------------------------------------- #
#  Клавиатура (U13) и мышь
# --------------------------------------------------------------------------- #
def test_keyboard_ctrl_f_focuses_search(controller):
    pid = add_party(controller, "ООО Фокус", "7701234560")["id"]
    flow = controller.registry_flow
    controller.select_page("registry")
    controller.window.activateWindow()
    flow.table.setFocus()
    controller.app.processEvents()

    QTest.keyClick(flow.table, Qt.Key.Key_F, Qt.KeyboardModifier.ControlModifier)
    controller.app.processEvents()

    assert flow.search_edit.hasFocus(), "Ctrl+F ставит фокус в поиск (U13)"


def test_keyboard_arrows_move_selection_without_toggle(controller):
    ids = [add_party(controller, f"ООО Стрелка-{i}", f"770123456{i}")["id"]
           for i in range(3)]
    flow = controller.registry_flow
    controller.select_page("registry")
    flow.table.setFocus()
    flow.table.set_selection(f"p:{ids[0]}")
    controller.app.processEvents()

    QTest.keyClick(flow.table, Qt.Key.Key_Down)
    controller.app.processEvents()
    assert flow._registry_selected_id == ids[1], "стрелка вниз двигает выделение"
    assert flow.table.selected_iid() == f"p:{ids[1]}"

    QTest.keyClick(flow.table, Qt.Key.Key_Down)
    controller.app.processEvents()
    assert flow._registry_selected_id == ids[2], \
        "стрелки не снимают выделение (без toggle, U13)"

    QTest.keyClick(flow.table, Qt.Key.Key_Up)
    controller.app.processEvents()
    assert flow._registry_selected_id == ids[1]


def test_keyboard_enter_activates_drill(controller):
    pid = add_party(controller, "ООО Энтер", "7701234560")["id"]
    flow = controller.registry_flow
    controller.select_page("registry")
    flow.table.setFocus()
    flow.table.set_selection(f"p:{pid}")
    controller.app.processEvents()

    QTest.keyClick(flow.table, Qt.Key.Key_Return)
    controller.app.processEvents()

    assert flow.is_drill_mode, "Enter открывает документы контрагента (U13)"


def test_keyboard_delete_triggers_delete_selected(controller):
    pid = add_party(controller, "ООО Делит", "7701234560")["id"]
    flow = controller.registry_flow
    controller.select_page("registry")
    flow.table.setFocus()
    flow.table.set_selection(f"p:{pid}")
    controller.app.processEvents()

    QTest.keyClick(flow.table, Qt.Key.Key_Delete)
    controller.app.processEvents()

    assert pid in [r["id"] for r in controller.store.trashed_counterparties()], \
        "клавиша Delete удаляет выделенную запись (через подтверждение)"


def test_click_toggle_clears_and_restores_selection(controller):
    pid = add_party(controller, "ООО Тоггл", "7701234560")["id"]
    flow = controller.registry_flow
    controller.select_page("registry")
    controller.app.processEvents()
    model = flow.table.model
    row_rect = flow.table.visualRect(model.index(0, 0))
    viewport = flow.table.viewport()
    center = row_rect.center()

    QTest.mouseClick(viewport, Qt.MouseButton.LeftButton, pos=center)
    controller.app.processEvents()
    assert flow._registry_selected_id == pid, "клик выделяет запись"

    QTest.mouseClick(viewport, Qt.MouseButton.LeftButton, pos=center)
    controller.app.processEvents()
    assert flow._registry_selected_id is None, \
        "повторный клик снимает выделение (user=True)"


def test_double_click_opens_and_keeps_selection(controller):
    pid = add_party(controller, "ООО Дабл", "7701234560")["id"]
    flow = controller.registry_flow
    controller.select_page("registry")
    controller.app.processEvents()
    row_rect = flow.table.visualRect(flow.table.model.index(0, 0))
    viewport = flow.table.viewport()

    QTest.mouseDClick(viewport, Qt.MouseButton.LeftButton, pos=row_rect.center())
    controller.app.processEvents()

    assert flow.is_drill_mode, "двойной клик проваливается в документы"


# --------------------------------------------------------------------------- #
#  Теги сроков и баннер (2.18.0)
# --------------------------------------------------------------------------- #
def _additional_doc(controller, party_id: str, until_iso: str) -> dict:
    return controller.store.add_document(
        storage.KIND_ADDITIONAL, party_id, "3", "01.10.2026",
        data={}, extra={"valid_until_iso": until_iso}, text="",
    )


def test_expiry_tags_color_rows_by_theme(controller):
    pid = add_party(controller, "ООО Сроки", "7701234560")["id"]
    expired = _additional_doc(controller, pid, "2026-01-01")
    soon = _additional_doc(controller, pid, "2026-12-15")
    flow = controller.registry_flow
    flow.refresh()
    flow.open_party_documents(controller.store.get_counterparty(pid))
    model = flow.table.model
    assert model.order() == (f"d:{expired['id']}", f"d:{soon['id']}")

    from PySide6.QtGui import QBrush
    light_expired = model.data(model.index(0, 0), Qt.ItemDataRole.ForegroundRole)
    light_soon = model.data(model.index(1, 0), Qt.ItemDataRole.ForegroundRole)
    assert isinstance(light_expired, QBrush)
    assert light_expired.color().name() == qt_theme.color("light", "error_text").lower()
    assert light_soon.color().name() == qt_theme.color("light", "warning_text").lower()

    controller.set_theme("dark")
    dark_expired = model.data(model.index(0, 0), Qt.ItemDataRole.ForegroundRole)
    assert dark_expired.color().name() == qt_theme.color("dark", "error_text").lower(), \
        "цвета тегов следуют теме (живое чтение, как Tk-слушатель)"
    controller.set_theme("light")


def test_expiry_banner_lists_expired(controller):
    pid = add_party(controller, "ООО Баннер", "7701234560")["id"]
    _additional_doc(controller, pid, "2026-01-01")
    flow = controller.registry_flow
    flow.refresh()

    banner_texts = [
        flow._banners_layout.itemAt(i).widget().text()
        for i in range(flow._banners_layout.count())
        if flow._banners_layout.itemAt(i).widget() is not None
    ]
    assert any("Просрочено доп. соглашений: 1" in text for text in banner_texts), \
        "баннер сроков над списком (2.18.0)"


# --------------------------------------------------------------------------- #
#  Отложенный поиск (Tk 200 мс)
# --------------------------------------------------------------------------- #
def test_search_debounce_uses_single_shot_timer(controller):
    flow = controller.registry_flow
    assert flow._search_timer.interval() == 200
    assert flow._search_timer.isSingleShot()
    add_party(controller, "ООО Отложка", "7701234560")
    flow.refresh()
    flow.search_edit.setText("отложка")
    # Таймер перезапущен, refresh ещё не выполнен (задача отложена).
    assert flow._search_timer.isActive()
    assert flow._registry_caption.text() == "Контрагенты (1)"

    flow._search_timer.timeout.emit()  # срабатывание таймера
    assert "Найдено по запросу" in flow._registry_caption.text()


def test_apply_filter_cancels_pending_search(controller):
    flow = controller.registry_flow
    add_party(controller, "ООО Отмена", "7701234560")
    flow.refresh()
    flow.search_edit.setText("отмена")
    assert flow._search_timer.isActive()

    flow.apply_filter()
    assert flow._search_timer.isActive() is False, \
        "Enter в поиске применяет фильтр сразу (U11)"


# --------------------------------------------------------------------------- #
#  Реальная база 10 001: порция, достижимая запись № 10 000, поиск № 5 000
# --------------------------------------------------------------------------- #
def test_registry_10k_portion_reach_and_search(qapp, appdata, stub_dialogs):
    """Критерий M5 на реальном наборе: № 10 000 достижима «Показать ещё»,
    подпись «показаны все N», поиск находит № 5 000 независимо от порции."""
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    try:
        for i in range(10_001):
            ctrl.store.add_counterparty(
                "ooo", {"ИНН": f"77010{i % 10}{i:07d}",
                        "Сокращенное название компании": f"ООО Тест-{i}"})
        flow = ctrl.registry_flow
        ctrl.select_page("registry")
        model = flow.table.model

        order = model.order()
        assert len(order) == 10_001, "10 000 строк + служебная «Показать ещё»"
        assert order[-1] == "show-more"
        assert "показаны 10000 из 10001" in flow._registry_caption.text()
        assert order[9_999] == "p:10000", "запись № 10 000 — в первой порции"

        # Поиск находит № 5 000 независимо от порции (цифры запроса —
        # поиск по ИНН: подстрока «0005000» есть только у ИНН записи № 5000).
        flow.search_edit.setText("0005000")
        flow.apply_filter()
        found = flow.table.model.order()
        assert found == ("p:5001",), \
            "поиск по всей базе — не ограничен порцией отображения"
        assert flow._registry_caption.text() == "Найдено по запросу «0005000»: 1"

        flow.reset_filter()
        flow.show_more()
        assert len(model.order()) == 10_001 and "show-more" not in model.order()
        assert "показаны все 10001" in flow._registry_caption.text(), \
            "«Показать ещё» до конца + подпись «показаны все N»"

        # Прокрутка 10k: скроллбар жив, прокрутка в конец не падает.
        flow.search_edit.clear()
        flow.apply_filter()
        qapp.processEvents()
        bar = flow.table.verticalScrollBar()
        assert bar.maximum() > 0
        bar.setValue(bar.maximum())
        qapp.processEvents()
        assert bar.value() == bar.maximum()
    finally:
        ctrl._stop_timers()
        if ctrl.spec_flow.preview_timer is not None:
            ctrl.spec_flow.preview_timer.stop()
        ctrl.window.close()
        ctrl.window.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        qapp.processEvents()


# --------------------------------------------------------------------------- #
#  Кросс-совместимость баз (обязательный тест каждого этапа, M3+)
# --------------------------------------------------------------------------- #
def test_cross_base_qt_registry_ops_read_by_tk_store(controller):
    """База после операций Qt-реестра читается свежим (Tk-путь) DataStore."""
    pid = add_party(controller, "ООО Кросс", "7701234560")["id"]
    doc = controller.store.add_document(storage.KIND_CONTRACT, pid, "5",
                                        "01.10.2026", data={}, extra={}, text="")
    controller.lifecycle_flow.delete_counterparty(
        controller.store.get_counterparty(pid))

    fresh = storage.DataStore()
    assert fresh.schema_version == storage.SQLITE_SCHEMA_VERSION
    assert len(fresh.active_counterparties()) == 0, \
        "удалённый контрагент не в активных"
    assert len(fresh.trashed_counterparties()) == 1
    assert len(fresh.trashed_documents()) == 1
    trashed_rec = controller.store.trashed_counterparties()[0]
    controller.lifecycle_flow.restore_counterparty(trashed_rec)
    fresh2 = storage.DataStore()
    assert len(fresh2.active_counterparties()) == 1, \
        "восстановленный контрагент виден Tk-сборке"


def test_settings_file_unchanged_format(controller):
    """Сортировка пишется в прежний settings.json (формат не меняется)."""
    controller.registry_flow.on_sort_changed("По ИНН")
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        data = json.load(fh)
    assert data.get("registry_sort") == "inn"
    assert isinstance(data, dict) and "theme" not in data, \
        "ключи settings.json — прежний формат (theme появляется только при смене темы)"
