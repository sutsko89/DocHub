"""Юнит-тесты RegistryController — четвёртый контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): код страницы «Реестр» переехал в
registry_controller.py. Виджетная часть (построение страницы, полная и
частичная перерисовки, секции документов/«unknown»/корзины) покрывается
smoke-тестом под Xvfb; здесь проверяется логика без Tk:

  * поиск: подстрока по названию, цифры — по ИНН, пустой запрос — все;
  * сортировка списка (U2): по добавлению / по алфавиту / по ИНН,
    запоминание в settings.json;
  * значения колонок строки контрагента/документа (U1), теги сроков;
  * выбор контрагента: переключение, сброс, перенос на страницы документов,
    ленивый рендер при неотрисованной странице (п. 5.1, 9 этапа 5);
    выделение с клавиатуры (U13) — без toggle;
  * панель действий (V3): подписи кнопок по типу записи, Delete не
    касается выделенной записи корзины;
  * отложенный поиск: перенос задачи (after_cancel) при вводе;
  * фоновая сверка реестра с папкой (п. 2.4): результат на приложении,
    тихая обработка сбоя, обновление реестра только при открытой странице;
  * диалог разбора файлов вне реестра: нет файлов / «Восстановить»
    (успех и ошибка хранилища) / «В корзину» / «Оставить»;
  * фасады ContractApp с прежними именами делегируют контроллеру.

Живой виджет Treeview (популяция, диф, выделение, «Показать ещё», тема)
покрывается tests/test_registry_tree.py и smoke-тестом.
"""
from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest

import registry_controller as rc
import storage
from document_builder import F

RegistryController = rc.RegistryController


# --------------------------------------------------------------------------- #
#  Фейки
# --------------------------------------------------------------------------- #
class FakeRoot:
    def __init__(self):
        self.after_calls: list = []
        self.cancelled: list = []
        self._n = 0

    def after(self, ms, fn):
        self._n += 1
        self.after_calls.append((ms, fn))
        return f"after-{self._n}"

    def after_cancel(self, token):
        self.cancelled.append(token)


class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeApp:
    """Duck-typed приложение: состояние страницы с прежними именами."""

    def __init__(self, store):
        self.root = FakeRoot()
        self.store = store
        self._log = logging.getLogger("test-registry-controller")
        self.registry_search_var = FakeVar("")
        self.pages: dict = {}
        self._registry_selected_id = None
        self._registry_rows: dict = {}
        # 2.19.0: проваливание в контрагента, выделенный документ, времена кликов.
        self._registry_drill_party_id = None
        self._docs_selected_id = None
        self._docs_rows: dict = {}
        self._registry_last_click_id = None
        self._registry_last_click_ts = 0.0
        self._docs_last_click_id = None
        self._docs_last_click_ts = 0.0
        self._registry_rendered = False
        self._registry_search_after = None
        self._orphans: list = []
        # Этап 4 (TREE-1/UX-REG): сортировка, фильтр типов, порция,
        # сигнатура вида, выделение секций, сохранение настроек.
        self._registry_sort = "added"
        self._registry_sort_var = FakeVar("По добавлению")
        self._settings: dict = {}
        self._docs_type_filter = "all"
        self._docs_type_var = FakeVar("Все типы")
        self._registry_shown = rc.registry_tree.DISPLAY_SOFT_LIMIT
        self._registry_signature = None
        self._registry_selected_extra = None
        self.settings_saved: list = []
        self.calls: list = []
        self.choice_calls: list = []
        self.choice_result = "skip"
        self.docs_dir = "/tmp/registry-tests-docs"
        # Соседние контроллеры вызываются напрямую (v2.12.0, фасады удалены):
        # записывающие фейки.
        self._doc_context_flow = SimpleNamespace(
            set_current_party=lambda rec, source=None: self.calls.append(
                ("set_party", rec.get("id"))),
            refresh_party_combo_values=lambda: self.calls.append("combo"),
        )
        self._settings_flow = SimpleNamespace(
            documents_dir=lambda interactive=True: self.docs_dir)
        self._lifecycle_flow = SimpleNamespace(
            delete_document=lambda rec: self.calls.append(
                ("delete_doc", rec.get("id"))),
            delete_counterparty=lambda rec: self.calls.append(
                ("delete_party", rec.get("id"))),
            export_document=lambda rec: self.calls.append(
                ("open_word", rec.get("id"))),
        )

    def _update_status(self):
        self.calls.append("status")

    def save_setting(self, key: str, value) -> None:
        self.settings_saved.append((key, value))

    def _choice_dialog(self, *args):
        self.choice_calls.append(args)
        return self.choice_result

    def _show_in_folder(self, path):
        self.calls.append(("folder", path))


@pytest.fixture()
def app(store):
    return FakeApp(store)


@pytest.fixture()
def ui_alerts(monkeypatch):
    """Перехват ui.alert: настоящий алерт создаёт Tk-окно."""
    alerts: list = []
    monkeypatch.setattr(rc.ui, "alert", lambda root, *args, **kwargs: alerts.append(args))
    return alerts


@pytest.fixture()
def flow(app):
    controller = RegistryController(app)
    # Перерисовка реестра требует живого Tk-рендера: в юнит-тестах она
    # заменяется записывающей подменой (GUI-путь покрывает smoke-тест).
    controller.refresh = lambda *a, **kw: app.calls.append("registry")
    return controller


def make_party(inn="7701234567", short="ООО Ромашка", ptype="ooo"):
    return {
        "id": f"p-{short}",
        "type": ptype,
        "fields": {F.INN: inn, F.SHORT_NAME: short},
    }


# --------------------------------------------------------------------------- #
#  Поиск (parties)
# --------------------------------------------------------------------------- #
def test_parties_empty_query_returns_all(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7701234567",
                                   "Сокращенное название компании": "ООО Ромашка"})
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО Стрелка"})

    assert len(flow.parties()) == 2


def test_parties_filters_by_name_substring(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7701234567",
                                   "Сокращенное название компании": "ООО Ромашка"})
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО Стрелка"})
    app.registry_search_var.set("ромашка")

    names = [storage.display_name(rec) for rec in flow.parties()]
    assert names == ["ООО Ромашка"]


def test_parties_filters_by_inn_digits(store, app, flow):
    party = store.add_counterparty("ooo", {"ИНН": "7701234567",
                                           "Сокращенное название компании": "ООО Ромашка"})
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО Стрелка"})
    app.registry_search_var.set("01234567")  # подстрока ИНН без ведущих цифр

    found = flow.parties()
    assert [rec["id"] for rec in found] == [party["id"]]


def test_parties_no_match_returns_empty(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7701234567",
                                   "Сокращенное название компании": "ООО Ромашка"})
    app.registry_search_var.set("несуществующий")

    assert flow.parties() == []


# --------------------------------------------------------------------------- #
#  Текст строки контрагента
# --------------------------------------------------------------------------- #
def test_party_row_text_mentions_name_kind_inn_id():
    rec = make_party()

    text = RegistryController.party_row_text(rec, selected=False)

    assert "ООО Ромашка" in text
    assert "Юр. лицо" in text
    assert "ИНН 7701234567" in text
    assert "ID p-ООО Ромашка" in text
    assert not text.startswith("●")


def test_party_row_text_ip_type():
    rec = make_party(inn="321300001234", short="ИП Иванов", ptype="ip")
    rec["fields"] = {F.INN: "321300001234", F.IP_NAME: "Иванов И. И."}

    text = RegistryController.party_row_text(rec, selected=False)

    assert text.startswith("ИП Иванов")
    assert "ИП" in text and "Юр. лицо" not in text


def test_party_row_text_truncates_long_address():
    rec = make_party()
    rec["fields"][F.LEGAL_ADDRESS] = "г. Владивосток, " + "длинная улица, " * 8

    text = RegistryController.party_row_text(rec, selected=False)

    assert len([ln for ln in [text]]) == 1
    assert "…" in text, "длинный адрес усекается"
    assert text.index("…") < len(rec["fields"][F.LEGAL_ADDRESS])


def test_party_row_text_selected_marker_and_missing_inn():
    rec = make_party()
    rec["fields"][F.INN] = ""

    text = RegistryController.party_row_text(rec, selected=True)

    assert text.startswith("●  ")
    assert "ИНН —" in text


# --------------------------------------------------------------------------- #
#  Выбор контрагента (select_party)
# --------------------------------------------------------------------------- #
def test_select_party_selects_and_updates_document_pages(app, flow):
    rec = make_party()

    flow.select_party(rec)

    assert app._registry_selected_id == rec["id"]
    assert ("set_party", rec["id"]) in app.calls, "выбор переносится на страницы"


def test_select_party_second_click_resets(app, flow):
    rec = make_party()
    flow.select_party(rec)

    flow.select_party(rec)

    assert app._registry_selected_id is None


def test_select_party_lazy_render_when_page_not_built(app, flow):
    rec = make_party()

    flow.select_party(rec)  # _registry_rows пуст → полная перерисовка

    assert app.calls.count("registry") == 1, "ленивый рендер вызывает refresh_registry"


# --------------------------------------------------------------------------- #
#  Отложенный поиск
# --------------------------------------------------------------------------- #
def test_on_search_changed_schedules_deferred_refresh(app, flow):
    flow.on_search_changed()

    assert len(app.root.after_calls) == 1
    assert app.root.after_calls[0][0] == 200
    assert app._registry_search_after == "after-1"


def test_apply_filter_cancels_pending_and_refreshes(app, flow):
    flow.on_search_changed()

    flow.apply_filter()

    assert app.root.cancelled == ["after-1"], "незавершённая задача отменяется"
    assert app._registry_search_after is None
    assert app.calls.count("registry") == 1


def test_apply_filter_without_pending(app, flow):
    flow.apply_filter()

    assert app.root.cancelled == []
    assert app.calls.count("registry") == 1


def test_reset_filter_clears_query_and_applies(app, flow):
    app.registry_search_var.set("ромашка")

    flow.reset_filter()

    assert app.registry_search_var.get() == ""
    assert app.calls.count("registry") == 1


# --------------------------------------------------------------------------- #
#  Сверка реестра с папкой документов (п. 2.4)
# --------------------------------------------------------------------------- #
def test_scan_consistency_stores_orphans(app, flow, monkeypatch):
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: ["/x/a.docx"])

    flow.scan_consistency()

    assert app._orphans == ["/x/a.docx"]
    assert app.calls == [], "страница не открыта — реестр не перерисовывается"


def test_scan_consistency_refreshes_when_registry_page_open(app, flow, monkeypatch):
    app.pages = {"registry": object()}
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: [])

    flow.scan_consistency()

    assert app.calls == ["registry"]


def test_scan_consistency_swallows_background_errors(app, flow, monkeypatch):
    def broken(base, store):
        raise OSError("папка недоступна")

    monkeypatch.setattr(rc.reconcile, "find_orphans", broken)

    flow.scan_consistency()  # фоновая задача не должна падать

    assert app._orphans == []


# --------------------------------------------------------------------------- #
#  Диалог разбора файлов вне реестра
# --------------------------------------------------------------------------- #
def test_reconcile_dialog_no_orphans_alerts(app, flow, monkeypatch, ui_alerts):
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: [])

    flow.open_reconcile_dialog()

    assert app.choice_calls == [], "без файлов диалог не открывается"
    assert app.calls == ["registry"]
    assert ui_alerts, "пользователю сообщают, что всё учтено"


def test_reconcile_dialog_restore_success(app, flow, monkeypatch, ui_alerts):
    orphans = ["/x/a.docx"]
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: list(orphans))
    added = []
    monkeypatch.setattr(
        rc.reconcile, "restore_orphans",
        lambda store, base, files, sf: added.append(files) or 3,
    )
    app.choice_result = "restore"

    flow.open_reconcile_dialog()

    assert added == [orphans]
    assert app.calls == ["combo", "registry", "status"]


def test_reconcile_dialog_restore_store_error(app, flow, monkeypatch, ui_alerts):
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: ["/x/a.docx"])
    monkeypatch.setattr(
        rc.reconcile, "restore_orphans",
        lambda store, base, files, sf: setattr(store, "last_error", "ошибка записи") or 0,
    )
    app.choice_result = "restore"

    flow.open_reconcile_dialog()

    assert app.calls == ["combo", "registry", "status"], "обновления выполняются и при ошибке"


def test_reconcile_dialog_trash_choice(app, flow, monkeypatch, ui_alerts):
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: ["/x/a.docx"])
    moved = []
    monkeypatch.setattr(
        rc.reconcile, "move_orphans_to_trash",
        lambda base, files: moved.append(files) or 2,
    )
    app.choice_result = "trash"

    flow.open_reconcile_dialog()

    assert moved == [["/x/a.docx"]]
    assert "combo" not in app.calls and "status" not in app.calls
    assert app.calls.count("registry") == 1


def test_reconcile_dialog_skip_choice_does_nothing(app, flow, monkeypatch):
    monkeypatch.setattr(rc.reconcile, "find_orphans", lambda base, store: ["/x/a.docx"])
    app.choice_result = "skip"

    flow.open_reconcile_dialog()

    assert app.choice_calls, "диалог открывается"
    assert app.calls == [], "без выбора пользователя ничего не происходит"


# --------------------------------------------------------------------------- #
#  Значения колонок, сортировка и панель действий (TREE-1/UX-REG)
# --------------------------------------------------------------------------- #
def test_party_row_values_columns():
    rec = make_party()

    values = RegistryController.party_row_values(rec)

    assert values[0] == "ООО Ромашка"
    assert values[1] == "Юр. лицо"
    assert values[2] == "7701234567"
    assert values[3] == rec["id"]


def test_party_row_values_missing_inn_and_long_address():
    rec = make_party()
    rec["fields"][F.INN] = ""
    rec["fields"][F.LEGAL_ADDRESS] = "г. Владивосток, " + "длинная улица, " * 8

    values = RegistryController.party_row_values(rec)

    assert values[2] == "—"
    assert values[4].endswith("…"), "длинный адрес усечён"


def test_doc_row_values_columns():
    rec = {"id": 1, "kind": storage.KIND_CONTRACT, "number": "5",
           "date": "01.10.2026", "extra": {}}

    doc, date, expiry = RegistryController.doc_row_values(rec)

    assert "№ 5" in doc and "Договор" in doc
    assert date == "01.10.2026"
    assert expiry == ""


def test_doc_row_values_expiry_without_separator():
    rec = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "3",
           "date": "01.10.2026", "extra": {"valid_until_iso": "2026-01-01"}}

    _doc, _date, expiry = RegistryController.doc_row_values(rec)

    assert expiry == "истекло 01.01.2026", "разделитель уходит — срок в своей колонке"


def test_doc_row_tags_marks_expired_and_soon():
    expired = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "1",
               "extra": {"valid_until_iso": "2026-01-01"}}
    soon = {"id": 2, "kind": storage.KIND_ADDITIONAL, "number": "2",
            "extra": {"valid_until_iso": "2026-12-15"}}
    contract = {"id": 3, "kind": storage.KIND_CONTRACT, "number": "5", "extra": {}}

    assert RegistryController.doc_row_tags(expired) == ("expired",)
    assert RegistryController.doc_row_tags(soon) == ("soon",)
    assert RegistryController.doc_row_tags(contract) == ()


def test_action_labels_for_kinds():
    assert rc.action_labels_for("party") == ("Редактировать", "Удалить")
    assert rc.action_labels_for("doc") == ("Открыть Word", "Удалить")
    assert rc.action_labels_for("trash_party") == ("Восстановить", "Удалить навсегда")
    assert rc.action_labels_for("trash_doc") == ("Восстановить", "Удалить навсегда")
    assert rc.action_labels_for("unknown") == ("Показать в папке", "Удалить")
    assert rc.action_labels_for("none") is None


def test_sorted_parties_by_added_keeps_store_order(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7701234567",
                                   "Сокращенное название компании": "ООО Ромашка"})
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО Стрелка"})
    app._registry_sort = "added"

    names = [storage.display_name(rec) for rec in flow.sorted_parties()]

    assert names == ["ООО Ромашка", "ООО Стрелка"]


def test_sorted_parties_by_name(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7701234567",
                                   "Сокращенное название компании": "ООО Стрелка"})
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО Ромашка"})
    app._registry_sort = "name"

    names = [storage.display_name(rec) for rec in flow.sorted_parties()]

    assert names == ["ООО Ромашка", "ООО Стрелка"], "алфавитный порядок не зависит от добавления"


def test_sorted_parties_by_inn_empty_last(store, app, flow):
    store.add_counterparty("ooo", {"ИНН": "7709990000",
                                   "Сокращенное название компании": "ООО С ИНН"})
    store.add_counterparty("ooo", {"Сокращенное название компании": "ООО Без ИНН"})
    app._registry_sort = "inn"

    names = [storage.display_name(rec) for rec in flow.sorted_parties()]

    assert names == ["ООО С ИНН", "ООО Без ИНН"], "пустой ИНН — в конец списка"


def test_on_sort_changed_persists_setting(app, flow):
    app._registry_sort = "added"
    app._registry_sort_var.set("По ИНН")

    flow.on_sort_changed()

    assert app._registry_sort == "inn"
    assert ("registry_sort", "inn") in app.settings_saved


def test_on_sort_changed_same_value_is_noop(app, flow):
    app._registry_sort = "added"
    app._registry_sort_var.set("По добавлению")

    flow.on_sort_changed()

    assert app.settings_saved == [], "перестановка и запись не выполняются"


def test_on_docs_type_changed(app, flow):
    app._docs_type_var.set("Договоры")

    flow.on_docs_type_changed()

    assert app._docs_type_filter == storage.KIND_CONTRACT


def test_show_more_extends_portion(app, flow):
    limit = rc.registry_tree.DISPLAY_SOFT_LIMIT
    app._registry_shown = limit

    flow.show_more()

    assert app._registry_shown == limit * 2, "следующая порция R-05b"


def test_tree_select_sets_and_user_clears(app, flow):
    rec = make_party()
    app._registry_rows[rec["id"]] = {"rec": rec, "iid": f"p:{rec['id']}"}
    app.registry_list = SimpleNamespace(
        row=lambda iid: SimpleNamespace(data={"kind": "party", "rec": rec}),
        set_selection=lambda iid: None,
        focus_tree=lambda: None,
    )

    flow.on_tree_select(f"p:{rec['id']}")
    assert app._registry_selected_id == rec["id"], "клик по строке выделяет контрагента"

    flow.on_tree_select(f"p:{rec['id']}")
    assert app._registry_selected_id == rec["id"], \
        "событие дерева идемпотентно (SET-семантика), эхо не меняет состояние"

    flow.on_tree_select(None)  # эхо популяции/синхронизации
    assert app._registry_selected_id == rec["id"], "эхо не снимает запомненный выбор"

    flow.on_tree_select(None, user=True)  # click-toggle пользователя
    assert app._registry_selected_id is None, "явное действие пользователя снимает выбор"

    flow.on_tree_select(f"p:{rec['id']}")
    flow.on_tree_activate(f"p:{rec['id']}")
    assert app._registry_drill_party_id == rec["id"], \
        "двойной клик/Enter открывает документы контрагента"


def test_tree_activate_doc_opens_word(app, flow):
    doc = {"id": 42, "kind": storage.KIND_CONTRACT, "number": "5",
           "date": "01.10.2026", "extra": {}}
    app._docs_rows[42] = {"rec": doc, "iid": "d:42"}
    app.registry_list = SimpleNamespace(
        row=lambda iid: SimpleNamespace(data={"kind": "doc", "rec": doc}),
        set_selection=lambda iid: None,
        focus_tree=lambda: None,
    )

    flow.on_tree_activate("d:42")

    assert ("open_word", 42) in app.calls, "двойной клик по документу открывает Word"


def test_key_nav_select_does_not_toggle(app, flow):
    rec = make_party()

    flow._select_party_plain(rec)
    assert app._registry_selected_id == rec["id"]
    assert ("set_party", rec["id"]) in app.calls

    flow._select_party_plain(rec)

    assert app._registry_selected_id == rec["id"], "стрелки не снимают выделение"


def test_delete_selected_ignores_trash_selection(app, flow):
    app._registry_selected_extra = {"kind": "trash_party", "rec": {"id": "p-1"}}
    app._registry_selected_id = "p-1"

    assert flow.delete_selected() is False, \
        "Delete не удаляет навсегда запись корзины — только через панель"
    assert app.calls == []


def test_open_party_documents_drills_down_and_back(app, flow):
    rec = make_party()

    flow.open_party_documents(rec)
    assert app._registry_drill_party_id == rec["id"], "проваливание в документы"
    assert app._registry_selected_id == rec["id"], "проваливание выделяет строку"

    flow.back_to_parties()
    assert app._registry_drill_party_id is None, "«Назад» возвращает к контрагентам"
    assert app._docs_selected_id is None, "выделенный документ сброшен"


def test_select_other_party_does_not_break_drill(app, flow):
    rec_a = make_party(short="ООО А")
    rec_b = make_party(short="ООО Б")
    app._registry_drill_party_id = rec_a["id"]

    flow.select_party(rec_b)

    assert app._registry_drill_party_id == rec_a["id"], \
        "выделение другого контрагента не меняет текущий вид"


def test_on_page_open_resets_drill(app, flow):
    rec = make_party()
    flow.open_party_documents(rec)
    assert app._registry_drill_party_id == rec["id"]

    flow.on_page_open()

    assert app._registry_drill_party_id is None, \
        "возврат на страницу реестра показывает список контрагентов"


def test_delete_selected_in_drill_mode_ignores_party(store, app, flow):
    party = store.add_counterparty("ooo", {"Сокращенное название компании": "ООО А"})
    flow.open_party_documents({"id": party["id"]})
    app._docs_selected_id = None

    assert flow.delete_selected() is False, \
        "внутри документов Delete не удаляет контрагента"
    assert not any(call[0] == "delete_party" for call in app.calls)


def test_select_document_toggles(app, flow):
    doc = {"id": 42, "kind": storage.KIND_CONTRACT, "number": "5",
           "date": "01.10.2026", "extra": {}}

    flow.select_document(doc)
    assert app._docs_selected_id == 42

    flow.select_document(doc)
    assert app._docs_selected_id is None, "повторный клик снимает выделение"


def test_delete_selected_prefers_document(store, app, flow):
    doc = store.add_document(storage.KIND_CONTRACT, None, "5", "01.10.2026", {})
    app._docs_selected_id = doc["id"]
    app._registry_selected_id = "p-1"

    flow.delete_selected()

    assert ("delete_doc", doc["id"]) in app.calls
    assert not any(call[0] == "delete_party" for call in app.calls), \
        "документ приоритетнее контрагента"


def test_delete_selected_document_also_in_drill_mode(store, app, flow):
    doc = store.add_document(storage.KIND_CONTRACT, None, "7", "02.10.2026", {})
    flow.open_party_documents({"id": "p-1"})
    app._docs_selected_id = doc["id"]

    assert flow.delete_selected() is True, "в проваливании Delete удаляет выделенный документ"
    assert ("delete_doc", doc["id"]) in app.calls


def test_delete_selected_party_when_no_document(store, app, flow):
    party = store.add_counterparty("ooo", {"Сокращенное название компании": "ООО А"})
    app._docs_selected_id = None
    app._registry_selected_id = party["id"]

    flow.delete_selected()

    assert ("delete_party", party["id"]) in app.calls


def test_delete_selected_without_selection_is_noop(app, flow):
    assert flow.delete_selected() is False
    assert app.calls == []


def test_doc_row_text_includes_expiry_suffix(store, app, flow):
    rec = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "3",
           "date": "01.10.2026",
           "extra": {"valid_until_iso": "2026-01-01"}}
    text = flow.doc_row_text(rec)
    assert "Доп. соглашение № 3" in text
    assert "истекло 01.01.2026" in text, "просрочка видна в строке реестра"


def test_doc_row_color_marks_expired_and_soon(app, flow):
    expired = {"id": 1, "kind": storage.KIND_ADDITIONAL, "number": "1",
               "extra": {"valid_until_iso": "2026-01-01"}}
    soon = {"id": 2, "kind": storage.KIND_ADDITIONAL, "number": "2",
            "extra": {"valid_until_iso": "2026-12-15"}}
    contract = {"id": 3, "kind": storage.KIND_CONTRACT, "number": "5", "extra": {}}

    assert flow.doc_row_color(expired) == rc.ui.COLORS["error_text"]
    assert flow.doc_row_color(soon) == rc.ui.COLORS["warning_text"]
    assert flow.doc_row_color(contract) == rc.ui.COLORS["text"]
