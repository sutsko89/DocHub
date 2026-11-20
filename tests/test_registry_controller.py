"""Юнит-тесты RegistryController — четвёртый контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): код страницы «Реестр» переехал в
registry_controller.py. Виджетная часть (построение страницы, полная и
частичная перерисовки, секции документов/«unknown»/корзины) покрывается
smoke-тестом под Xvfb; здесь проверяется логика без Tk:

  * поиск: подстрока по названию, цифры — по ИНН, пустой запрос — все;
  * текст строки контрагента: тип, ИНН, ID, усечение адреса, маркер выбора;
  * выбор контрагента: переключение, сброс, перенос на страницы документов,
    ленивый рендер при неотрисованной странице (п. 5.1, 9 этапа 5);
  * отложенный поиск: перенос задачи (after_cancel) при вводе и «Найти»;
  * фоновая сверка реестра с папкой (п. 2.4): результат на приложении,
    тихая обработка сбоя, обновление реестра только при открытой странице;
  * диалог разбора файлов вне реестра: нет файлов / «Восстановить»
    (успех и ошибка хранилища) / «В корзину» / «Оставить»;
  * фасады ContractApp с прежними именами делегируют контроллеру.
"""
from __future__ import annotations

import logging
import time
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
#  Взаимодействие 2.19.0: выделение, двойной клик, проваливание, Delete, сроки
# --------------------------------------------------------------------------- #
def test_on_party_click_selects_then_double_click_drills(app, flow):
    rec = make_party()

    flow.on_party_click(rec)
    assert app._registry_selected_id == rec["id"], "первый клик — выделение"
    assert app._registry_drill_party_id is None, "ещё не провалились"

    # Второй клик в пределах окна двойного клика — проваливание в документы.
    flow.on_party_click(rec)
    assert app._registry_selected_id == rec["id"]
    assert app._registry_drill_party_id == rec["id"], "двойной клик открывает документы"


def test_on_party_click_second_click_on_other_party(app, flow):
    rec_a = make_party(short="ООО А")
    rec_b = make_party(short="ООО Б")
    app._registry_last_click_id = rec_a["id"]
    app._registry_last_click_ts = time.monotonic()

    flow.on_party_click(rec_b)

    assert app._registry_selected_id == rec_b["id"]
    assert app._registry_drill_party_id is None, "двойной клик по разным строкам не сработал"


def test_on_party_click_after_timeout_is_single_click(app, flow):
    rec = make_party()
    flow.on_party_click(rec)
    # Имитируем медленный клик: время предыдущего клика — давно.
    app._registry_last_click_ts = time.monotonic() - 10

    flow.on_party_click(rec)

    assert app._registry_selected_id is None, "это снятие выделения, не двойной клик"


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


def test_on_doc_click_double_click_opens_word(app, flow):
    doc = {"id": 42, "kind": storage.KIND_CONTRACT, "number": "5",
           "date": "01.10.2026", "extra": {}}

    flow.on_doc_click(doc)
    assert app._docs_selected_id == 42, "первый клик — выделение"
    assert ("open_word", 42) not in app.calls

    flow.on_doc_click(doc)
    assert ("open_word", 42) in app.calls, "двойной клик открывает документ в Word"


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
