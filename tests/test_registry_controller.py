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
        self._registry_docs_row = 1
        self._registry_docs_frame = None
        self._registry_rendered = False
        self._registry_search_after = None
        self._orphans: list = []
        self.calls: list = []
        self.choice_calls: list = []
        self.choice_result = "skip"
        self.docs_dir = "/tmp/registry-tests-docs"

    def refresh_registry(self):
        self.calls.append("registry")

    def _set_current_party(self, rec):
        self.calls.append(("set_party", rec.get("id")))

    def _update_status(self):
        self.calls.append("status")

    def _refresh_party_combo_values(self):
        self.calls.append("combo")

    def _documents_dir(self, interactive=True):
        return self.docs_dir

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
    return RegistryController(app)


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
#  Фасады ContractApp
# --------------------------------------------------------------------------- #
class RecordingFlow:
    def __init__(self):
        self.calls: list = []

    def build_page(self):
        self.calls.append("build_page")

    def refresh(self):
        self.calls.append("refresh")

    def apply_filter(self):
        self.calls.append("apply_filter")

    def select_party(self, rec):
        self.calls.append(("select_party", rec))

    def scan_consistency(self):
        self.calls.append("scan_consistency")


def test_facades_delegate_to_controller():
    """Зеркало smoke-вызовов: app.refresh_registry/_registry_apply_filter/…"""
    import main

    app = object.__new__(main.ContractApp)
    flow = RecordingFlow()
    app._registry_flow = flow

    app._build_registry_page()
    app.refresh_registry()
    app._registry_apply_filter()
    app._select_registry_party({"id": 1})
    app._scan_documents_consistency()

    assert flow.calls == [
        "build_page", "refresh", "apply_filter",
        ("select_party", {"id": 1}), "scan_consistency",
    ]
