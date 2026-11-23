"""Юнит-тесты волны 3 этапа 5: U8-полностью + U9 + V5 + UPD-UI (остатки).

Контракты:

* U8-полностью (UX-B): «Заполнить по ИНН» работает в окне редактирования
  контрагента — DadataController адресуется через DadataTarget (поля окна,
  фиксированный режим, собственная строка статуса, диалоги модальны к
  окну правки); страница добавления при этом не затрагивается; окно могло
  закрыться во время запроса — подстановка тихо отказывает (TclError);
* U9 (UX-B): пакетные решения импорта — один диалог сводки вместо серии
  диалогов по каждому файлу без ИНН; «создать по названиям» переиспользует
  существующие карточки и группирует файлы по распознанному названию
  (проверка по названию обязательна, дублей без ИНН не плодим);
* V5 (UX-B): акцентная точка «грязной» формы в шапке — точка на кнопке
  раздела появляется, когда форма страницы отличается от baseline, и
  исчезает после сохранения (сброса baseline);
* UPD-UI (остатки P-07): предпросмотр договора не рендерится на скрытой
  странице — его перерисовывает _select_page.
"""
from __future__ import annotations

import logging
import tkinter as tk
from types import SimpleNamespace

import pytest

import dadata_controller as dc
import document_page_controller as dpc
import import_controller as ic
import storage
import ui_macos as ui
from document_builder import F

from test_dadata_controller import (  # noqa: E402
    FakeEntry,
    FakeRoot,
    FakeVar,
    OOO_CARD,
    RecordingFactory,
    make_entries,
)

DadataController = dc.DadataController


# --------------------------------------------------------------------------- #
#  U8-полностью: DadataTarget и окно правки
# --------------------------------------------------------------------------- #
class FakeTypeVar(FakeVar):
    pass


def make_page_app() -> tuple[DadataController, SimpleNamespace]:
    """Fake-приложение: страница добавления в прежнем контракте."""
    app = SimpleNamespace(
        root=FakeRoot(),
        _cp_type_var=FakeVar("Юр. лицо"),
        _cp_status_var=FakeVar(""),
        _cp_entries={"ooo": {}, "ip": {}},
    )
    flow = DadataController(app)
    return flow, app


def make_edit_target(inn: str = "7701234567") -> tuple[dc.DadataTarget, dict, FakeVar, FakeRoot]:
    """Цель «окно правки»: живой FakeRoot, поля и собственный статус."""
    win = FakeRoot()
    entries = make_entries(inn)
    status = FakeVar("")
    return dc.DadataTarget(win, entries, "ooo", status), entries, status, win


def test_dadata_target_from_app_resolves_page():
    flow, app = make_page_app()
    app._cp_entries["ooo"] = make_entries()
    target = dc.DadataTarget.from_app(app)
    assert target.mode == "ooo"
    assert target.root is app.root
    assert target.status_var is app._cp_status_var
    assert target.entries is app._cp_entries["ooo"]
    assert flow is not None


def test_edit_target_fill_validates_inn_and_alerts_to_window(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_page_app()
    target, _entries, _status, win = make_edit_target(inn="123")  # не 10 цифр
    factory = RecordingFactory()

    flow.fill_from_inn(target=target, thread_factory=factory)

    assert factory.created == [], "воркер не запускается при невалидном ИНН"
    assert alerts and alerts[0][0] is win, "диалог модален к окну правки"
    assert flow.busy is False
    # Страница добавления не тронута: валидация читала ИНН окна правки.
    assert app._cp_status_var.value == ""


def test_edit_target_fill_starts_worker_and_busy_status(monkeypatch):
    monkeypatch.setattr(dc.dadata, "get_api_key", lambda: "test-key")
    flow, _app = make_page_app()
    target, _entries, status, _win = make_edit_target()
    factory = RecordingFactory()

    flow.fill_from_inn(target=target, thread_factory=factory)

    assert flow.busy is True
    assert len(factory.created) == 1
    thread = factory.created[0]
    assert thread.args[0] == "7701234567"      # ИНН окна правки
    assert thread.args[2] == "ooo"             # фиксированный режим цели
    assert thread.args[3] is flow.result_queue
    assert "Загружаю" in status.value, "анимация идёт в строке статуса окна"


def test_edit_target_party_loaded_fills_edit_entries(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, _app = make_page_app()
    target, entries, status, _win = make_edit_target()

    # Подмена рабочего результата: результат приходит из очереди (U8).
    flow._target = target
    flow.party_loaded("ooo", dict(OOO_CARD))

    assert entries[F.FULL_NAME].get() == OOO_CARD["full_name"]
    assert entries[F.INN].get() == OOO_CARD["inn"]
    assert entries[F.BIK].get() == OOO_CARD["bik"]
    assert status.value == "Реквизиты заполнены из ЕГРЮЛ/ЕГРИП"
    assert alerts == [], "успешная подстановка без диалогов"
    assert flow.busy is False and flow._target is None


def test_edit_target_destroyed_mid_request_is_quiet(monkeypatch):
    """Окно правки закрылось, пока запрос летел: тихий отказ (U8)."""
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_page_app()
    win = FakeRoot()

    class DeadEntry:
        def delete(self, *_a):
            raise tk.TclError("window destroyed")

        def insert(self, *_a):
            raise tk.TclError("window destroyed")

        def get(self):
            return ""

    entries = {F.INN: DeadEntry()}
    target = dc.DadataTarget(win, entries, "ooo", FakeVar(""))
    flow._target = target

    # Подстановка в разрушенное окно не роняет контроллер и не показывает
    # диалоги (показывать уже некому).
    flow.party_loaded("ooo", dict(OOO_CARD))
    assert alerts == []
    assert flow.busy is False


def test_page_fill_isolated_from_edit_window(monkeypatch):
    """Подстановка в окно правки не трогает поля страницы добавления."""
    flow, app = make_page_app()
    page_entries = make_entries("9999999999")
    app._cp_entries["ooo"] = page_entries
    target, edit_entries, _status, _win = make_edit_target("7701234567")
    monkeypatch.setattr(dc.dadata, "get_api_key", lambda: "test-key")
    factory = RecordingFactory()

    flow.fill_from_inn(target=target, thread_factory=factory)
    flow.party_loaded("ooo", dict(OOO_CARD))

    # Поля страницы добавления не изменились.
    assert page_entries[F.INN].get() == "9999999999"
    assert page_entries[F.FULL_NAME].get() == ""
    # Поля окна правки заполнены.
    assert edit_entries[F.INN].get() == OOO_CARD["inn"]


# --------------------------------------------------------------------------- #
#  U9: пакетные решения импорта без ИНН
# --------------------------------------------------------------------------- #
class U9Store:
    """Минимум Storage для планировщика U9: поиск по ИНН и названию."""

    def __init__(self):
        self.counterparties: list[dict] = []
        self.created: list[dict] = []
        self.documents: list[dict] = []

    def find_by_inn(self, inn):
        for rec in self.counterparties:
            if str((rec.get("fields") or {}).get(F.INN, "")).strip() == inn:
                return rec
        return None

    def find_by_name(self, display):
        for rec in self.counterparties:
            if storage.display_name(rec) == display:
                return rec
        return None

    def add_counterparty(self, mode, fields, save=True):
        rec = {"id": str(len(self.counterparties) + 1), "type": mode,
               "fields": dict(fields)}
        self.counterparties.append(rec)
        self.created.append(rec)
        return rec

    def documents_for(self, party_id):
        return []

    def add_document(self, kind, party_id, number, date,
                     data=None, extra=None, text="", save=True):
        self.documents.append({
            "kind": kind, "party_id": party_id,
            "number": number, "date": date,
            "data": data or {}, "extra": extra or {}, "text": text,
        })


def no_inn_result(name: str) -> dict:
    return {
        "ok": True, "recognized": False, "name": f"{name}.docx",
        "party": {"type": "ooo", "display": name, "fields": {}, "missing": "ИНН"},
        "kind": "contract", "number": "", "date": "", "copied": "x",
    }


def make_u9_flow(choice_result: str) -> tuple[ic.ImportController, SimpleNamespace, U9Store]:
    store = U9Store()
    app = SimpleNamespace(
        root=None, scale=1.0, store=store,
        _log=__import__("logging").getLogger("test-wave3"),
        choice_result=choice_result,
        current_party_rec=None,
    )
    app._choice_dialog = lambda *a, **kw: app.choice_result
    app.current_party = lambda: None
    app._settings_flow = SimpleNamespace(documents_dir_for_import=lambda: "")
    flow = ic.ImportController(app)
    return flow, app, store


def test_batch_no_inn_groups_and_one_dialog(monkeypatch):
    """20 файлов без ИНН — один диалог сводки (критерий U9)."""
    dialogs: list = []
    flow, _app, _store = make_u9_flow("create_by_name")
    flow.app._choice_dialog = lambda *a, **kw: (
        dialogs.append(a), flow.app.choice_result)[1]
    results = [no_inn_result("ООО Ромашка") for _ in range(10)] + \
        [no_inn_result("ООО Кнопка") for _ in range(10)]

    decisions = flow.plan_batch(results)

    assert len(dialogs) == 1, "сводка — один диалог на весь пакет"
    text = dialogs[0][1]
    assert "ООО Ромашка" in text and "ООО Кнопка" in text
    assert len(decisions) == 20
    kinds = {kind for kind, _c in decisions.values()}
    assert kinds == {"create_by_name"}


def test_batch_create_by_name_reuses_existing_card():
    """Проверка по названию: «создать для всех» не плодит дубли (U9)."""
    flow, _app, store = make_u9_flow("create_by_name")
    existing = {"id": "1", "type": "ooo",
                "fields": {F.SHORT_NAME: "ООО Ромашка", F.INN: ""}}
    store.counterparties.append(existing)
    results = [no_inn_result("ООО Ромашка") for _ in range(3)]

    decisions = flow.plan_batch(results)

    # Решение исполняется при регистрации: все три файла — к существующей
    # карточке с тем же названием, новых карточек нет.
    stats = {"added": 0, "new_parties": 0, "skipped": 0,
             "duplicates": 0, "cancelled": 0}
    attached = []
    for r in results:
        entry = flow.new_entry(r)
        if flow.register_result(r, entry, stats, decisions):
            attached.append(entry)
    assert store.created == [], "новых карточек не создаётся при совпадении названия"
    assert len(attached) == 3 and stats["added"] == 3
    assert all("привязан" in e["detail"] for e in attached)


def test_batch_create_by_name_groups_same_names():
    """Два файла с одинаковым нераспознанным названием — одна карточка."""
    flow, _app, store = make_u9_flow("create_by_name")
    results = [no_inn_result("ООО Новая") for _ in range(2)]

    decisions = flow.plan_batch(results)

    entries = []
    stats = {"added": 0, "new_parties": 0, "skipped": 0,
             "duplicates": 0, "cancelled": 0}
    for r in results:
        entry = flow.new_entry(r)
        entries.append(entry)
        changed = flow.register_result(r, entry, stats, decisions)
        assert changed is True
    # Регистрация создала ОДНУ карточку на оба файла.
    assert len(store.created) == 1
    assert stats["new_parties"] == 1
    assert stats["added"] == 2
    assert all(e["status"] == "added" for e in entries)


def test_batch_attach_current_and_skip_all():
    flow, _app, store = make_u9_flow("attach_current")
    current = {"id": "7", "type": "ooo", "fields": {}}
    flow.current_party = lambda: current
    results = [no_inn_result("ООО Без ИНН") for _ in range(3)]

    decisions = flow.plan_batch(results)
    assert all(kind == "attach" and rec["id"] == "7"
               for kind, rec in decisions.values())

    # skip_all: все файлы уходят в unknown, карточки не создаются.
    flow2, _app2, store2 = make_u9_flow("skip_all")
    moved: list = []
    flow2.move_to_unknown = lambda path: moved.append(path) or path
    results2 = [no_inn_result("ООО Пропуск") for _ in range(2)]
    decisions2 = flow2.plan_batch(results2)
    assert all(d == ("skip",) for d in decisions2.values())
    assert store2.created == []


def test_batch_per_file_falls_back_to_ask_party():
    """Явный выбор «решить по каждому» — прежний поток ask_party."""
    flow, _app, _store = make_u9_flow("per_file")
    asked: list = []
    flow.ask_party = lambda result: asked.append(result) or (None, "skip")
    results = [no_inn_result("ООО Один"), no_inn_result("ООО Два")]

    decisions = flow.plan_batch(results)

    assert decisions == {}, "per_file не предрешает файлы"
    stats = {"added": 0, "new_parties": 0, "skipped": 0,
             "duplicates": 0, "cancelled": 0}
    for r in results:
        entry = flow.new_entry(r)
        flow.register_result(r, entry, stats, decisions)
    assert len(asked) == 2, "файлы решаются по одному — выбор пользователя"


def test_batch_conflicts_recognized_inn_existing_name(monkeypatch):
    """ИНОЙ контур U9: ИНН распознан, но название занято другой карточкой."""
    dialogs: list = []
    flow, _app, store = make_u9_flow("attach_all")
    flow.app._choice_dialog = lambda *a, **kw: (
        dialogs.append(a), flow.app.choice_result)[1]
    existing = {"id": "5", "type": "ooo",
                "fields": {F.SHORT_NAME: "ООО Занято", F.INN: "1111111111"}}
    store.counterparties.append(existing)
    result = {
        "ok": True, "recognized": True, "name": "zanyato.docx",
        "party": {"type": "ooo", "display": "ООО Занято",
                  "fields": {F.INN: "2222222222"}},
        "kind": "contract", "number": "5", "date": "01.02.2026",
        "copied": "y",
    }

    decisions = flow.plan_batch([result])

    assert len(dialogs) == 1, "конфликт названий — сводка, не мгновенный диалог"
    assert decisions[id(result)] == ("attach", existing)


def test_register_result_uses_batch_decision(monkeypatch):
    """register_result исполняет предрешённое «skip» без диалогов."""
    flow, _app, store = make_u9_flow("skip_all")
    moved: list = []
    flow.move_to_unknown = lambda path: moved.append(path) or path
    result = no_inn_result("ООО Отклонено")

    decisions = {id(result): ("skip",)}
    stats = {"added": 0, "new_parties": 0, "skipped": 0,
             "duplicates": 0, "cancelled": 0}
    entry = flow.new_entry(result)
    changed = flow.register_result(result, entry, stats, decisions)

    assert changed is False
    assert stats["skipped"] == 1
    assert moved, "отклонённый файл перенесён в unknown"
    assert store.created == []


# --------------------------------------------------------------------------- #
#  V5: акцентная точка «грязной» формы в шапке
# --------------------------------------------------------------------------- #
def test_topnav_item_dot_toggles():
    """Точка появляется/исчезает на кнопке раздела (V5)."""
    root = tk.Tk()
    try:
        item = ui.TopNavItem(root, "Спецификация", "bank", lambda: None)
        assert item.has_dot() is False
        item.set_dot(True)
        assert item.has_dot() is True
        item.set_dot(True)   # повторная установка — без перерисовки-дубля
        assert item.has_dot() is True
        item.set_dot(False)
        assert item.has_dot() is False
    finally:
        root.destroy()


def test_dirty_dot_contract_via_doc_context_stub():
    """_update_dirty_dots читает dirty_pages() и ставит точки по ключам."""
    class NavStub:
        def __init__(self):
            self.dots: list = []

        def set_dot(self, value):
            self.dots.append(value)

    app = SimpleNamespace(
        nav_items={
            "contract": NavStub(), "additional": NavStub(),
            "spec": NavStub(), "counterparty": NavStub(),
            "registry": NavStub(), "settings": NavStub(),
        },
        _doc_context_flow=SimpleNamespace(
            dirty_pages=lambda: ["spec", "counterparty"],
        ),
    )
    # Метод ContractApp работает с duck-typed состоянием: только nav_items
    # и dirty_pages() потока контекста форм.
    import main as main_mod
    main_mod.ContractApp._update_dirty_dots(app)
    assert app.nav_items["spec"].dots[-1] is True
    assert app.nav_items["counterparty"].dots[-1] is True
    assert app.nav_items["registry"].dots[-1] is False
    assert app.nav_items["settings"].dots[-1] is False


# --------------------------------------------------------------------------- #
#  UPD-UI: drill-вид обновляет только список документов (остатки P-07)
# --------------------------------------------------------------------------- #
class RecordingTree:
    """Записывающий фейк дерева реестра: populate против apply_diff."""

    def __init__(self):
        self.calls: list = []
        self._iids: list = []

    def set_columns(self, columns):
        self.calls.append(("columns", columns))

    def populate(self, rows):
        self.calls.append(("populate", len(rows)))
        self._iids = [row.iid for row in rows]

    def apply_diff(self, rows):
        self.calls.append(("diff", len(rows)))
        self._iids = [row.iid for row in rows]

    def order(self):
        return tuple(self._iids)

    def set_selection(self, iid):
        self.calls.append(("select", iid))

    def focus_tree(self):
        pass


def test_drill_refresh_updates_doc_list_only(store):
    """UPD-UI (P-07): в drill-виде операция с документом обновляет только
    список документов (диф), без полной популяции и смены колонок."""
    import registry_controller as rc
    from test_registry_controller import FakeApp as RegFakeApp

    app = RegFakeApp(store)
    flow = rc.RegistryController(app)
    party = store.add_counterparty(
        "ooo", {"Сокращенное название компании": "ООО Дрилл"})
    store.add_document(storage.KIND_CONTRACT, party["id"], "1", "01.01.2026", {})

    tree = RecordingTree()
    app.registry_list = tree
    app._registry_rendered = True

    flow.open_party_documents(party)
    assert app._registry_drill_party_id == party["id"]
    kinds = [c[0] for c in tree.calls]
    assert "populate" in kinds, "первый вход в drill — полная популяция"

    # Операция с документом (добавление) при открытом drill-виде:
    store.add_document(storage.KIND_SPEC, party["id"], "2", "02.01.2026", {})
    tree.calls.clear()
    flow.refresh()

    kinds = [c[0] for c in tree.calls]
    assert "diff" in kinds, "обновление списка документов идёт дифом (R-04)"
    assert "populate" not in kinds, "полная популяция не выполняется"
    assert "columns" not in kinds, "колонки drill-вида не пересобираются"


# --------------------------------------------------------------------------- #
#  UPD-UI: предпросмотр договора — только на видимой странице
# --------------------------------------------------------------------------- #
def test_contract_preview_skips_hidden_page():
    """P-07: предпросмотр договора не рендерится на скрытой странице."""
    rendered: list = []
    app = SimpleNamespace(
        current_page="registry",
        _preview_contract_after=None,
        contract_party_var=FakeVar("ООО Ромашка"),
        contract_number_var=FakeVar("1"), contract_day_var=FakeVar("1"),
        contract_month_var=FakeVar("января"),
        contract_preview=None,
        _set_text=lambda *a, **kw: rendered.append(a),
        _doc_context_flow=SimpleNamespace(
            party_from_var=lambda var: None,
        ),
    )
    flow = dpc.DocumentPageController(app)
    flow.render_contract_preview()
    assert rendered == [], "скрытая страница не рендерит предпросмотр"
    assert app._preview_contract_after is None


def test_contract_preview_renders_visible_page(monkeypatch):
    rendered: list = []
    app = SimpleNamespace(
        current_page="contract",
        _preview_contract_after=None,
        contract_party_var=FakeVar("ООО Ромашка"),
        contract_number_var=FakeVar("1"), contract_day_var=FakeVar("1"),
        contract_month_var=FakeVar("января"),
        contract_preview=SimpleNamespace(),
        _set_text=lambda *a, **kw: rendered.append(a),
        _doc_context_flow=SimpleNamespace(party_from_var=lambda var: None),
    )
    flow = dpc.DocumentPageController(app)
    flow.render_contract_preview()
    assert rendered, "видимая страница рендерит предпросмотр"
