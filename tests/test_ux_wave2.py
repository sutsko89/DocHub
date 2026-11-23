"""Юнит-тесты волны 2 этапа 5: UPD-UI (P-07) + UX-B частично (U10, U12).

Контракты:

* U12: строка «Позиций: N · Итого: …» под таблицей позиций — сумма из ТОГО
  же document_builder.spec_total(rows), что пишется в реестр и печатную
  форму: расхождение сумм исключено по построению (критерий карточки);
* U10: «Открыть реестр» в отчёте импорта — обновляет реестр, статус и
  открывает страницу реестра;
* UPD-UI (P-07): предпросмотры и информация договора обновляются только
  при видимой странице; на скрытой странице трейсы не трогают виджеты —
  при переключении их перерисовывает _select_page.
"""
from __future__ import annotations

from types import SimpleNamespace

import pytest

import document_save_controller as dsc
import document_page_controller as dpc
import document_context_controller as dcc
import import_controller as ic
import ui_macos as ui
from document_builder import spec_total
from render_helpers import SPEC_ITEMS, spec_total as fixture_spec_total


class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def set(self, value):
        self.value = value

    def get(self):
        return self.value


class FakeSpecTable:
    def __init__(self, rows):
        self._rows = rows

    @property
    def rows(self):
        return self._rows


class FakeApp:
    def __init__(self, current_page="registry"):
        self.current_page = current_page
        self._spec_totals_var = FakeVar()
        self.spec_table = FakeSpecTable([])
        self.calls: list = []

    def _update_status(self):
        self.calls.append("status")

    def _select_page(self, key):
        self.calls.append(f"page:{key}")

    def refresh(self):
        self.calls.append("refresh")


def _make_save_flow(app):
    return dsc.DocumentSaveController(app)


# --------------------------------------------------------------------------- #
#  U12: итог под таблицей позиций
# --------------------------------------------------------------------------- #
def test_spec_totals_line_matches_registry_total():
    app = FakeApp()
    flow = _make_save_flow(app)
    app.spec_table = FakeSpecTable(list(SPEC_ITEMS))
    flow.update_spec_totals()
    expected = fixture_spec_total()
    assert app._spec_totals_var.value == (
        f"Позиций: {len(SPEC_ITEMS)} · Итого: {expected:.2f}".replace(".", ",")
    )
    # Критерий: один spec_total — сумма строки равна сумме реестра/печати.
    assert spec_total(SPEC_ITEMS) == spec_total(app.spec_table.rows)


def test_spec_totals_line_empty_table():
    app = FakeApp()
    flow = _make_save_flow(app)
    flow.update_spec_totals()
    assert app._spec_totals_var.value == "Позиции не добавлены"


def test_on_spec_item_changed_updates_totals(monkeypatch):
    app = FakeApp()
    flow = _make_save_flow(app)
    app.spec_table = FakeSpecTable(list(SPEC_ITEMS))
    monkeypatch.setattr(flow, "regenerate_spec", lambda: None)
    flow.on_spec_item_changed()
    assert "Итого" in app._spec_totals_var.value


# --------------------------------------------------------------------------- #
#  U10: «Открыть реестр» в отчёте импорта
# --------------------------------------------------------------------------- #
def test_show_report_passes_open_registry_callback(monkeypatch):
    captured: dict = {}

    def fake_window(*args, **kwargs):
        captured["on_open_registry"] = kwargs.get("on_open_registry")

    app = SimpleNamespace(root=None, scale=1.0,
                          _registry_flow=SimpleNamespace(refresh=lambda: None),
                          calls=[])
    app._update_status = lambda: app.calls.append("status")
    app._select_page = lambda key: app.calls.append(f"page:{key}")
    app._registry_flow = SimpleNamespace(refresh=lambda: app.calls.append("refresh"))
    monkeypatch.setattr(ic.ui, "import_report_window", fake_window)
    flow = ic.ImportController(app)
    flow.show_report([], "итог")
    assert callable(captured["on_open_registry"])
    captured["on_open_registry"]()
    assert app.calls == ["refresh", "status", "page:registry"]


# --------------------------------------------------------------------------- #
#  UPD-UI (P-07): гейты видимости
# --------------------------------------------------------------------------- #
def test_render_spec_preview_skipped_on_hidden_page():
    app = FakeApp(current_page="registry")
    app._spec_preview_after = "pending"
    flow = _make_save_flow(app)
    # Канарейка: без гейта метод упал бы на отсутствующих атрибутах формы.
    flow.render_spec_preview()
    assert app._spec_preview_after is None  # pending сброшен
    assert not hasattr(app, "spec_text")    # виджет предпросмотра не тронут


def test_render_additional_preview_skipped_on_hidden_page():
    app = FakeApp(current_page="registry")
    app._preview_additional_after = "pending"
    flow = dpc.DocumentPageController(app)
    flow.render_additional_preview()
    assert app._preview_additional_after is None


def test_update_contract_info_skipped_on_hidden_page():
    app = FakeApp(current_page="registry")
    flow = dcc.DocumentContextController(app)
    flow.update_contract_info("spec")  # без spec_contract_var — гейт сработал
    assert not hasattr(app, "spec_contract_var")
