"""Юнит-тесты DocumentSaveController — пятый контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): страница «Спецификация» и сценарии
сохранения договора / доп. соглашения / спецификации переехали в
document_save_controller.py и проверяются без Tk — контроллер получает
фейковое приложение (duck typing) и настоящее хранилище DataStore, а
smoke-тест продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами (из аудита 2.0.1):
  * таблица позиций: фильтр по наименованию, суммы Decimal на лету;
  * договор-основание подставляется автоматически при выборе контрагента
    (п. 5.1), без договоров — честная подсказка, а не пустой список;
  * предпросмотр спецификации собирается отложенно (п. 5.6): таймер живёт
    на приложении (``app._spec_preview_after`` — его сбрасывает smoke-тест),
    «Сохранить» дособирает его синхронно;
  * повтор номера документа у того же контрагента требует подтверждения
    (п. 4.1): в тексте назван тип, номер и дата существующего документа;
  * транзакция «файл + запись реестра» (п. 9.2): папка контрагента
    санитизируется и не выходит за пределы папки документов; сбой записи
    реестра не оставляет «полусохранённого» документа;
  * открытие файла после сохранения идёт через фасад приложения
    ``ContractApp._open_file`` — именно модульное main.open_file подменяет
    smoke-тест;
  * соседние контроллеры вызываются напрямую через app._flow.*
    (фасады ContractApp удалены, v2.12.0).
"""
from __future__ import annotations

import logging
import os
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pytest

import document_save_controller as dsc
import main
import services
import storage
from document_builder import F

DocumentSaveController = dsc.DocumentSaveController


# --------------------------------------------------------------------------- #
#  Фейки: приложение, таблица позиций, текстовый виджет
# --------------------------------------------------------------------------- #
class FakeRoot:
    def __init__(self):
        self.after_calls: list = []
        self.cancelled: list = []
        self._n = 0

    def after(self, ms, fn):
        self._n += 1
        token = f"after-{self._n}"
        self.after_calls.append((token, ms, fn))
        return token

    def after_cancel(self, token):
        self.cancelled.append(token)


class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeText:
    """Минимальный аналог CTkTextbox: insert/delete/get в памяти."""

    def __init__(self):
        self._text = ""

    def get(self, start, end):
        return self._text

    def configure(self, **_kwargs):
        pass

    def delete(self, _start, _end=None):
        self._text = ""

    def insert(self, _start, text):
        self._text = text


class FakeCombo:
    def __init__(self):
        self.values: list | None = None

    def set_values(self, values):
        self.values = list(values)


class FakeSpecTable:
    """Duck-typed SpecTable: строки и позиции задаются тестом."""

    def __init__(self, items=None):
        self.rows: list = []
        self._items = list(items or [])
        self.added: list = []
        self.removed = 0

    def get_items(self):
        return [dict(i) for i in self._items]

    def add_row(self, values):
        self.added.append(dict(values))
        return len(self.added)

    def remove_last(self):
        self.removed += 1


class FakeApp:
    """Duck-typed приложение: только то, что трогает DocumentSaveController."""

    def __init__(self, store, tmp_path):
        self.root = FakeRoot()
        self.store = store
        self._log = logging.getLogger("test-document-save-controller")
        self.calls: list = []
        self.opened: list = []
        self.open_ok = True
        self.docs_dir = str(tmp_path / "docs")
        os.makedirs(self.docs_dir, exist_ok=True)
        self.inside = True

        # Страница «Спецификация».
        self.spec_table = FakeSpecTable()
        self.spec_text = FakeText()
        self.spec_party_var = FakeVar("")
        self.spec_contract_var = FakeVar("")
        self.spec_contract_info_var = FakeVar("")
        self.spec_contract_combo = FakeCombo()
        self.spec_number_var = FakeVar("")
        self.spec_day_var = FakeVar("20")
        self.spec_month_var = FakeVar("января")
        self.spec_prices_until_var = FakeVar("")
        self.spec_delivery_term_var = FakeVar("")
        self.spec_payment_terms_var = FakeVar("")
        self.spec_delivery_basis_var = FakeVar("")
        self._spec_preview_after = None

        # Страница «Договор».
        self.contract_party_var = FakeVar("")
        self.contract_number_var = FakeVar("")
        self.contract_day_var = FakeVar("")
        self.contract_month_var = FakeVar("")

        # Страница «Доп. соглашение».
        self.addl_party_var = FakeVar("")
        self.addl_contract_var = FakeVar("")
        self.addl_number_var = FakeVar("")
        self.addl_valid_var = FakeVar("31.12.2026")
        self._limit_var = FakeVar("")
        self.term_var = FakeVar("")

        # Что вернут разборчики комбобоксов (настраивается тестом).
        self.party_for_var: dict | None = None
        self.contract_for_var: dict | None = None

        # Соседние контроллеры вызываются напрямую (v2.12.0, фасады
        # ContractApp удалены): записывающие фейки.
        self._doc_context_flow = SimpleNamespace(
            party_from_var=lambda var: self.party_for_var,
            contract_from_var=lambda var, party_id: self.contract_for_var,
            contract_display=lambda rec:
                f"Договор поставки № {rec.get('number')}",
            reset_form_baseline=lambda page: self.calls.append(("baseline", page)),
        )
        self._settings_flow = SimpleNamespace(
            documents_dir=lambda interactive=True: self.docs_dir,
            is_inside_directory=lambda base, path: self.inside,
        )
        self._registry_flow = SimpleNamespace(
            refresh=lambda: self.calls.append("registry"))

    # -- страницы и обновления UI ------------------------------------------ #
    def _set_text(self, textbox, text):
        textbox.insert("1.0", text)

    def _clear_page_errors(self, page):
        self.calls.append(("clear_errors", page))

    def _set_page_error(self, page, key, message):
        self.calls.append(("page_error", page, key, message))

    def _focus_page_widget(self, page, key):
        self.calls.append(("focus", page, key))

    def _update_status(self):
        self.calls.append("status")

    def _notify(self, message, path=None):
        self.calls.append(("notify", message, path))

    # -- зависимый фасад приложения ---------------------------------------- #
    def _open_file(self, path):
        self.opened.append(("open", path))
        return self.open_ok


@pytest.fixture()
def ui_stub(monkeypatch):
    """Перехват ui.confirm/ui.alert: возврат по очереди + запись текстов."""
    state = {
        "confirm_results": [],  # очередь ответов; пусто → True
        "confirms": [],
        "alerts": [],
    }

    def fake_confirm(root, *args, **kwargs):
        state["confirms"].append((args, kwargs))
        if state["confirm_results"]:
            return state["confirm_results"].pop(0)
        return True

    def fake_alert(root, *args, **kwargs):
        state["alerts"].append((args, kwargs))

    monkeypatch.setattr(dsc.ui, "confirm", fake_confirm)
    monkeypatch.setattr(dsc.ui, "alert", fake_alert)
    return state


@pytest.fixture()
def app(store, tmp_path):
    return FakeApp(store, tmp_path)


@pytest.fixture()
def flow(app):
    return DocumentSaveController(app)


def seed_party(store, ooo_fields):
    """Контрагент без договорных полей в реестре."""
    return store.add_counterparty(
        "ooo", {k: v for k, v in ooo_fields.items()
                if k not in (F.NUMBER, F.DAY, F.MONTH)}
    )


def seed_contract(store, party_id, ooo_fields):
    """Договор поставки в реестре (для выбора договор-основания)."""
    return store.add_document(
        storage.KIND_CONTRACT, party_id, ooo_fields[F.NUMBER],
        f"{ooo_fields[F.DAY]} {ooo_fields[F.MONTH]}", dict(ooo_fields),
    )


ITEM_BOLT = {"name": "Болт М8", "qty": Decimal("5"), "unit": "шт.",
             "price": Decimal("200")}


# --------------------------------------------------------------------------- #
#  Таблица позиций
# --------------------------------------------------------------------------- #
def test_spec_items_reads_table(app, flow):
    app.spec_table = FakeSpecTable(items=[ITEM_BOLT])
    assert flow.spec_items() == [ITEM_BOLT]


def test_spec_total_sums_decimal(app, flow):
    app.spec_table = FakeSpecTable(
        items=[ITEM_BOLT, {"name": "Гайка", "qty": Decimal("2"),
                           "unit": "шт.", "price": Decimal("100.50")}]
    )
    assert flow.spec_total() == Decimal("1201.00")


def test_spec_rows_property_reads_table(app, flow):
    app.spec_table.rows = [{"name": "x"}]
    assert flow.spec_rows == [{"name": "x"}]


def test_add_spec_row_routes_to_table(app, flow):
    flow.add_spec_row("Болт", "5", "200", "шт.")
    assert app.spec_table.added == [
        {"name": "Болт", "qty": "5", "unit": "шт.", "price": "200"}
    ]


def test_remove_spec_row_routes_to_table(app, flow):
    flow.remove_spec_row()
    assert app.spec_table.removed == 1


# --------------------------------------------------------------------------- #
#  Договор-основание и пересборка предпросмотра (п. 5.1, 5.6)
# --------------------------------------------------------------------------- #
def test_on_spec_party_changed_substitutes_contract(
    store, ooo_fields, app, flow,
):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party

    flow.on_spec_party_changed()

    assert app.spec_contract_combo.values == [
        f"Договор поставки № {ooo_fields[F.NUMBER]}"
    ]
    assert app.spec_contract_var.value == app.spec_contract_combo.values[0]
    assert "подставлен автоматически" in app.spec_contract_info_var.value
    # Пересборка предпросмотра запланирована (п. 5.6).
    assert app._spec_preview_after is not None
    assert app.root.after_calls


def test_on_spec_party_changed_without_contracts(store, ooo_fields, app, flow):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party

    flow.on_spec_party_changed()

    assert app.spec_contract_combo.values == ["— нет сохранённых договоров —"]
    assert app.spec_contract_var.value == ""
    assert "сохраните договор поставки" in app.spec_contract_info_var.value


def test_on_spec_party_changed_without_party(app, flow):
    app.party_for_var = None

    flow.on_spec_party_changed()

    assert app.spec_contract_combo.values == ["— нет сохранённых договоров —"]


def test_regenerate_spec_cancels_pending_token(app, flow):
    app._spec_preview_after = "after-old"

    flow.regenerate_spec()

    assert app.root.cancelled == ["after-old"]
    assert app._spec_preview_after == app.root.after_calls[-1][0]


def test_flush_spec_preview_renders_synchronously(store, ooo_fields, app, flow):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.spec_table = FakeSpecTable(items=[ITEM_BOLT])
    app._spec_preview_after = "after-old"

    flow.flush_spec_preview()

    assert app.root.cancelled == ["after-old"]
    assert app._spec_preview_after is None, "синхронный рендер не оставляет таймер"
    assert "СПЕЦИФИКАЦИЯ" in app.spec_text.get("1.0", "end")


# --------------------------------------------------------------------------- #
#  Предпросмотр спецификации
# --------------------------------------------------------------------------- #
def test_render_preview_without_party_shows_hint(app, flow):
    app.party_for_var = None

    flow.render_spec_preview()

    assert app.spec_text.get("1.0", "end") == \
        "Выберите контрагента и договор-основание слева."
    assert app._spec_preview_after is None


def test_render_preview_without_items(store, ooo_fields, app, flow):
    app.party_for_var = seed_party(store, ooo_fields)

    flow.render_spec_preview()

    assert "Позиции не заполнены" in app.spec_text.get("1.0", "end")


def test_render_preview_with_items(store, ooo_fields, app, flow):
    app.party_for_var = seed_party(store, ooo_fields)
    app.spec_number_var = FakeVar("7")
    app.spec_delivery_term_var = FakeVar("отсрочка 30 дней")
    app.spec_table = FakeSpecTable(items=[ITEM_BOLT])

    flow.render_spec_preview()

    text = app.spec_text.get("1.0", "end")
    assert "СПЕЦИФИКАЦИЯ № 7" in text
    assert "Итого:" in text and "одна тысяча" in text.lower(), "сумма прописью"
    assert "в т.ч. НДС" in text
    assert "отсрочка 30 дней" in text
    # «Цены действительны до» обновляются при каждом рендере.
    assert app.spec_prices_until_var.value
    assert app._spec_preview_after is None


def test_render_preview_mentions_contract(store, ooo_fields, app, flow):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = {
        "id": "c1", "number": ooo_fields[F.NUMBER], "data": dict(ooo_fields),
    }

    flow.render_spec_preview()

    assert f"Приложение № 1 к Договору поставки" in app.spec_text.get("1.0", "end")


# --------------------------------------------------------------------------- #
#  Папка контрагента и транзакция сохранения (п. 9.2)
# --------------------------------------------------------------------------- #
def test_party_documents_dir_without_party(app, flow):
    assert flow.party_documents_dir(None) == app.docs_dir


def test_party_documents_dir_creates_party_folder(store, ooo_fields, app, flow):
    party = seed_party(store, ooo_fields)

    path = flow.party_documents_dir(party)

    assert path.startswith(app.docs_dir)
    assert os.path.isdir(path), "папка контрагента создаётся на диске"
    assert path != app.docs_dir


def test_party_documents_dir_escape_falls_back_to_base(store, ooo_fields, app,
                                                       flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.inside = False

    path = flow.party_documents_dir(party)

    assert path == app.docs_dir
    assert ui_stub["alerts"], "показано предупреждение о недопустимом имени"
    assert "Недопустимое имя папки" in ui_stub["alerts"][0][0]


def test_party_documents_dir_oserror_falls_back_to_base(
    store, ooo_fields, app, flow, ui_stub, monkeypatch,
):
    party = seed_party(store, ooo_fields)

    def broken_makedirs(path, exist_ok=False):
        raise OSError("disk full")

    monkeypatch.setattr(dsc.os, "makedirs", broken_makedirs)

    path = flow.party_documents_dir(party)

    assert path == app.docs_dir
    assert "Не удалось создать папку контрагента" in ui_stub["alerts"][0][0]


def test_persist_document_uses_party_folder(store, ooo_fields, app, flow,
                                            monkeypatch):
    party = seed_party(store, ooo_fields)
    seen = {}

    def fake_save_document(store_, draft, docs_dir):
        seen["store"] = store_
        seen["draft"] = draft
        seen["docs_dir"] = docs_dir
        return {"ok": True, "path": "путь.docx", "record": None, "error": None}

    monkeypatch.setattr(services, "save_document", fake_save_document)

    result = flow.persist_document({"kind": "contract", "party_id": party["id"]})

    assert result["ok"] is True
    assert seen["store"] is store
    assert seen["docs_dir"] == flow.party_documents_dir(party)


def test_persist_document_result_alert_text(app, flow, ui_stub):
    flow.persist_document_result_alert({"error": "диск переполнен"})

    (args, _kwargs) = ui_stub["alerts"][0]
    assert args[0] == "Не удалось сохранить документ"
    assert "диск переполнен" in args[1]
    assert "запись в реестр не добавлена" in args[1]


# --------------------------------------------------------------------------- #
#  Повтор номера документа (п. 4.1)
# --------------------------------------------------------------------------- #
def test_confirm_duplicate_empty_number_silent(app, flow, ui_stub):
    assert flow.confirm_document_duplicate("p1", storage.KIND_CONTRACT, "  ", "12.01") is True
    assert ui_stub["confirms"] == []


def test_confirm_duplicate_same_number_asks(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)
    ui_stub["confirm_results"] = [False]

    allowed = flow.confirm_document_duplicate(
        party["id"], storage.KIND_CONTRACT, ooo_fields[F.NUMBER], "12.01.2026",
    )

    assert allowed is False
    (args, kwargs) = ui_stub["confirms"][0]
    assert args[0] == "Повтор номера"
    assert "уже есть" in args[1]
    assert f"№ {ooo_fields[F.NUMBER]}" in args[1]
    assert kwargs.get("ok_text") == "Создать ещё"


def test_confirm_duplicate_new_number_silent(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)

    allowed = flow.confirm_document_duplicate(
        party["id"], storage.KIND_CONTRACT, "777", "12.01.2026",
    )

    assert allowed is True
    assert ui_stub["confirms"] == []


def test_confirm_duplicate_other_kind_ignored(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)

    allowed = flow.confirm_document_duplicate(
        party["id"], storage.KIND_SPEC, ooo_fields[F.NUMBER], "12.01.2026",
    )

    assert allowed is True
    assert ui_stub["confirms"] == []


# --------------------------------------------------------------------------- #
#  save_contract
# --------------------------------------------------------------------------- #
def test_save_contract_without_party_alerts(app, flow, ui_stub):
    app.party_for_var = None

    flow.save_contract()

    assert ui_stub["alerts"]
    assert "Контрагент не выбран" in ui_stub["alerts"][0][0]
    assert app.calls == []


def test_save_contract_validation_errors(app, flow, ui_stub):
    app.party_for_var = {"id": "p1", "type": "ooo", "fields": {}}
    app.contract_number_var = FakeVar("")
    app.contract_day_var = FakeVar("")

    flow.save_contract()

    assert ("clear_errors", "contract") in app.calls
    assert any(call[0] == "page_error" for call in app.calls)
    assert "Ошибки в данных" in ui_stub["alerts"][0][0]
    assert not app.store.documents


def test_save_contract_duplicate_rejected(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party
    app.contract_number_var = FakeVar(ooo_fields[F.NUMBER])
    app.contract_day_var = FakeVar("12")
    app.contract_month_var = FakeVar("января")
    ui_stub["confirm_results"] = [False]

    flow.save_contract()

    assert len(app.store.documents) == 1, "повтор номера отклонён — новый договор не сохранён"
    assert "registry" not in app.calls


def test_save_contract_success(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_number_var = FakeVar(ooo_fields[F.NUMBER])
    app.contract_day_var = FakeVar(ooo_fields[F.DAY])
    app.contract_month_var = FakeVar(ooo_fields[F.MONTH])

    flow.save_contract()

    assert len(app.store.documents) == 1
    rec = app.store.documents[0]
    assert rec["kind"] == storage.KIND_CONTRACT
    assert rec["extra"]["is_ip"] is False
    # UI: базовая линия сброшена, реестр и статус обновлены, файл открыт.
    assert ("baseline", "contract") in app.calls
    assert "registry" in app.calls and "status" in app.calls
    assert app.opened and app.opened[0][0] == "open"
    # Путь к файлу приходит из транзакции сервисного слоя (п. 9.2) и лежит
    # в индивидуальной папке контрагента внутри папки документов.
    notify_kind, message, path = app.calls[-1]
    assert notify_kind == "notify"
    assert "сохранён" in message
    assert Path(path).exists()
    assert str(path).startswith(app.docs_dir)


def test_save_contract_failed_open_still_saved(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_number_var = FakeVar("42")
    app.contract_day_var = FakeVar("12")
    app.contract_month_var = FakeVar("января")
    app.open_ok = False

    flow.save_contract()

    assert len(app.store.documents) == 1, "неудачное открытие не отменяет сохранение"
    message = app.calls[-1][1]
    assert "сохранён" in message and "и открыт" not in message


# --------------------------------------------------------------------------- #
#  save_additional
# --------------------------------------------------------------------------- #
def test_save_additional_without_contract_alerts(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = None

    flow.save_additional()

    assert "Договор не выбран" in ui_stub["alerts"][0][0]
    assert app.calls == []


def test_save_additional_missing_limit_and_term(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    contract = seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party
    app.contract_for_var = contract
    app._limit_var = FakeVar("")
    app.term_var = FakeVar("")

    flow.save_additional()

    keys = {call[2] for call in app.calls if call[0] == "page_error"}
    assert keys == {F.LIMIT, F.TERM}
    assert "Не хватает данных" in ui_stub["alerts"][0][0]
    assert len(app.store.documents) == 1, "высеванный договор не задвоился"


def test_save_additional_non_digit_term(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    contract = seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party
    app.contract_for_var = contract
    app._limit_var = FakeVar("1 000 000")
    app.term_var = FakeVar("полгода")

    flow.save_additional()

    errors = [call[3] for call in app.calls if call[0] == "page_error"
              and call[2] == F.TERM]
    assert errors == ["Должно быть целое число дней"]


def test_save_additional_success(store, ooo_fields, app, flow, ui_stub, monkeypatch):
    party = seed_party(store, ooo_fields)
    contract = seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party
    app.contract_for_var = contract
    app._limit_var = FakeVar("1 500 000")
    app.term_var = FakeVar("45")
    app.addl_number_var = FakeVar("1")
    seen = {}

    def fake_save_document(store_, draft, docs_dir):
        seen["draft"] = draft
        return {"ok": True, "path": f"{docs_dir}/доп.docx",
                "record": None, "error": None}

    monkeypatch.setattr(services, "save_document", fake_save_document)

    flow.save_additional()

    draft = seen["draft"]
    assert draft["kind"] == storage.KIND_ADDITIONAL
    assert draft["extra"]["contract_id"] == contract["id"]
    assert draft["extra"]["limit_digits"] == "1500000", "лимит без разделителей"
    assert draft["extra"]["term"] == "45"
    assert draft["extra"]["valid_until"] == "31.12.2026", (
        "2.18.0: срок действия сохраняется в записи реестра")
    assert draft["extra"]["valid_until_iso"] == "2026-12-31"
    assert ("baseline", "additional") in app.calls
    assert "сохранено" in app.calls[-1][1]


# --------------------------------------------------------------------------- #
#  save_spec
# --------------------------------------------------------------------------- #
def test_save_spec_missing_number(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = {"id": "c1", "number": "516", "data": dict(ooo_fields)}
    app.spec_number_var = FakeVar("")

    flow.save_spec()

    assert ("page_error", "spec", F.NUMBER, "Заполните поле") in app.calls
    assert ("focus", "spec", F.NUMBER) in app.calls
    assert "Укажите номер спецификации" in ui_stub["alerts"][0][0][1]


def test_save_spec_without_items(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = {"id": "c1", "number": "516", "data": dict(ooo_fields)}
    app.spec_number_var = FakeVar("1")

    flow.save_spec()

    assert "Добавьте хотя бы одну позицию" in ui_stub["alerts"][0][0][1]


def test_save_spec_zero_quantity(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = {"id": "c1", "number": "516", "data": dict(ooo_fields)}
    app.spec_number_var = FakeVar("1")
    app.spec_table = FakeSpecTable(
        items=[{"name": "Болт", "qty": Decimal("0"), "unit": "шт.",
                "price": Decimal("200")}]
    )

    flow.save_spec()

    # U6 (UX-A): сводка ошибок позиций одним алертом.
    assert "количество должно быть больше нуля" in ui_stub["alerts"][0][0][1]


def test_save_spec_zero_price(store, ooo_fields, app, flow, ui_stub):
    party = seed_party(store, ooo_fields)
    app.party_for_var = party
    app.contract_for_var = {"id": "c1", "number": "516", "data": dict(ooo_fields)}
    app.spec_number_var = FakeVar("1")
    app.spec_table = FakeSpecTable(
        items=[{"name": "Болт", "qty": Decimal("5"), "unit": "шт.",
                "price": Decimal("0")}]
    )

    flow.save_spec()

    assert "цена должна быть больше нуля" in ui_stub["alerts"][0][0][1]


def test_save_spec_success(store, ooo_fields, app, flow, ui_stub, monkeypatch):
    party = seed_party(store, ooo_fields)
    contract = seed_contract(store, party["id"], ooo_fields)
    app.party_for_var = party
    app.contract_for_var = contract
    app.spec_number_var = FakeVar("1")
    app.spec_table = FakeSpecTable(items=[ITEM_BOLT])
    app.spec_delivery_term_var = FakeVar("отсрочка 30 дней")
    seen = {}

    def fake_save_document(store_, draft, docs_dir):
        seen["draft"] = draft
        return {"ok": True, "path": f"{docs_dir}/спец.docx",
                "record": None, "error": None}

    monkeypatch.setattr(services, "save_document", fake_save_document)

    flow.save_spec()

    draft = seen["draft"]
    # Текст записи — актуальный предпросмотр, дособранный синхронно.
    assert "СПЕЦИФИКАЦИЯ № 1" in draft["text"]
    assert "Итого:" in draft["text"]
    # Decimal не сериализуется в JSON — кол-во/цены float'ами.
    assert draft["extra"]["items"] == [
        {"name": "Болт М8", "qty": 5.0, "unit": "шт.", "price": 200.0}
    ]
    assert draft["extra"]["total"] == 1000.0
    assert draft["extra"]["contract_id"] == contract["id"]
    assert draft["extra"]["delivery_term"] == "отсрочка 30 дней"
    assert ("baseline", "spec") in app.calls
    assert "сохранена" in app.calls[-1][1]


def test_save_spec_duplicate_rejected(store, ooo_fields, app, flow, ui_stub,
                                      monkeypatch):
    party = seed_party(store, ooo_fields)
    seed_contract(store, party["id"], ooo_fields)
    # Спецификация с тем же номером уже есть — сохранение потребует подтверждения.
    store.add_document(
        storage.KIND_SPEC, party["id"], "1", "20 января",
        dict(ooo_fields), extra={"items": []},
    )
    app.party_for_var = party
    app.contract_for_var = {"id": "c1", "number": ooo_fields[F.NUMBER],
                            "data": dict(ooo_fields)}
    app.spec_number_var = FakeVar("1")
    app.spec_table = FakeSpecTable(items=[ITEM_BOLT])
    ui_stub["confirm_results"] = [False]
    called = []
    monkeypatch.setattr(services, "save_document",
                        lambda *a, **kw: called.append(a) or {"ok": True})

    flow.save_spec()

    assert not called, "после отказа от повтора номера сохранения нет"


# --------------------------------------------------------------------------- #
#  Печатные формы записей (маршрутизация в сервисный слой)
# --------------------------------------------------------------------------- #
def test_save_and_open_document_routes_through_app(store, ooo_fields, app, flow,
                                                   monkeypatch):
    rec = {"kind": storage.KIND_CONTRACT, "party_id": "p1", "number": "5",
           "date": "10.01.2026", "data": {}, "extra": {}, "text": ""}

    def fake_write(doc, rec_, store_, directory):
        return f"{directory}/договор.docx"

    monkeypatch.setattr(services, "write_docx", fake_write)
    monkeypatch.setattr(services, "build_record_docx",
                        lambda rec_, store_: "документ")

    path = flow.save_and_open_document(rec)

    assert path == f"{app.docs_dir}/договор.docx"
    assert app.opened == [("open", path)], "открытие через фасад приложения"


def test_render_docx_helpers_route_to_services(app, flow, monkeypatch):
    monkeypatch.setattr(services, "render_contract",
                        lambda data, is_ip: ("contract", data, is_ip))
    monkeypatch.setattr(
        services, "render_additional",
        lambda data, is_ip, number, limit, term, valid_until="":
        ("additional", data, number),
    )
    monkeypatch.setattr(services, "render_spec",
                        lambda rec, store_, settings: ("spec", rec))
    monkeypatch.setattr(services, "rec_is_ip",
                        lambda rec, store_: True)
    monkeypatch.setattr(services, "build_record_docx",
                        lambda rec, store_: ("record", rec))

    assert flow.render_contract_docx({"x": 1}, True) == ("contract", {"x": 1}, True)
    assert flow.render_additional_docx({}, False, "1", "1 000", "30") == \
        ("additional", {}, "1")
    assert flow.render_additional_docx({}, False, "1", "1 000", "30", "31.12.2026") == \
        ("additional", {}, "1"), "2.18.0: срок действия проходит в печатную форму"
    assert flow.render_spec_docx({"y": 2}) == ("spec", {"y": 2})
    assert flow.rec_is_ip({"z": 3}) is True
    assert flow.build_record_docx({"w": 4}) == ("record", {"w": 4})


def test_render_spec_docx_uses_print_settings(store, ooo_fields, app, flow,
                                              monkeypatch):
    """Печатные настройки берутся из main.get_spec_settings на момент вызова."""
    seen = {}
    monkeypatch.setattr(services, "render_spec",
                        lambda rec, store_, settings: seen.update(settings=settings))

    flow.render_spec_docx({"kind": storage.KIND_SPEC})

    settings = seen["settings"]
    assert settings["spec_year"] > 0, "год вычисляется на момент вызова (аудит 2.0.1)"
    assert "spec_supplier" in settings


# --------------------------------------------------------------------------- #
#  DEF-4: честное сообщение при неудаче отката
# --------------------------------------------------------------------------- #
def test_persist_document_result_alert_orphan_names_file(app, flow, ui_stub):
    flow.persist_document_result_alert({
        "error": "Ошибка записи реестра.",
        "orphan_path": "/docs/ООО Ромашка/Договор №516.docx",
    })
    (args, _kwargs) = ui_stub["alerts"][0]
    assert "Договор №516.docx" in args[1], "имя остаточного файла названо"
    assert "файл удалён" not in args[1].lower(), "ложного «файл удалён» нет"


def test_persist_document_result_alert_removed_states_removal(app, flow, ui_stub):
    flow.persist_document_result_alert({
        "error": "Ошибка записи реестра.\n\nСозданный файл удалён, "
                 "запись в реестр не добавлена.",
    })
    (args, _kwargs) = ui_stub["alerts"][0]
    assert "запись в реестр не добавлена" in args[1]
