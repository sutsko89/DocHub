"""Юнит-тесты DocumentPageController — седьмой контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): страницы «Договор поставки» и
«Доп. соглашение» — сборка, отложенные предпросмотры (п. 5.6), поиск по
ИНН, подстановка договора-основания (п. 5.1, 4.2), «сумма/срок прописью» —
переехали в document_page_controller.py и проверяются без Tk: контроллер
получает фейковое приложение (duck typing) и настоящее хранилище DataStore,
а smoke-тест продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами:
  * предпросмотр собирается отложенно: новый ввод отменяет ещё не
    сработавшую перерисовку, таймер живёт на приложении
    (``app._preview_contract_after`` / ``app._preview_additional_after``);
  * предпросмотр договора подставляет реквизиты контрагента + номер/дату;
    тип «ip» уходит в печать как ИП; сбой сборки — честное сообщение,
    а не пустой предпросмотр;
  * предпросмотр доп. соглашения берёт данные сохранённого договора-
    основания (копию реквизитов на момент создания) + номер/лимит/срок;
  * поиск по ИНН: найден — выбор подставлен и поле поиска очищено;
    не найден — ошибка с ИНН в тексте (пустой ИНН — «—»), выбор не менялся;
  * договор-основание подставляется автоматически при выборе контрагента
    (п. 5.1) с честной подсказкой (п. 4.2); без договоров — не пустой
    список, а «— нет сохранённых договоров —»;
  * «сумма/срок прописью»: пусто — тире, число — словами, мусор — подсказка;
  * фасады ContractApp с прежними именами делегируют контроллеру.
"""
from __future__ import annotations

import logging
import os
from types import SimpleNamespace

import pytest
from docx import Document

import document_page_controller as dpc
import main
import storage
from document_builder import F

DocumentPageController = dpc.DocumentPageController


# --------------------------------------------------------------------------- #
#  Фейки: корень Tk, переменные, виджеты, приложение
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


class FakeLabel:
    """Минимальный аналог CTkLabel: помнит последний текст configure()."""

    def __init__(self):
        self.text = ""

    def configure(self, **kwargs):
        if "text" in kwargs:
            self.text = kwargs["text"]


class FakeInnerEntry:
    """Внутренний tk.Entry: index() и after_idle() в памяти."""

    def __init__(self):
        self.caret = 7
        self.idle_calls: list = []

    def index(self, _which):
        return self.caret

    def after_idle(self, fn):
        self.idle_calls.append(fn)


class FakeEntry:
    """Обёртка ui.entry: у неё доступен .entry."""

    def __init__(self):
        self.entry = FakeInnerEntry()


class FakeApp:
    """Duck-typed приложение: только то, что трогает DocumentPageController."""

    def __init__(self, store, tmp_path):
        self.root = FakeRoot()
        self.store = store
        self._log = logging.getLogger("test-document-page-controller")
        self.calls: list = []
        self.scale = 1.0

        # Страница «Договор».
        self.contract_preview = FakeText()
        self.contract_party_var = FakeVar("")
        self.contract_number_var = FakeVar("")
        self.contract_day_var = FakeVar("")
        self.contract_month_var = FakeVar("")
        self.contract_search_var = FakeVar("")
        self._preview_contract_after = None

        # Страница «Доп. соглашение».
        self.addl_preview = FakeText()
        self.addl_party_var = FakeVar("")
        self.addl_contract_var = FakeVar("")
        self.addl_contract_combo = FakeCombo()
        self.addl_contract_info_var = FakeVar("")
        self.addl_number_var = FakeVar("")
        self._limit_var = FakeVar("")
        self.limit_entry = FakeEntry()
        self.limit_words_label = FakeLabel()
        self.term_var = FakeVar("")
        self.term_entry = FakeEntry()
        self.words_label = FakeLabel()
        self._preview_additional_after = None

        # Что вернут разборчики комбобоксов (настраивается тестом).
        self.party_for_var: dict | None = None
        self.contract_for_var: dict | None = None

    # -- разборчики значений (остаются на приложении) ---------------------- #
    def _party_from_var(self, var):
        return self.party_for_var

    def _contract_from_var(self, var, party_id):
        return self.contract_for_var

    def _contract_display(self, rec):
        return f"Договор поставки № {rec.get('number')}"

    def _set_text(self, textbox, text):
        textbox.insert("1.0", text)

    # -- выбор контрагента (фасады приложения) ----------------------------- #
    def _set_current_party(self, rec, source=None):
        self.calls.append(("set_current_party", rec.get("id") if rec else None))

    def _on_party_selected(self, key):
        self.calls.append(("on_party_selected", key))

    # -- печать предпросмотра: фасады DocumentSaveController ---------------- #
    def _render_contract_docx(self, data, is_ip):
        self.calls.append(("render_contract_docx", dict(data), is_ip))
        doc = Document()
        doc.add_paragraph(f"Договор для {data.get(F.SHORT_NAME, '?')}, № {data.get(F.NUMBER, '')}")
        return doc

    def _render_additional_docx(self, data, is_ip, number, limit_display, term):
        self.calls.append(
            ("render_additional_docx", dict(data), is_ip, number, limit_display, term))
        doc = Document()
        doc.add_paragraph(f"Доп. соглашение № {number}, лимит {limit_display}, срок {term}")
        return doc


@pytest.fixture()
def app(store, tmp_path):
    a = FakeApp(store, tmp_path)
    flow = DocumentPageController(a)
    a._doc_page_flow = flow
    # Фасады ContractApp, которые контроллер вызывает повторно: тот же
    # маршрут «фасад → контроллер», что и в настоящем приложении.
    a._render_contract_preview = flow.render_contract_preview
    a._render_additional_preview = flow.render_additional_preview
    a._update_limit_words = flow.update_limit_words
    return a


@pytest.fixture()
def flow(app):
    return app._doc_page_flow


@pytest.fixture()
def ui_alert_stub(monkeypatch):
    """Перехват ui.alert: запись заголовков и сообщений."""
    alerts: list = []

    def fake_alert(root, title, message, kind="info"):
        alerts.append((title, message, kind))

    monkeypatch.setattr(dpc.ui, "alert", fake_alert)
    return alerts


def seed_party(store, ooo_fields):
    """Контрагент ООО в реестре."""
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


# --------------------------------------------------------------------------- #
#  Отложенный предпросмотр договора
# --------------------------------------------------------------------------- #
def test_preview_contract_debounces(app, flow):
    app._preview_contract_after = "pending-1"
    flow.preview_contract("evt")

    assert app.root.cancelled == ["pending-1"]
    token, ms, fn = app.root.after_calls[-1]
    assert ms == DocumentPageController.PREVIEW_DELAY_MS == 120
    assert app._preview_contract_after == token
    # Перерисовка планируется через атрибут приложения — маршрут подмен сохранён:
    # колбэк резолвится из app._render_contract_preview на каждый вызов.
    assert fn is app._render_contract_preview


def test_render_contract_preview_without_party(app, flow):
    app.party_for_var = None
    flow.render_contract_preview()

    assert app._preview_contract_after is None
    assert "Выберите контрагента" in app.contract_preview.get("1.0", "end")
    assert app.calls == []  # печать не запускалась


def test_render_contract_preview_collects_fields(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    app.party_for_var = rec
    app.contract_number_var.set("5")
    app.contract_day_var.set("12")
    app.contract_month_var.set("января")
    flow.render_contract_preview()

    render_calls = [c for c in app.calls if c[0] == "render_contract_docx"]
    assert len(render_calls) == 1
    _name, data, is_ip = render_calls[0]
    # Реквизиты контрагента + номер/дата из полей страницы.
    assert data[F.SHORT_NAME] == ooo_fields[F.SHORT_NAME]
    assert data[F.INN] == ooo_fields[F.INN]
    assert data[F.NUMBER] == "5"
    assert data[F.DAY] == "12"
    assert data[F.MONTH] == "января"
    assert is_ip is False
    assert "Договор для" in app.contract_preview.get("1.0", "end")
    assert "№ 5" in app.contract_preview.get("1.0", "end")


def test_render_contract_preview_ip_type(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    rec_ip = dict(rec, type="ip")
    app.party_for_var = rec_ip
    flow.render_contract_preview()

    render_calls = [c for c in app.calls if c[0] == "render_contract_docx"]
    assert render_calls[0][2] is True


def test_render_contract_preview_error_is_honest(app, flow, store, ooo_fields):
    app.party_for_var = seed_party(store, ooo_fields)

    def broken(data, is_ip):
        raise RuntimeError("шаблон недоступен")

    app._render_contract_docx = broken
    flow.render_contract_preview()

    assert "Не удалось собрать предпросмотр" in app.contract_preview.get("1.0", "end")


# --------------------------------------------------------------------------- #
#  Отложенный предпросмотр доп. соглашения
# --------------------------------------------------------------------------- #
def test_preview_additional_debounces(app, flow):
    app._preview_additional_after = "pending-2"
    flow.preview_additional("evt")

    assert app.root.cancelled == ["pending-2"]
    token, ms, _fn = app.root.after_calls[-1]
    assert ms == 120
    assert app._preview_additional_after == token


def test_render_additional_preview_without_party(app, flow):
    app.party_for_var = None
    flow.render_additional_preview()

    assert "Выберите контрагента и договор-основание" in app.addl_preview.get("1.0", "end")
    assert app.calls == []


def test_render_additional_preview_without_contract(app, flow, store, ooo_fields):
    app.party_for_var = seed_party(store, ooo_fields)
    app.contract_for_var = None
    flow.render_additional_preview()

    assert "Выберите контрагента и договор-основание" in app.addl_preview.get("1.0", "end")
    assert app.calls == []


def test_render_additional_preview_uses_contract_data(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    contract = seed_contract(store, rec["id"], ooo_fields)
    app.party_for_var = rec
    # Договор-основание распознан по значению комбобокса.
    app.contract_for_var = contract
    app.addl_number_var.set("1")
    app._limit_var.set("1 500 000")
    app.term_var.set("45")
    flow.render_additional_preview()

    render_calls = [c for c in app.calls if c[0] == "render_additional_docx"]
    assert len(render_calls) == 1
    _name, data, is_ip, number, limit_display, term = render_calls[0]
    # Данные берутся из сохранённого договора (копия на момент создания).
    assert data[F.SHORT_NAME] == ooo_fields[F.SHORT_NAME]
    assert is_ip is False
    assert (number, limit_display, term) == ("1", "1 500 000", "45")
    text = app.addl_preview.get("1.0", "end")
    assert "Доп. соглашение № 1" in text
    assert "лимит 1 500 000" in text


def test_render_additional_preview_error_is_honest(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    app.party_for_var = rec
    app.contract_for_var = seed_contract(store, rec["id"], ooo_fields)

    def broken(data, is_ip, number, limit_display, term):
        raise RuntimeError("шаблон недоступен")

    app._render_additional_docx = broken
    flow.render_additional_preview()

    assert "Не удалось собрать предпросмотр" in app.addl_preview.get("1.0", "end")


# --------------------------------------------------------------------------- #
#  Поиск контрагента по ИНН со страницы договора
# --------------------------------------------------------------------------- #
def test_find_party_by_inn_found(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    app.contract_search_var.set(ooo_fields[F.INN])
    flow.find_party_by_inn()

    assert app.calls == [("set_current_party", rec["id"])]
    assert app.contract_search_var.get() == ""


def test_find_party_by_inn_not_found_alerts(app, flow, store, ooo_fields, ui_alert_stub):
    seed_party(store, ooo_fields)
    app.contract_search_var.set("7707083893")
    flow.find_party_by_inn()

    assert app.calls == []  # выбор не менялся
    assert len(ui_alert_stub) == 1
    title, message, kind = ui_alert_stub[0]
    assert title == "Контрагент не найден"
    assert "7707083893" in message
    assert kind == "error"


def test_find_party_by_inn_empty_shows_dash(app, flow, store, ui_alert_stub):
    app.contract_search_var.set("")
    flow.find_party_by_inn()

    _title, message, _kind = ui_alert_stub[0]
    assert "—" in message


# --------------------------------------------------------------------------- #
#  Договор-основание подставляется при выборе контрагента
# --------------------------------------------------------------------------- #
def test_on_addl_party_changed_substitutes_contract(app, flow, store, ooo_fields):
    rec = seed_party(store, ooo_fields)
    contract = seed_contract(store, rec["id"], ooo_fields)
    app.party_for_var = rec
    flow.on_addl_party_changed("evt")

    display = f"Договор поставки № {ooo_fields[F.NUMBER]}"
    assert app.addl_contract_combo.values == [display]
    assert app.addl_contract_var.get() == display
    assert "подставлен автоматически" in app.addl_contract_info_var.get()


def test_on_addl_party_changed_without_contracts(app, flow, store, ooo_fields):
    app.party_for_var = seed_party(store, ooo_fields)
    flow.on_addl_party_changed("evt")

    assert app.addl_contract_combo.values == ["— нет сохранённых договоров —"]
    assert app.addl_contract_var.get() == ""
    assert "сохраните договор поставки" in app.addl_contract_info_var.get()


def test_on_addl_party_changed_without_party(app, flow):
    app.party_for_var = None
    flow.on_addl_party_changed("evt")

    assert app.addl_contract_combo.values == ["— нет сохранённых договоров —"]
    assert "сохраните договор поставки" in app.addl_contract_info_var.get()


# --------------------------------------------------------------------------- #
#  «Сумма/срок прописью» и формат лимита
# --------------------------------------------------------------------------- #
def test_update_words_empty(app, flow):
    app.term_var.set("")
    flow.update_words()
    assert app.words_label.text == "Срок прописью: —"


def test_update_words_digits(app, flow):
    app.term_var.set("45")
    flow.update_words()
    assert app.words_label.text == "Срок прописью: сорок пять"


def test_update_words_invalid(app, flow):
    app.term_var.set("45 дней")
    flow.update_words()
    assert app.words_label.text == "Срок прописью: введите целое число дней"


def test_format_limit_groups_digits(app, flow):
    app._limit_var.set("1500000")
    flow.format_limit()

    assert app._limit_var.get() == "1 500 000"
    # Курсор восстанавливается после переформатирования (after_idle).
    assert app.limit_entry.entry.idle_calls
    assert app.limit_words_label.text == "Сумма прописью: один миллион пятьсот тысяч рублей"


def test_format_limit_already_formatted_keeps_var(app, flow):
    app._limit_var.set("1 500 000")
    flow.format_limit()

    assert app._limit_var.get() == "1 500 000"
    assert app.limit_words_label.text == "Сумма прописью: один миллион пятьсот тысяч рублей"


def test_update_limit_words_empty(app, flow):
    app._limit_var.set("")
    flow.update_limit_words()
    assert app.limit_words_label.text == "Сумма прописью: —"


def test_update_limit_words_garbage(app, flow):
    app._limit_var.set("abc")
    flow.update_limit_words()
    assert app.limit_words_label.text == "Сумма прописью: —"


# --------------------------------------------------------------------------- #
#  Фасады ContractApp делегируют контроллеру
# --------------------------------------------------------------------------- #
class RecordingFlow:
    def __init__(self):
        self.calls: list = []

    def _record(self, name, *args):
        self.calls.append((name, *args))
        return "recorded"

    def build_contract_page(self):
        return self._record("build_contract_page")

    def preview_contract(self, *args):
        return self._record("preview_contract", *args)

    def render_contract_preview(self):
        return self._record("render_contract_preview")

    def find_party_by_inn(self):
        return self._record("find_party_by_inn")

    def build_additional_page(self):
        return self._record("build_additional_page")

    def preview_additional(self, *args):
        return self._record("preview_additional", *args)

    def render_additional_preview(self):
        return self._record("render_additional_preview")

    def on_addl_party_changed(self, *args):
        return self._record("on_addl_party_changed", *args)

    def update_words(self, *args):
        return self._record("update_words", *args)

    def format_limit(self, *args):
        return self._record("format_limit", *args)

    def update_limit_words(self):
        return self._record("update_limit_words")


def make_facade_app():
    flow = RecordingFlow()
    return SimpleNamespace(_doc_page_flow=flow), flow


def test_facade_contract_page_delegates():
    app, flow = make_facade_app()

    main.ContractApp._build_contract_page(app)
    main.ContractApp._preview_contract(app, "evt")
    main.ContractApp._render_contract_preview(app)
    main.ContractApp._find_party_by_inn(app)

    names = [call[0] for call in flow.calls]
    assert names == [
        "build_contract_page", "preview_contract",
        "render_contract_preview", "find_party_by_inn",
    ]
    assert flow.calls[1][1] == "evt"


def test_facade_additional_page_delegates():
    app, flow = make_facade_app()

    main.ContractApp._build_additional_page(app)
    main.ContractApp._preview_additional(app, "evt")
    main.ContractApp._render_additional_preview(app)
    main.ContractApp._on_addl_party_changed(app, "evt")
    main.ContractApp._update_words(app, "evt")
    main.ContractApp._format_limit(app, "evt")
    main.ContractApp._update_limit_words(app)

    names = [call[0] for call in flow.calls]
    assert names == [
        "build_additional_page", "preview_additional",
        "render_additional_preview", "on_addl_party_changed",
        "update_words", "format_limit", "update_limit_words",
    ]


def test_facade_traces_use_app_methods():
    """Трассы переменных привязаны к фасадам ContractApp (маршрут подмен)."""
    import inspect

    src = inspect.getsource(dpc.DocumentPageController.build_contract_page)
    assert "app._preview_contract" in src
    assert "app._find_party_by_inn" in src
    src_addl = inspect.getsource(dpc.DocumentPageController.build_additional_page)
    assert "app._on_addl_party_changed" in src_addl
    assert "app._preview_additional" in src_addl
    assert "app._update_words" in src_addl
    assert "app._format_limit" in src_addl
