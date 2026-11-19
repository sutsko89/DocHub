"""Юнит-тесты LifecycleController — третий контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): сценарии реестра «экспорт/открытие документа»
и «корзина» переехали в lifecycle_controller.py и проверяются без Tk —
контроллер получает фейковое приложение (duck typing) и настоящее
хранилище DataStore, а smoke-тест продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами (из аудита 2.0.1, п. 2.5/2.6/2.9):
  * удаление документа/контрагента — только после подтверждения, в тексте
    которого тип, номер, контрагент и судьба файла на диске;
  * контрагент с большим числом документов требует повторного подтверждения;
  * восстановление и окончательное удаление честно сообщают об ошибках
    хранилища (store.last_error), а не молчат;
  * «Очистить корзину» при пустой корзине не задаёт лишних вопросов;
  * экспорт: «неопознанные» записи открывают папку, импортированные —
    исходный файл; утерянный оригинал не подменяется молча (диалог п. 2.9);
  * открытие файлов и показ папки идут через фасады приложения, тела которых
    резолвят модульные main.open_file / main.show_in_folder — именно их
    подменяет smoke-тест;
  * фасады ContractApp с прежними именами делегируют контроллеру.
"""
from __future__ import annotations

import logging
from types import SimpleNamespace
import os

import pytest
from docx import Document

import lifecycle_controller as lc
import main
import storage
from render_helpers import OOO_FIELDS

LifecycleController = lc.LifecycleController


# --------------------------------------------------------------------------- #
#  Фейки: приложение и перехват диалогов
# --------------------------------------------------------------------------- #
class FakeApp:
    """Duck-typed приложение: только то, что трогает LifecycleController."""

    def __init__(self, store, tmp_path):
        self.root = object()  # только parent для диалогов, Tk не нужен
        self.store = store
        self._log = logging.getLogger("test-lifecycle-controller")
        self._registry_selected_id = None
        # Фасады оболочки: контроллер обязан вызывать открытие файлов через
        # приложение (smoke-тест подменяет модульные main.open_file/…).
        self.opened: list = []
        # Обновления UI, запрошенные контроллером.
        self.calls: list = []
        # Диалог выбора при утере оригинала (instance-атрибут — так же
        # подменяет smoke-тест).
        self.choice_calls: list = []
        self.choice_result = "cancel"
        self.docs_dir = str(tmp_path / "docs")
        os.makedirs(self.docs_dir, exist_ok=True)
        self.asksave_calls: list = []

    # Соседние контроллеры вызываются напрямую (v2.12.0, фасады удалены):
    # записывающие фейки.
        self._doc_page_flow = SimpleNamespace(
            preview_contract=lambda *a: self.calls.append("preview_contract"),
            preview_additional=lambda *a: self.calls.append("preview_additional"),
        )
        self._doc_context_flow = SimpleNamespace(
            refresh_party_combo_values=lambda: self.calls.append("combo"),
        )
        self._doc_save_flow = SimpleNamespace(
            build_record_docx=lambda rec: Document(),
            party_documents_dir=lambda party_rec, interactive=True: self.docs_dir,
        )
        self._settings_flow = SimpleNamespace(
            documents_dir=lambda interactive=True: self.docs_dir)
        self._registry_flow = SimpleNamespace(
            refresh=lambda: self.calls.append("registry"))

    # -- оболочка приложения (остаётся на ContractApp) ---------------------- #
    def _update_status(self):
        self.calls.append("status")

    def _open_file(self, path):
        self.opened.append(("open", path))
        return True

    def _show_in_folder(self, path):
        self.opened.append(("folder", path))

    def _choice_dialog(self, *args):
        self.choice_calls.append(args)
        return self.choice_result


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

    monkeypatch.setattr(lc.ui, "confirm", fake_confirm)
    monkeypatch.setattr(lc.ui, "alert", fake_alert)
    return state


@pytest.fixture()
def app(store, tmp_path):
    return FakeApp(store, tmp_path)


@pytest.fixture()
def flow(app):
    return LifecycleController(app)


def seed(store, ooo_fields):
    """Контрагент + один договор в реестре."""
    party = store.add_counterparty("ooo", dict(ooo_fields))
    doc = store.add_document(
        storage.KIND_CONTRACT, party["id"], "5", "10.01.2026", dict(ooo_fields),
    )
    return party, doc


# --------------------------------------------------------------------------- #
#  Удаление документа
# --------------------------------------------------------------------------- #
def test_delete_document_cancel_keeps_record(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)
    ui_stub["confirm_results"] = [False]

    flow.delete_document(doc)

    assert not store.trashed_documents()
    assert app.calls == [], "после «Отмена» UI не обновляется"


def test_delete_document_ok_trashes_and_refreshes(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)

    flow.delete_document(doc)

    assert len(store.trashed_documents()) == 1
    assert app.calls == ["registry", "status", "preview_contract", "preview_additional"]


def test_delete_document_confirm_mentions_consequences(
    store, ooo_fields, app, flow, ui_stub,
):
    party, doc = seed(store, ooo_fields)
    ui_stub["confirm_results"] = [False]

    flow.delete_document(doc)

    (args, kwargs) = ui_stub["confirms"][-1]
    title, message = args[0], args[1]
    assert title == "Удалить документ"
    assert "Договор поставки № 5 от 10.01.2026" in message
    assert "Переместить документ в корзину" in message
    assert "Файл на диске останется" in message
    assert "Корзина" in message
    assert kwargs.get("ok_text") == "В корзину"


def test_delete_document_unknown_id_alerts(store, app, flow, ui_stub):
    flow.delete_document({"id": 999_999, "kind": storage.KIND_CONTRACT})

    assert not store.trashed_documents()
    assert ui_stub["alerts"], "неудача хранилища должна быть озвучена"
    assert "Запись не найдена" in str(ui_stub["alerts"][-1])


# --------------------------------------------------------------------------- #
#  Удаление контрагента
# --------------------------------------------------------------------------- #
def test_delete_counterparty_lists_documents(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)
    ui_stub["confirm_results"] = [False]

    flow.delete_counterparty(party)

    (args, _kwargs) = ui_stub["confirms"][-1]
    title, message = args[0], args[1]
    assert title == "Удалить контрагента"
    assert "Договор поставки № 5" in message, "в подтверждении перечень документов"
    assert "карточка и документы (1 шт.)" in message
    assert "Файлы на диске останутся" in message
    assert app.docs_dir in message
    assert not store.trashed_documents()


def test_delete_counterparty_ok_resets_selection_and_refreshes(
    store, ooo_fields, app, flow, ui_stub,
):
    party, _doc = seed(store, ooo_fields)
    app._registry_selected_id = party["id"]

    flow.delete_counterparty(party)

    assert len(store.trashed_counterparties()) == 1
    assert app._registry_selected_id is None, "выбор удалённой карточки снят"
    assert app.calls == ["combo", "preview_contract", "preview_additional",
                         "registry", "status"]


def test_delete_counterparty_selection_kept_for_other_party(
    store, ooo_fields, app, flow, ui_stub,
):
    party, _doc = seed(store, ooo_fields)
    app._registry_selected_id = "другая-карточка"

    flow.delete_counterparty(party)

    assert app._registry_selected_id == "другая-карточка"


def test_delete_counterparty_many_docs_second_confirm(
    store, ooo_fields, app, flow, ui_stub, tmp_path,
):
    party, _doc = seed(store, ooo_fields)
    for i in range(21):
        store.add_document(
            storage.KIND_CONTRACT, party["id"], str(i + 1), "", dict(ooo_fields),
        )
    ui_stub["confirm_results"] = [True, False]  # второе подтверждение — отказ

    flow.delete_counterparty(party)

    assert len(ui_stub["confirms"]) == 2, "много документов → двойное подтверждение"
    assert not store.trashed_counterparties(), "после отказа ничего не удалено"

    ui_stub["confirms"].clear()
    ui_stub["confirm_results"] = [True, True]
    flow.delete_counterparty(party)
    assert len(store.trashed_counterparties()) == 1


# --------------------------------------------------------------------------- #
#  Восстановление и окончательное удаление
# --------------------------------------------------------------------------- #
def test_restore_document_returns_from_trash(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)
    store.trash_document(doc["id"])
    app.calls.clear()

    flow.restore_document(doc)

    assert not store.trashed_documents()
    assert app.calls == ["combo", "preview_contract", "preview_additional",
                         "registry", "status"]


def test_restore_document_unknown_id_alerts(store, app, flow, ui_stub):
    flow.restore_document({"id": 999_999})

    assert ui_stub["alerts"], "неудача восстановления озвучивается"
    assert "Запись не найдена" in str(ui_stub["alerts"][-1])


def test_restore_counterparty_returns_card_and_docs(
    store, ooo_fields, app, flow, ui_stub,
):
    party, doc = seed(store, ooo_fields)
    store.trash_counterparty(party["id"])

    flow.restore_counterparty(party)

    assert not store.trashed_documents()
    assert not store.trashed_counterparties()
    assert app.calls == ["combo", "preview_contract", "preview_additional",
                         "registry", "status"]


def test_restore_counterparty_store_error_alerts(
    store, ooo_fields, app, flow, ui_stub, monkeypatch,
):
    party, _doc = seed(store, ooo_fields)
    store.trash_counterparty(party["id"])

    def broken(_party_id):
        store.last_error = "диск недоступен"
        return 0

    monkeypatch.setattr(store, "restore_counterparty", broken)
    flow.restore_counterparty(party)

    assert ui_stub["alerts"], "ошибка хранилища озвучивается"
    assert "диск недоступен" in str(ui_stub["alerts"][-1])
    assert len(store.trashed_counterparties()) == 1, "запись осталась в корзине"


def test_purge_document_confirm_and_cancel(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)
    store.trash_document(doc["id"])

    ui_stub["confirm_results"] = [False]
    flow.purge_document(doc)
    assert len(store.trashed_documents()) == 1, "после «Отмена» запись в корзине"

    ui_stub["confirm_results"] = [True]
    flow.purge_document(doc)
    assert not store.trashed_documents(), "после подтверждения запись удалена"
    assert app.calls == ["registry", "status"]


def test_purge_document_unknown_id_alerts(store, app, flow, ui_stub):
    flow.purge_document({"id": 999_999})

    assert ui_stub["alerts"]
    assert "Запись не найдена" in str(ui_stub["alerts"][-1])


def test_purge_counterparty_removes_everything(
    store, ooo_fields, app, flow, ui_stub,
):
    party, doc = seed(store, ooo_fields)
    store.trash_counterparty(party["id"])

    flow.purge_counterparty(party)

    assert not store.trashed_counterparties()
    assert not store.trashed_documents()
    assert app.calls == ["combo", "registry", "status"]


def test_purge_counterparty_store_error_alerts(
    store, ooo_fields, app, flow, ui_stub, monkeypatch,
):
    party, _doc = seed(store, ooo_fields)
    store.trash_counterparty(party["id"])

    def broken(_party_id):
        store.last_error = "ошибка записи"
        return 0

    monkeypatch.setattr(store, "purge_counterparty", broken)
    flow.purge_counterparty(party)

    assert ui_stub["alerts"]
    assert "ошибка записи" in str(ui_stub["alerts"][-1])


# --------------------------------------------------------------------------- #
#  Очистка корзины
# --------------------------------------------------------------------------- #
def test_empty_trash_no_dialog_when_empty(store, app, flow, ui_stub):
    flow.empty_trash()

    assert not ui_stub["confirms"], "пустая корзина не требует подтверждения"
    assert app.calls == []


def test_empty_trash_purges_everything(store, ooo_fields, app, flow, ui_stub):
    party, doc = seed(store, ooo_fields)
    store.trash_counterparty(party["id"])

    flow.empty_trash()

    assert store.trash_count() == 0
    assert app.calls == ["combo", "registry", "status"]


def test_empty_trash_store_error_alerts(store, ooo_fields, app, flow, ui_stub,
                                        monkeypatch):
    party, doc = seed(store, ooo_fields)
    store.trash_document(doc["id"])

    def broken():
        store.last_error = "не удалось записать реестр"
        return 0

    monkeypatch.setattr(store, "purge_all_trash", broken)
    flow.empty_trash()

    assert ui_stub["alerts"]
    assert "не удалось записать реестр" in str(ui_stub["alerts"][-1])
    assert store.trash_count() == 1


# --------------------------------------------------------------------------- #
#  Экспорт / открытие документа
# --------------------------------------------------------------------------- #
def test_export_unknown_kind_shows_folder_only(store, app, flow, ui_stub):
    rec = {"kind": storage.KIND_UNKNOWN, "extra": {"source_file": "путь/x.docx"}}

    flow.export_document(rec)

    assert app.opened == [("folder", "путь/x.docx")]


def test_export_imported_opens_original_source(store, ooo_fields, app, flow,
                                               ui_stub, tmp_path):
    source = tmp_path / "оригинал.docx"
    source.write_text("x", encoding="utf-8")
    rec = {
        "kind": storage.KIND_CONTRACT,
        "party_id": "нет",
        "extra": {"imported": True, "source_file": str(source)},
    }

    flow.export_document(rec)

    assert app.opened == [("open", str(source))], "открывается исходный файл"
    assert not ui_stub["confirms"]


def test_export_imported_missing_source_cancel(store, ooo_fields, app, flow,
                                               ui_stub, monkeypatch):
    _party, _doc = seed(store, ooo_fields)
    rec = {
        "kind": storage.KIND_CONTRACT, "party_id": None, "number": "777",
        "extra": {"imported": True, "source_file": "нет_такого.docx"},
    }
    app.choice_result = "cancel"

    def no_dialog(**_kwargs):
        raise AssertionError("диалог сохранения не должен открываться")

    monkeypatch.setattr(lc.filedialog, "asksaveasfilename", no_dialog)
    flow.export_document(rec)

    assert app.choice_calls, "утеря оригинала поднимает диалог п. 2.9"
    assert "Оригинал не найден" in str(app.choice_calls[0])
    assert "Пересобрать из шаблона" in str(app.choice_calls[0])
    assert app.opened == [], "без выбора пользователя суррогат не создаётся"


def test_export_imported_missing_source_folder_choice(
    store, ooo_fields, app, flow, ui_stub,
):
    _party, _doc = seed(store, ooo_fields)
    rec = {
        "kind": storage.KIND_CONTRACT, "party_id": None,
        "extra": {"imported": True, "source_file": "нет_такого.docx"},
    }
    app.choice_result = "folder"

    flow.export_document(rec)

    assert app.opened == [("folder", "нет_такого.docx")]


def test_export_imported_missing_source_rebuild_saves(
    store, ooo_fields, app, flow, ui_stub, tmp_path, monkeypatch,
):
    party, _doc = seed(store, ooo_fields)
    rec = {
        "kind": storage.KIND_CONTRACT, "party_id": party["id"], "number": "777",
        "extra": {"imported": True, "source_file": "нет_такого.docx"},
    }
    app.choice_result = "rebuild"
    target = tmp_path / "пересобран.docx"
    monkeypatch.setattr(
        lc.filedialog, "asksaveasfilename", lambda **_kwargs: str(target),
    )

    flow.export_document(rec)

    assert target.exists(), "документ пересобран и сохранён"
    assert app.opened == [("open", str(target))]
    assert ui_stub["alerts"], "пользователь предупреждён о пересборке"
    assert "собран из шаблона" in str(ui_stub["alerts"][-1])


def test_export_created_document_saves_and_opens(
    store, ooo_fields, app, flow, ui_stub, tmp_path, monkeypatch,
):
    party, doc = seed(store, ooo_fields)
    target = tmp_path / "договор.docx"
    captured = {}

    def fake_dialog(**kwargs):
        captured.update(kwargs)
        return str(target)

    monkeypatch.setattr(lc.filedialog, "asksaveasfilename", fake_dialog)
    flow.export_document(doc)

    assert target.exists()
    assert app.opened == [("open", str(target))]
    assert "Договор поставки №5" in captured.get("initialfile", "")
    assert store.party_name(party["id"]) in captured.get("initialfile", "")
    assert not ui_stub["alerts"], "без сюрпризов предупреждений нет"


def test_export_failure_alerts(store, app, flow, ui_stub, monkeypatch):
    rec = {"kind": storage.KIND_CONTRACT, "party_id": None, "number": "1"}

    def broken(**_kwargs):
        raise OSError("диск переполнен")

    monkeypatch.setattr(lc.filedialog, "asksaveasfilename", broken)
    flow.export_document(rec)

    assert ui_stub["alerts"], "ошибка экспорта озвучивается"
    assert "Не удалось сохранить документ" in str(ui_stub["alerts"][-1])


# --------------------------------------------------------------------------- #
#  Фасады ContractApp
# --------------------------------------------------------------------------- #
def test_facade_open_file_routes_through_main(monkeypatch):
    """Зеркало smoke-проверки: подмена main.open_file видна фасаду приложения."""
    seen = []
    monkeypatch.setattr(main, "open_file", lambda path: seen.append(path) or True)

    result = main.ContractApp._open_file(object(), "документ.docx")

    assert result is True
    assert seen == ["документ.docx"]


def test_facade_show_in_folder_routes_through_main(monkeypatch):
    seen = []
    monkeypatch.setattr(main, "show_in_folder", lambda path: seen.append(path))

    main.ContractApp._show_in_folder(object(), "документ.docx")

    assert seen == ["документ.docx"]


