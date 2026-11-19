"""Юнит-тесты SettingsController — восьмой контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): страница «Настройки», ключ DaData и вся
логика папки документов переехали в settings_controller.py и проверяются
без Tk — контроллер получает фейковое приложение (duck typing), настройки
живут в словаре, а ключ DaData — в изолированном APPDATA (фикстура
``appdata`` из conftest). Smoke-тест продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами (из аудита 2.0.1):
  * ключ DaData: пустое поле — ошибка, не «сохранено»; сохранённый ключ
    не подставляется в поле (аудит P2-1), статус без самого значения;
    удаление — с подтверждением; сбой записи — честное сообщение с путём;
  * выбор папки документов сохраняется немедленно (п. 6.2), а не только
    при закрытии; сбой записи — без ложного «Папка сохранена»;
  * папка по умолчанию — видимая «Документы/EW.DocHub» (п. 6.2), при
    недоступных «Документах» — возврат в профиль пользователя;
  * недоступная выбранная папка не подменяется молча (п. 2.8): диалог
    «повторить / выбрать другую / по умолчанию», лимит ретраев, guard от
    повторного диалога;
  * фоновый импорт: без Tk-диалогов — ошибка попадёт в отчёт импорта;
  * path не выходит за пределы папки документов (п. 2.7) — staticmethod,
    как его ждут smoke-тест и ImportController;
  * фасады ContractApp с прежними именами делегируют контроллеру.
"""
from __future__ import annotations

import logging
import os
from types import SimpleNamespace

import pytest

import dadata_client as dadata
import main
import settings_controller as scc
import win_shell
from settings_controller import SettingsController


# --------------------------------------------------------------------------- #
#  Фейки
# --------------------------------------------------------------------------- #
class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeKeyEntry:
    """Поле ввода ключа: get/delete в памяти."""

    def __init__(self, value=""):
        self._value = value

    def get(self):
        return self._value

    def delete(self, _first, _last=None):
        self._value = ""


class FakeApp:
    """Duck-typed приложение: только то, что трогает SettingsController."""

    def __init__(self, tmp_path):
        self.root = object()
        self._log = logging.getLogger("test-settings-controller")
        self.scale = 1.0

        self._settings: dict = {}
        self._documents_dir_var = FakeVar("")
        self._dadata_key_entry = FakeKeyEntry()
        self._dadata_key_status_var = FakeVar("")
        self._docs_dir_fallback_active = False
        self._docs_dir_retry_count = 0
        self._docs_dir_dialog_active = False

        self.calls: list = []
        self.persist_ok = True
        self.choice_result = "default"
        self.choice_calls: list = []
        self.opened: list = []

    # -- диалоговые примитивы и фасады приложения -------------------------- #
    def _choice_dialog(self, title, message, options):
        self.choice_calls.append((title, message, options))
        return self.choice_result

    def _persist_settings(self):
        self.calls.append("persist")
        return self.persist_ok

    def _open_file(self, path):
        self.opened.append(path)
        return True

    def _on_theme_change(self, mode):
        self.calls.append(("theme", mode))


@pytest.fixture()
def app(tmp_path):
    a = FakeApp(tmp_path)
    flow = SettingsController(a)
    a._settings_flow = flow
    return a


@pytest.fixture()
def flow(app):
    return app._settings_flow


@pytest.fixture()
def alerts(monkeypatch):
    """Перехват ui.alert/ui.confirm в модуле контроллера."""
    state = {"alerts": [], "confirm_result": True, "confirms": []}

    def fake_alert(root, title, message, kind="info"):
        state["alerts"].append((title, message, kind))

    def fake_confirm(root, *args, **kwargs):
        state["confirms"].append((args, kwargs))
        return state["confirm_result"]

    monkeypatch.setattr(scc.ui, "alert", fake_alert)
    monkeypatch.setattr(scc.ui, "confirm", fake_confirm)
    return state


@pytest.fixture(autouse=True)
def _isolate_key_storage(appdata):
    """APPDATA изолирован для файла ключа DaData (фикстура из conftest)."""


def broken_dir_path(tmp_path):
    """«Папка», которая на самом деле файл: makedirs на ней падает."""
    path = tmp_path / "not_a_dir"
    path.write_text("x", encoding="utf-8")
    return str(path)


# --------------------------------------------------------------------------- #
#  Ключ DaData
# --------------------------------------------------------------------------- #
def test_dadata_key_status_no_key(flow):
    assert flow.dadata_key_status_text() == "Ключ не задан."


def test_dadata_key_status_saved_without_value(flow):
    assert dadata.save_api_key("test-key-123") is True
    text = flow.dadata_key_status_text()
    assert text.startswith("Ключ сохранён")
    assert "test-key-123" not in text  # значение не показывается
    dadata.clear_api_key()


def test_save_dadata_key_empty_field(flow, alerts):
    app = flow.app
    app._dadata_key_entry = FakeKeyEntry("   ")
    flow.save_dadata_key()

    assert dadata.load_api_key() == ""  # ничего не сохранено
    assert alerts["alerts"] == [("Ключ не указан",
                                 "Введите API-ключ DaData или оставьте поле пустым.",
                                 "error")]


def test_save_dadata_key_success(flow, alerts):
    app = flow.app
    app._dadata_key_entry = FakeKeyEntry("  my-secret-key  ")
    flow.save_dadata_key()

    assert dadata.load_api_key() == "my-secret-key"
    assert app._dadata_key_entry.get() == ""  # поле очищено, ключ не виден
    assert app._dadata_key_status_var.get().startswith("Ключ сохранён")
    title, message, kind = alerts["alerts"][0]
    assert (title, kind) == ("Ключ сохранён", "success")
    assert "my-secret-key" not in message


def test_save_dadata_key_failure_alerts_path(flow, alerts, monkeypatch):
    app = flow.app
    app._dadata_key_entry = FakeKeyEntry("abc")
    monkeypatch.setattr(scc.dadata, "save_api_key", lambda key: False)
    flow.save_dadata_key()

    title, message, kind = alerts["alerts"][0]
    assert (title, kind) == ("Не удалось сохранить ключ", "error")
    assert dadata.key_storage_path() in message
    assert app._dadata_key_entry.get() == "abc"  # значение не потеряно


def test_clear_dadata_key_confirmed(flow, alerts):
    app = flow.app
    dadata.save_api_key("to-remove")
    flow.clear_dadata_key()

    assert dadata.load_api_key() == ""
    assert app._dadata_key_entry.get() == ""
    assert app._dadata_key_status_var.get() == "Ключ не задан."
    assert alerts["alerts"] == [("Ключ удалён",
                                 "Ключ DaData удалён из профиля пользователя.",
                                 "success")]


def test_clear_dadata_key_cancelled(flow, alerts):
    dadata.save_api_key("keep-me")
    alerts["confirm_result"] = False
    flow.clear_dadata_key()

    assert dadata.load_api_key() == "keep-me"  # ключ остался
    assert alerts["alerts"] == []  # подтверждение отменено — тишина
    dadata.clear_api_key()


def test_clear_dadata_key_confirm_message(flow, alerts):
    flow.clear_dadata_key()
    _args, kwargs = alerts["confirms"][0]
    title, message = _args  # (title, message) — позиционные у ui.confirm
    assert title == "Удалить ключ DaData"
    assert "Автозаполнение по ИНН перестанет работать" in message
    assert kwargs.get("ok_text") == "Удалить"


# --------------------------------------------------------------------------- #
#  Выбор папки документов
# --------------------------------------------------------------------------- #
def test_choose_documents_dir_cancelled(app, flow, monkeypatch):
    monkeypatch.setattr(scc.filedialog, "askdirectory", lambda **kw: "")
    flow.choose_documents_dir()

    assert "documents_dir" not in app._settings
    assert app.calls == []  # persist не звался


def test_choose_documents_dir_saves_immediately(app, flow, alerts, monkeypatch, tmp_path):
    target = tmp_path / "docs"
    monkeypatch.setattr(scc.filedialog, "askdirectory", lambda **kw: str(target))
    flow.choose_documents_dir()

    assert app._settings["documents_dir"] == str(target)
    assert app._documents_dir_var.get() == str(target)
    assert app.calls == ["persist"]
    title, message, kind = alerts["alerts"][0]
    assert (title, kind) == ("Папка сохранена", "success")
    assert str(target) in message


def test_choose_documents_dir_persist_failure_no_success_alert(app, flow, alerts, monkeypatch, tmp_path):
    app.persist_ok = False
    monkeypatch.setattr(scc.filedialog, "askdirectory", lambda **kw: str(tmp_path / "x"))
    flow.choose_documents_dir()

    assert app._settings["documents_dir"] == str(tmp_path / "x")
    assert alerts["alerts"] == []  # «Папка сохранена» не обещана


def test_open_folder_button_routes_through_app_facade(app):
    """Кнопка «Открыть папку» идёт через _open_file и папку контроллера."""
    import inspect

    src = inspect.getsource(scc.SettingsController.build_settings_page)
    assert "app._open_file(app._settings_flow.documents_dir())" in src


# --------------------------------------------------------------------------- #
#  Папка документов по умолчанию (п. 6.2)
# --------------------------------------------------------------------------- #
def test_default_documents_dir_creates_visible_folder(app, flow, monkeypatch, tmp_path):
    monkeypatch.setattr(win_shell, "documents_known_folder", lambda: str(tmp_path / "Документы"))
    result = flow.default_documents_dir(interactive=False)

    assert result == str(tmp_path / "Документы" / "EW.DocHub")
    assert os.path.isdir(result)


def test_default_documents_dir_fallback_to_appdata(app, flow, alerts, monkeypatch, tmp_path):
    # «Документы» — файл: создать папку внутри нельзя → возврат в профиль.
    documents_file = tmp_path / "Документы"
    documents_file.write_text("x", encoding="utf-8")
    monkeypatch.setattr(win_shell, "documents_known_folder", lambda: str(documents_file))
    result = flow.default_documents_dir(interactive=False)

    assert result == os.path.join(os.environ["APPDATA"],
                                  main.SETTINGS_DIR_NAME, "documents")
    assert os.path.isdir(result)
    assert alerts["alerts"] == []  # неинтерактивный путь — без диалога


# --------------------------------------------------------------------------- #
#  Недоступная папка: честное решение (п. 2.8)
# --------------------------------------------------------------------------- #
def test_documents_dir_custom_ok_resets_retries(app, flow, tmp_path):
    target = tmp_path / "custom"
    app._settings["documents_dir"] = str(target)
    app._docs_dir_retry_count = 2

    assert flow.documents_dir() == str(target)
    assert os.path.isdir(target)
    assert app._docs_dir_retry_count == 0


def test_documents_dir_no_custom_uses_default(app, flow, monkeypatch, tmp_path):
    monkeypatch.setattr(win_shell, "documents_known_folder", lambda: str(tmp_path / "Документы"))
    assert flow.documents_dir() == str(tmp_path / "Документы" / "EW.DocHub")


def test_documents_dir_broken_interactive_dialog_default(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    result = flow.documents_dir(interactive=True)

    # «Папка по умолчанию» из диалога: путь уведён в fallback.
    assert app._docs_dir_fallback_active is True
    assert result != app._settings["documents_dir"]
    assert os.path.isdir(result)
    assert len(app.choice_calls) == 1
    title, message, options = app.choice_calls[0]
    assert title == "Папка документов недоступна"
    assert "Не удалось открыть папку" in message
    assert app._settings["documents_dir"] in message
    assert [o[0] for o in options] == ["default", "choose", "retry"]


def test_documents_dir_broken_non_interactive_no_dialog(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    result = flow.documents_dir(interactive=False)

    assert app.choice_calls == []  # фоновая задача — без Tk-диалогов
    assert os.path.isdir(result)


def test_resolve_dialog_retry_calls_documents_dir(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app.choice_result = "retry"
    retried = []
    flow.documents_dir = lambda interactive=True: retried.append(interactive) or "RETRY"

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert result == "RETRY"
    assert app._docs_dir_retry_count == 1
    assert retried == [True]


def test_resolve_dialog_choose_persists(app, flow, tmp_path, monkeypatch):
    target = tmp_path / "chosen"
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app._docs_dir_retry_count = 1
    app.choice_result = "choose"
    monkeypatch.setattr(scc.filedialog, "askdirectory", lambda **kw: str(target))

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert result == str(target)
    assert app._settings["documents_dir"] == str(target)
    assert app.calls == ["persist"]
    assert app._docs_dir_retry_count == 0  # счётчик сброшен
    assert app._docs_dir_fallback_active is False


def test_resolve_dialog_choose_cancelled_falls_back(app, flow, tmp_path, monkeypatch):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app.choice_result = "choose"
    monkeypatch.setattr(scc.filedialog, "askdirectory", lambda **kw: "")

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert app._docs_dir_fallback_active is True
    assert os.path.isdir(result)


def test_resolve_retry_limit_exhausted_no_dialog(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app._docs_dir_retry_count = SettingsController.MAX_DIR_RETRIES

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert app.choice_calls == []  # лимит ретраев — диалог не показывается
    assert app._docs_dir_fallback_active is True
    assert os.path.isdir(result)


def test_resolve_guard_fallback_active(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app._docs_dir_fallback_active = True

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert app.choice_calls == []
    assert os.path.isdir(result)


def test_resolve_guard_dialog_active(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)
    app._docs_dir_dialog_active = True

    result = flow.resolve_unavailable_documents_dir(app._settings["documents_dir"],
                                                   OSError("нет доступа"))

    assert app.choice_calls == []
    assert os.path.isdir(result)


# --------------------------------------------------------------------------- #
#  Папка документов для фонового импорта (без Tk)
# --------------------------------------------------------------------------- #
def test_documents_dir_for_import_custom_ok(app, flow, tmp_path):
    target = tmp_path / "import-docs"
    app._settings["documents_dir"] = str(target)

    assert flow.documents_dir_for_import() == str(target)
    assert os.path.isdir(target)


def test_documents_dir_for_import_broken_raises(app, flow, tmp_path):
    app._settings["documents_dir"] = broken_dir_path(tmp_path)

    with pytest.raises(OSError) as exc_info:
        flow.documents_dir_for_import()
    assert "Папка документов недоступна" in str(exc_info.value)
    assert app._settings["documents_dir"] in str(exc_info.value)


def test_documents_dir_for_import_no_custom(app, flow, monkeypatch, tmp_path):
    monkeypatch.setattr(win_shell, "documents_known_folder", lambda: str(tmp_path / "Документы"))
    assert flow.documents_dir_for_import() == str(tmp_path / "Документы" / "EW.DocHub")


# --------------------------------------------------------------------------- #
#  Проверка выхода пути за пределы папки (п. 2.7)
# --------------------------------------------------------------------------- #
def test_is_inside_directory_positive():
    base = os.path.join("base", "docs")
    assert SettingsController.is_inside_directory(base, os.path.join(base, "a.docx"))


def test_is_inside_directory_negative():
    assert not SettingsController.is_inside_directory(
        os.path.join("base", "docs"), os.path.join("base", "other", "x.docx"))


def test_is_inside_directory_empty_path():
    assert not SettingsController.is_inside_directory("base", "")


# --------------------------------------------------------------------------- #
#  is_inside_directory: staticmethod контроллера (маршрут smoke-теста)
# --------------------------------------------------------------------------- #
def test_is_inside_directory_static_route():
    """Smoke зовёт SettingsController.is_inside_directory без экземпляра."""
    assert SettingsController.is_inside_directory("base", os.path.join("base", "f.docx"))
    assert not SettingsController.is_inside_directory("base", os.path.join("out", "f.docx"))
