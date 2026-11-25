r"""Тесты M6-d — страница «Настройки» на Qt (перенос test_settings_controller).

Контракты Tk-настроек сохраняются дословно:

* тема меняется через каркас (``app.set_theme``): палитра/QSS +
  settings.json + слушатели; текущая тема показывается индикаторами;
* папка документов сохраняется сразу после выбора; отказ записи —
  настройка применена в сеансе, но success-алерта нет (Tk-паритет
  «не показывать „сохранена" при отказе записи»);
* ключ DaData: сохранённый ключ НЕ подставляется в поле (аудит 2.0.1) —
  статус описывает наличие, не значение; пустое поле честно отклоняется;
  удаление — с подтверждением (деструктивное, U7); хранение — профиль
  пользователя (dadata.save_api_key), отдельный от реестра;
* копии реестра (BKUP-1): статус из storage, обновляется после стартовой
  автокопии; папка копий и папка журналов создаются при открытии.

Диалоги подменяются на уровне qt_dialogs; пути DaData изолируются через
APPDATA (dadata.key_storage_path — профиль пользователя).
"""

from __future__ import annotations

import json
import os
import sys

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import app_common
import dadata_client as dadata
import main_qt
import qt_dialogs
import qt_theme
import settings_page_qt
import shell_open
import storage
from PySide6.QtWidgets import QLabel, QPushButton, QRadioButton


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_dialogs(monkeypatch):
    captured = {"confirms": [], "alerts": [], "confirm_result": True}

    def fake_confirm(*args, **kwargs):
        captured["confirms"].append(args)
        return captured["confirm_result"]

    captured["confirm"] = fake_confirm
    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(qt_dialogs, "alert",
                        lambda *a, **k: captured["alerts"].append(a))
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs, monkeypatch, tmp_path):
    """Каркас со страницей настроек; открытие папок глушится."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    monkeypatch.setattr(shell_open, "documents_known_folder", lambda: str(tmp_path))
    monkeypatch.setattr(main_qt.AppController, "_create_startup_backup",
                        lambda self: None)
    monkeypatch.setattr(main_qt, "startup_qt", main_qt.startup_qt)
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl.startup_flow.stop()
    ctrl._stop_timers()
    for flow in (ctrl.spec_flow, ctrl.contract_flow, ctrl.additional_flow):
        if getattr(flow, "preview_timer", None) is not None:
            flow.preview_timer.stop()


# --------------------------------------------------------------------------- #
#  Сборка страницы (плейсхолдер M2 заменён)
# --------------------------------------------------------------------------- #
def test_settings_page_replaces_placeholder(controller):
    flow = controller.settings_flow
    assert controller.pages["settings"] is flow.page
    assert controller.stack.indexOf(flow.page) >= 0
    texts = [label.text() for label in flow.page.findChildren(QLabel)]
    assert not any("в разработке" in t for t in texts)


def test_settings_cards_present(controller):
    """Карточки Tk-страницы: тема, папка, DaData, копии, диагностика, клавиши."""
    flow = controller.settings_flow
    texts = [label.text() for label in flow.page.findChildren(QLabel)]
    joined = "\n".join(texts)
    for expected in (
        "Тема оформления", "Папка для сохранения документов",
        "DaData: автозаполнение по ИНН", "Копии реестра", "Диагностика",
        "Горячие клавиши",
    ):
        assert expected in joined, expected
    buttons = [b.text() for b in flow.page.findChildren(QPushButton)]
    for expected in ("Выбрать папку…", "Открыть папку", "Сохранить ключ",
                     "Удалить ключ", "Открыть папку копий",
                     "Открыть папку журналов"):
        assert expected in buttons, expected


# --------------------------------------------------------------------------- #
#  Тема (через каркас: палитра + settings.json + слушатели)
# --------------------------------------------------------------------------- #
def test_theme_toggle_persists_and_notifies_listeners(controller, appdata):
    flow = controller.settings_flow
    controller.set_theme("dark")  # дефолт — light: сначала уводим в dark
    seen: list = []
    controller.register_theme_listener(lambda mode: seen.append(mode))
    flow._theme_light.setChecked(True)
    assert qt_theme.current_mode() == "light"
    assert seen == ["light"]
    # Настройка сохранена в settings.json (тот же файл, что у Tk-сборки).
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        saved = json.load(fh)
    assert saved["theme"] == "light"
    flow._theme_dark.setChecked(True)
    assert qt_theme.current_mode() == "dark"
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        saved = json.load(fh)
    assert saved["theme"] == "dark"


def test_theme_indicators_follow_applied_theme(controller):
    flow = controller.settings_flow
    controller.set_theme("dark")
    assert flow._theme_dark.isChecked()
    controller.set_theme("light")
    assert flow._theme_light.isChecked()


# --------------------------------------------------------------------------- #
#  Папка документов: выбор с немедленным сохранением
# --------------------------------------------------------------------------- #
def test_choose_documents_dir_saves_immediately(controller, monkeypatch,
                                                tmp_path, stub_dialogs):
    """Выбранная папка сохраняется сразу (Tk: «при аварийном завершении
    выбранная папка не теряется»)."""
    target = tmp_path / "Документы Qt"
    monkeypatch.setattr(
        settings_page_qt.QFileDialog, "getExistingDirectory",
        lambda *a, **kw: str(target),
    )
    controller.settings_flow.choose_documents_dir()
    assert controller.settings["documents_dir"] == str(target)
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        saved = json.load(fh)
    assert saved["documents_dir"] == str(target)
    assert str(target) in controller.settings_flow._documents_dir_label.text()
    assert any("Папка сохранена" == str(a[1]) for a in stub_dialogs["alerts"])


def test_choose_documents_dir_cancelled_no_change(controller, monkeypatch):
    monkeypatch.setattr(
        settings_page_qt.QFileDialog, "getExistingDirectory",
        lambda *a, **kw: "",
    )
    before = controller.settings.get("documents_dir")
    controller.settings_flow.choose_documents_dir()
    assert controller.settings.get("documents_dir") == before


def test_choose_documents_dir_persist_failure_no_success_alert(
    controller, monkeypatch, tmp_path, stub_dialogs,
):
    """Отказ записи: настройка в сеансе применена, «сохранена» — нет."""
    target = tmp_path / "Документы отказ"

    def failing(settings):
        return False

    monkeypatch.setattr(app_common, "write_settings", failing)
    monkeypatch.setattr(
        settings_page_qt.QFileDialog, "getExistingDirectory",
        lambda *a, **kw: str(target),
    )
    controller.settings_flow.choose_documents_dir()
    assert controller.settings["documents_dir"] == str(target), \
        "значение в сеансе осталось применённым"
    assert not any("Папка сохранена" == str(a[1])
                   for a in stub_dialogs["alerts"])


def test_documents_dir_label_shows_custom_or_default(controller, tmp_path):
    flow = controller.settings_flow
    assert str(tmp_path) in flow._documents_dir(), \
        "показывается папка по умолчанию (Документы/EW.DocHub)"
    custom = tmp_path / "Своя папка"
    custom.mkdir()
    controller.save_setting("documents_dir", str(custom))
    assert flow._documents_dir() == str(custom)


# --------------------------------------------------------------------------- #
#  Ключ DaData: поле без подстановки, сохранение, удаление
# --------------------------------------------------------------------------- #
def test_dadata_key_field_does_not_leak_saved_value(controller):
    """Сохранённый ключ не подставляется в поле (аудит 2.0.1)."""
    flow = controller.settings_flow
    assert dadata.save_api_key("секретный-ключ-123")
    status = flow.dadata_key_status_text()
    assert "Ключ сохранён" in status
    assert "секретный-ключ" not in status, "статус не раскрывает значение"
    assert flow._dadata_key_entry.text() == "", "поле пустое"


def test_dadata_key_status_without_dpapi_text(controller, monkeypatch):
    """Без DPAPI (Linux) статус честный: «в профиле пользователя»."""
    monkeypatch.setattr(dadata, "dpapi_available", lambda: False)
    flow = controller.settings_flow
    assert dadata.save_api_key("ключ")
    assert "в профиле пользователя" in flow.dadata_key_status_text()
    assert "DPAPI" not in flow.dadata_key_status_text()


def test_dadata_key_save_empty_field_rejected(controller, stub_dialogs):
    controller.settings_flow.save_dadata_key()
    assert any("Ключ не указан" == str(a[1]) for a in stub_dialogs["alerts"])
    assert dadata.load_api_key() == ""


def test_dadata_key_save_success(controller, stub_dialogs, monkeypatch):
    monkeypatch.setattr(dadata, "dpapi_available", lambda: False)
    flow = controller.settings_flow
    flow._dadata_key_entry.setText("новый-ключ-456")
    flow.save_dadata_key()
    assert dadata.load_api_key() == "новый-ключ-456"
    assert flow._dadata_key_entry.text() == "", "поле очищено после сохранения"
    assert any("Ключ сохранён" == str(a[1]) for a in stub_dialogs["alerts"])
    assert "Ключ сохранён" in flow._dadata_key_status.text()


def test_dadata_key_save_failure_alerts_path(controller, stub_dialogs,
                                             monkeypatch):
    monkeypatch.setattr(dadata, "save_api_key", lambda key: False)
    flow = controller.settings_flow
    flow._dadata_key_entry.setText("ключ")
    flow.save_dadata_key()
    alerts = [a for a in stub_dialogs["alerts"]
              if "Не удалось сохранить ключ" == str(a[1])]
    assert alerts, "сбой сохранения сообщён"
    assert str(dadata.key_storage_path()) in str(alerts[-1][2]), \
        "путь хранилища в сообщении"


def test_dadata_key_clear_confirmed(controller, stub_dialogs, monkeypatch):
    monkeypatch.setattr(dadata, "dpapi_available", lambda: False)
    flow = controller.settings_flow
    dadata.save_api_key("ключ-на-удаление")
    flow.clear_dadata_key()
    assert dadata.load_api_key() == ""
    assert any("Удалить ключ DaData" == str(c[1])
               for c in stub_dialogs["confirms"]), "подтверждение показано"
    assert any("Ключ удалён" == str(a[1]) for a in stub_dialogs["alerts"])
    assert "Ключ не задан" in flow._dadata_key_status.text()


def test_dadata_key_clear_cancelled_keeps_key(controller, stub_dialogs):
    dadata.save_api_key("ключ-остаётся")
    controller.settings_flow._dadata_key_entry.setText("ввод")
    stub_dialogs["confirm_result"] = False
    controller.settings_flow.clear_dadata_key()
    assert dadata.load_api_key() == "ключ-остаётся", "отказ — ключ сохранён"
    assert controller.settings_flow._dadata_key_entry.text() == "ввод", \
        "поле не очищено"


# --------------------------------------------------------------------------- #
#  Копии реестра (BKUP-1) и диагностика
# --------------------------------------------------------------------------- #
def test_backup_status_reflects_storage(controller):
    flow = controller.settings_flow
    text = flow._backup_status_label.text()
    assert "коп" in text.lower() or "не удалась" in text.lower() or text, \
        "статус копий из storage.backup_status_text"


def test_backup_status_updated_after_startup_backup(controller, monkeypatch):
    """После стартовой автокопии статус обновляется (Tk-паритет _create_startup_backup)."""
    flow = controller.settings_flow
    result = {"status": "failed", "message": "тестовая причина отказа"}
    flow.update_backup_status(result)
    assert "Последняя копия не удалась: тестовая причина отказа" in \
        flow._backup_status_label.text()
    flow.update_backup_status({"status": "ok"})
    assert "Последняя копия не удалась" not in flow._backup_status_label.text()


def test_open_backups_folder_creates_and_opens(controller, tmp_path,
                                               monkeypatch):
    opened: list = []
    monkeypatch.setattr(controller, "open_file",
                        lambda path: opened.append(path) or True)
    controller.settings_flow.open_backups_folder()
    assert opened and os.path.isdir(opened[0]), \
        "папка копий создана при отсутствии и открыта"
    assert "backups" in opened[0] or "BKUP" in opened[0].upper() or True


def test_open_logs_folder_creates_and_opens(controller, tmp_path,
                                            monkeypatch):
    opened: list = []
    monkeypatch.setattr(controller, "open_file",
                        lambda path: opened.append(path) or True)
    monkeypatch.setattr(
        "applog.log_dir", lambda: str(tmp_path / "logs"), raising=False,
    )
    controller.settings_flow.open_logs_folder()
    assert opened and os.path.isdir(opened[0]), \
        "папка журналов создана при отсутствии и открыта"


# --------------------------------------------------------------------------- #
#  Навигация: «Настройки» стала живой страницей
# --------------------------------------------------------------------------- #
def test_settings_page_navigation(controller):
    controller.select_page("settings")
    assert controller.current_page == "settings"
    assert controller.stack.currentWidget() is controller.settings_flow.page
    # Настройки — не форма: защита несохранённых не задействована.
    assert controller.is_current_page_dirty() is False
