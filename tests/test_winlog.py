"""WINLOG-1 — диагностируемость Win32-слоя и кнопка «Открыть папку журналов».

Этап 2 roadmap «Отзывчивость оболочки и диагностируемость»: история
вылетов 2.17.0–2.19.3 с пустым crash.log показала цену недиагностируемости.
Закрепляются контракты:

  * сбой Win32-вызова оставляет запись в app.log с именем операции
    (_log_win32_failure — до и после правок жестов перенос/растягивание);
  * DragFinish освобождает HDROP при любом исходе, включая сбой
    DragQueryFileW посреди списка (D7);
  * деградация мьютекса Global → Local не молчалива (D10);
  * кнопка на «Настройках» открывает папку журналов через общий фасад
    ``app._open_file`` и работает на свежей установке без папки logs.

Тесты идут без Win32: лог-хелперы и _drag_files проверяются на
duck-typed фейках, платформозависимые ветки покрыты smoke-тестом.
"""
from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import applog
import settings_controller as settings_ctrl
import single_instance
import win_shell


def capture_logger(child_name: str, records: list) -> logging.Handler:
    """Ловушка записей дочернего логгера ewdochub (propagate=False)."""

    class _Capture(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            records.append(record)

    logger = logging.getLogger(applog.LOGGER_NAME + "." + child_name)
    handler = _Capture(level=logging.WARNING)
    logger.addHandler(handler)
    return handler


@pytest.fixture()
def winlog_records():
    records: list = []
    handler = capture_logger("win_shell", records)
    try:
        yield records
    finally:
        logging.getLogger(applog.LOGGER_NAME + ".win_shell").removeHandler(handler)


@pytest.fixture()
def mutex_records():
    records: list = []
    handler = capture_logger("single_instance", records)
    try:
        yield records
    finally:
        logging.getLogger(applog.LOGGER_NAME + ".single_instance").removeHandler(handler)


class FakeApp:
    """Приложение с одним фиксируемым фасадом открытия файлов."""

    def __init__(self):
        self.opened: list[str] = []

    def _open_file(self, path):
        self.opened.append(path)


# --------------------------------------------------------------------------- #
#  Сбой Win32-вызова → запись в app.log с именем операции
# --------------------------------------------------------------------------- #
def test_win32_failure_logs_operation_name(winlog_records):
    win_shell._log_win32_failure("MonitorFromPoint", OSError("сбой"))
    assert len(winlog_records) == 1
    message = winlog_records[0].getMessage()
    assert "MonitorFromPoint" in message, "имя операции в записи лога"
    assert "OSError" in message, "тип исключения в записи лога"


def test_win32_failure_logger_never_raises(monkeypatch):
    """Лог-хелпер не поднимает исключение, даже если журнал недоступен."""

    def boom(name):
        raise RuntimeError("журнал недоступен")

    monkeypatch.setattr(applog, "get_logger", boom)
    win_shell._log_win32_failure("GetMonitorInfoW", OSError())  # не падает


def test_all_win32_degradations_are_logged_not_silent():
    """Контракт WINLOG-1: молчаливых ``except Exception: pass`` без имени
    операции в win_shell.py не осталось. Tk-fallback'и (не Win32-вызовы)
    помечены ``tk-fallback`` — их тишина допустима и закомментирована."""
    source = Path(win_shell.__file__).read_text(encoding="utf-8")
    lines = source.splitlines()
    silent = []
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped == "except Exception:":
            body = lines[index + 1].strip() if index + 1 < len(lines) else ""
            if body.startswith("pass") and "tk-fallback" not in body:
                silent.append(index + 1)
    assert silent == [], (
        f"молчаливые except в win_shell.py (строки {silent}) должны "
        "именовать операцию через _log_win32_failure"
    )


# --------------------------------------------------------------------------- #
#  D7: DragFinish в finally
# --------------------------------------------------------------------------- #
class FakeBuffer:
    """Аналог create_unicode_buffer с записываемым .value."""

    def __init__(self, size):
        self.value = None
        self.size = size


class FakeShell32:
    """Duck-typed shell32: DragQueryFileW отдаёт имена или ломается."""

    def __init__(self, files=None, fail_on_call=None):
        self.files = list(files or [])
        self.fail_on_call = fail_on_call
        self.query_calls = 0
        self.finished: list[int] = []

    def DragQueryFileW(self, drop, index, buffer, size):
        self.query_calls += 1
        if self.fail_on_call is not None and self.query_calls >= self.fail_on_call:
            raise OSError("сбой чтения HDROP")
        if index == 0xFFFFFFFF:
            return len(self.files)
        value = self.files[index]
        if buffer is not None:
            buffer.value = value
        return len(value)

    def DragFinish(self, drop):
        self.finished.append(drop)


def test_drag_files_returns_names_and_finishes_drop(monkeypatch):
    monkeypatch.setattr(win_shell.ctypes, "create_unicode_buffer", FakeBuffer)
    shell = FakeShell32(files=["C:\\дог\\акт.docx", "C:\\дог\\счёт.docx"])
    files = win_shell._drag_files(shell, 42)
    assert files == ["C:\\дог\\акт.docx", "C:\\дог\\счёт.docx"]
    assert len(shell.finished) == 1, "DragFinish вызван ровно один раз"


def test_drag_files_finishes_drop_even_when_query_fails(monkeypatch):
    """D7: сбой DragQueryFileW посреди списка не пропускает DragFinish."""
    monkeypatch.setattr(win_shell.ctypes, "create_unicode_buffer", FakeBuffer)
    shell = FakeShell32(files=["a.docx", "b.docx"], fail_on_call=3)
    with pytest.raises(OSError):
        win_shell._drag_files(shell, 7)
    assert len(shell.finished) == 1, "HDROP освобождён в finally"


def test_drag_files_swallows_dragfinish_failure(monkeypatch):
    """Сбой самого DragFinish не перекрывает исход чтения файлов."""
    monkeypatch.setattr(win_shell.ctypes, "create_unicode_buffer", FakeBuffer)
    shell = FakeShell32(files=["a.docx"])
    shell.DragFinish = lambda drop: (_ for _ in ()).throw(OSError("нет доступа"))
    files = win_shell._drag_files(shell, 1)
    assert files == ["a.docx"], "сбой освобождения не теряет список файлов"


# --------------------------------------------------------------------------- #
#  D10: деградация Global → Local видна в журнале
# --------------------------------------------------------------------------- #
def test_mutex_degradation_is_logged(mutex_records):
    single_instance._log_mutex_degradation(
        "Global\\EWDocHub_SingleInstance_sid", 5,
    )
    assert len(mutex_records) == 1
    message = mutex_records[0].getMessage()
    assert "Global\\EWDocHub_SingleInstance_sid" in message
    assert "Local" in message, "деградация названа прямо"
    assert mutex_records[0].levelname == "WARNING"


def test_mutex_degradation_logger_never_raises(monkeypatch):
    def boom(name):
        raise RuntimeError("журнал недоступен")

    monkeypatch.setattr(applog, "get_logger", boom)
    single_instance._log_mutex_degradation("Global\\X", 5)  # не падает


def test_mutex_candidates_keep_local_fallback():
    names = single_instance.InstanceLock._mutex_candidates()
    assert names[0].startswith("Global\\") and names[2].startswith("Local\\"), \
        "после недоступного Global идёт Local с тем же именем"


# --------------------------------------------------------------------------- #
#  Кнопка «Открыть папку журналов» (Настройки)
# --------------------------------------------------------------------------- #
def test_open_logs_folder_creates_and_opens(appdata):
    app = FakeApp()
    settings_ctrl.SettingsController(app).open_logs_folder()
    assert app.opened == [applog.log_dir()], "открывается именно папка журналов"
    assert os.path.isdir(applog.log_dir()), "папка создана на свежей установке"


def test_open_logs_folder_failure_still_calls_opener(monkeypatch, appdata):
    def boom(path, exist_ok=False):
        raise OSError("недоступно")

    monkeypatch.setattr(settings_ctrl.os, "makedirs", boom)
    app = FakeApp()
    settings_ctrl.SettingsController(app).open_logs_folder()
    assert app.opened == [applog.log_dir()], "сбой создания не блокирует открытие"
