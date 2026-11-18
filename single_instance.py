"""Защита от второго экземпляра программы (п. 2.2 плана).

Два запущенных экземпляра писали в один ``database.json`` и могли затереть
изменения друг друга. Здесь используется именованный Win32-mutex: он
освобождается автоматически при завершении процесса (в том числе аварийном).

На не-Windows платформах блокировка считается успешной: проект
ориентирован на Windows, а на других ОС программа не поставляется.
"""

from __future__ import annotations

import ctypes
import sys

import applog

MUTEX_NAME = "Local\\EWDocHub_SingleInstance"
WINDOW_TITLE = "EW.DocHub — контрагенты и документы"
ERROR_ALREADY_EXISTS = 183


class InstanceLock:
    """Именованный mutex; ``acquire()`` возвращает False, если программа уже запущена."""

    def __init__(self) -> None:
        self._handle = None

    def acquire(self) -> bool:
        if sys.platform != "win32":
            return True
        try:
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
            if not handle:
                # Не удалось создать mutex — не блокируем запуск.
                return True
            if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
                kernel32.CloseHandle(handle)
                return False
            self._handle = handle
            return True
        except Exception as exc:  # noqa: BLE001 — блокировка не должна ронять запуск
            applog.get_logger("single_instance").warning(
                "Не удалось установить блокировку экземпляра: %s", type(exc).__name__,
            )
            return True

    def release(self) -> None:
        if self._handle is None or sys.platform != "win32":
            return
        try:
            ctypes.windll.kernel32.CloseHandle(self._handle)
        except Exception:  # noqa: BLE001
            pass
        finally:
            self._handle = None


def activate_existing_window() -> bool:
    """Пытается показать и активировать окно уже запущенной программы."""
    if sys.platform != "win32":
        return False
    try:
        user32 = ctypes.windll.user32
        hwnd = user32.FindWindowW(None, WINDOW_TITLE)
        if not hwnd:
            return False
        SW_RESTORE = 9
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)
        return True
    except Exception:  # noqa: BLE001
        return False


def show_already_running() -> None:
    """Сообщает пользователю, что программа уже запущена."""
    message = (
        "EW.DocHub уже запущен.\n\n"
        "Одновременная работа двух копий запрещена: они пишут в один файл "
        "реестра и могут потерять данные.\n\n"
        "Переключитесь на открытое окно программы."
    )
    if sys.platform == "win32":
        try:
            MB_ICONINFORMATION = 0x40
            ctypes.windll.user32.MessageBoxW(
                0, message, "EW.DocHub", MB_ICONINFORMATION,
            )
            return
        except Exception:  # noqa: BLE001
            pass
    print(message, file=sys.stderr)
