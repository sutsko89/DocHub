"""Защита от второго экземпляра программы (п. 2.2 плана).

Два запущенных экземпляра писали в один ``database.json`` и могли затереть
изменения друг друга. Здесь используются именованные Win32-mutex'ы: они
освобождаются автоматически при завершении процесса (в том числе аварийном).

Пространство имён (аудит 2.0.1, P2 «Мьютекс сессионный»):
  * ``Global\\EWDocHub_SingleInstance_<SID>`` — защита данных на пользователя:
    одна копия во всех сессиях (консоль + RDP), у разных пользователей свои
    реестры и свои блокировки;
  * ``Global\\EWDocHub_InstallerMutex`` — фиксированный «инсталляторный»
    мьютекс: по нему Inno Setup находит запущенную копию (AppMutex) независимо
    от имени пользователя; если ``Global\\`` недоступен — ``Local\\``-вариант
    с тем же именем (оба перечислены в installer.iss).

На не-Windows платформах блокировка считается успешной: проект
ориентирован на Windows, а на других ОС программа не поставляется.
"""

from __future__ import annotations

import ctypes
import sys

import applog

INSTALLER_MUTEX_NAME = "Global\\EWDocHub_InstallerMutex"
WINDOW_TITLE = "EW.DocHub — контрагенты и документы"
ERROR_ALREADY_EXISTS = 183

if sys.platform == "win32":
    from ctypes import wintypes


def _user_sid() -> str:
    """SID текущего пользователя Windows (строка вида «S-1-5-21-…»).

    SID надёжнее имени: он не меняется при переименовании учётной записи.
    Пустая строка — определить не удалось (тогда мьютекс без SID).
    """
    if sys.platform != "win32":
        return ""
    try:
        advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        TOKEN_QUERY = 0x0008
        TokenUser = 1

        token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(
            kernel32.GetCurrentProcess(), TOKEN_QUERY, ctypes.byref(token),
        ):
            return ""
        try:
            length = wintypes.DWORD(0)
            advapi32.GetTokenInformation(token, TokenUser, None, 0, ctypes.byref(length))
            if not length.value:
                return ""
            buffer = ctypes.create_string_buffer(length.value)
            if not advapi32.GetTokenInformation(
                token, TokenUser, buffer, length.value, ctypes.byref(length),
            ):
                return ""

            class _SID_AND_ATTRIBUTES(ctypes.Structure):
                _fields_ = [("Sid", ctypes.c_void_p), ("Attributes", wintypes.DWORD)]

            user = ctypes.cast(buffer, ctypes.POINTER(_SID_AND_ATTRIBUTES)).contents
            sid_string = wintypes.LPWSTR()
            if not advapi32.ConvertSidToStringSidW(user.Sid, ctypes.byref(sid_string)):
                return ""
            try:
                return sid_string.value or ""
            finally:
                kernel32.LocalFree(sid_string)
        finally:
            kernel32.CloseHandle(token)
    except Exception:
        return ""


def mutex_name() -> str:
    """Имя мьютекса защиты данных: ``Global\\`` + SID пользователя.

    ``Global\\`` закрывает дыру ``Local\\`` (аудит 2.0.1): две сессии одного
    пользователя (консоль + RDP) получают одну и ту же блокировку. SID не даёт
    одному пользователю блокировать другого. Если SID определить не удалось —
    общее имя в ``Global\\``: лишняя блокировка безопаснее её отсутствия.
    """
    sid = _user_sid()
    if sid:
        return f"Global\\EWDocHub_SingleInstance_{sid}"
    return "Global\\EWDocHub_SingleInstance"


# Имя текущей сессии — для проверок smoke-теста и совместимости.
MUTEX_NAME = mutex_name()


class InstanceLock:
    """Именованные mutex'и; ``acquire()`` возвращает False, если программа уже запущена."""

    def __init__(self) -> None:
        self._handles: list = []

    def acquire(self) -> bool:
        if sys.platform != "win32":
            return True
        try:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.CreateMutexW.argtypes = [
                ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR,
            ]
            kernel32.CreateMutexW.restype = wintypes.HANDLE
            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            kernel32.CloseHandle.restype = wintypes.BOOL
            for name in self._mutex_candidates():
                handle = kernel32.CreateMutexW(None, False, name)
                # GetLastError читаем сразу после вызова: use_last_error=True
                # сохраняет именно ошибку этого вызова (аудит 2.0.1).
                exists = ctypes.get_last_error() == ERROR_ALREADY_EXISTS
                if exists and handle:
                    kernel32.CloseHandle(handle)
                    self._release_all()
                    return False
                if not handle:
                    # Создать не удалось (например, нет прав на Global\ в этой
                    # сессии) — пробуем следующий вариант имени, но не роняем
                    # и не блокируем запуск.
                    continue
                self._handles.append(handle)
            return True
        except Exception as exc:  # noqa: BLE001 — блокировка не должна ронять запуск
            applog.get_logger("single_instance").warning(
                "Не удалось установить блокировку экземпляра: %s", type(exc).__name__,
            )
            self._release_all()
            return True

    @staticmethod
    def _mutex_candidates() -> list[str]:
        """Имена мьютексов: Global, при отказе в доступе — Local (та же семантика)."""
        names = [mutex_name(), INSTALLER_MUTEX_NAME]
        return names + [name.replace("Global\\", "Local\\", 1) for name in names]

    def release(self) -> None:
        self._release_all()

    def _release_all(self) -> None:
        handles, self._handles = self._handles, []
        if sys.platform != "win32" or not handles:
            return
        try:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            kernel32.CloseHandle.restype = wintypes.BOOL
            for handle in handles:
                kernel32.CloseHandle(handle)
        except Exception:
            pass


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
