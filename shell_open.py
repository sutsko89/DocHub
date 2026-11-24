"""Открытие файлов и показ в проводнике без интерфейса (этап M3 миграции).

GUI-free ядро операций «Открыть файл» и «Показать в папке» — общий для
Qt-страниц (CP1: тост сохранения спецификации) и будущих этапов M4+.
Tk-версии main.open_file / main.show_in_folder остаются как есть до M8
(они связаны с Tk-алертом ui.alert и уходят вместе с Tk-слоем): здесь —
те же системные вызовы и та же диагностика в applog, но без диалогов;
функции возвращают успех, а сообщение об отказе показывает вызывающий
слой (main_qt._open_path / _show_path — qt-алерты, как и было в M2).

Контракты сохранены дословно (п. 6.2):

* отсутствующий файл — сразу отказ (xdg-open возвращает 0 даже для
  несуществующего пути), вызывающий код честно сообщает «сохранён»
  без ложного «открыт»;
* show_in_folder показывает файл в проводнике, НЕ запуская его.
"""

from __future__ import annotations

import ctypes
import os
import platform
import subprocess
import sys

import applog


def documents_known_folder() -> str:
    """Папка «Документы» пользователя (п. 6.2) — GUI-free, для Qt-страниц.

    Windows: SHGetKnownFolderPath(FOLDERID_Documents). На не-Windows и при
    любом сбое — fallback «~/Documents», как у Tk-версии
    win_shell.documents_known_folder. Копия сознательна: win_shell
    импортирует tkinter (хуки окна) и в Qt-сборку не включается; Tk-версия
    уходит вместе с Tk-слоем (M8), тогда копия объединяется.
    """
    if sys.platform == "win32":
        try:
            import uuid as _uuid

            class _GUID(ctypes.Structure):
                _fields_ = [
                    ("Data1", ctypes.c_ulong),
                    ("Data2", ctypes.c_ushort),
                    ("Data3", ctypes.c_ushort),
                    ("Data4", ctypes.c_ubyte * 8),
                ]

            guid = _GUID()
            known = _uuid.UUID("FDD39AD0-238F-46AF-ADB4-6C85480369C7")  # FOLDERID_Documents
            guid.Data1, guid.Data2, guid.Data3 = (
                known.time_low, known.time_mid, known.time_hi_version,
            )
            guid.Data4[:] = known.bytes[8:]
            out = ctypes.c_void_p()
            hr = ctypes.windll.shell32.SHGetKnownFolderPath(
                ctypes.byref(guid), 0, None, ctypes.byref(out)
            )
            if hr == 0 and out.value:
                documents = ctypes.wstring_at(out.value)
                ctypes.windll.ole32.CoTaskMemFree(out)
                return documents
            if out.value:
                ctypes.windll.ole32.CoTaskMemFree(out)
        except Exception as exc:  # noqa: BLE001 — есть USERPROFILE-fallback
            applog.get_logger("files").warning(
                "SHGetKnownFolderPath не сработал (%s) — fallback ~/Documents",
                type(exc).__name__,
            )
        profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        return os.path.join(profile, "Documents")
    return os.path.join(os.path.expanduser("~"), "Documents")


def open_file_system(file_path) -> bool:
    """Открывает файл системным приложением и возвращает успех.

    Отсутствующий файл — False и запись в журнал (без диалога). Ошибки
    запуска тоже False: сообщение формирует вызывающий слой.
    """
    if not file_path or not os.path.exists(file_path):
        applog.get_logger("files").warning(
            "Открытие отменено: файла нет на диске (%s)",
            os.path.splitext(str(file_path or ""))[1].lower(),
        )
        return False
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)  # noqa: S606
        elif platform.system() == "Darwin":
            if subprocess.run(["open", file_path], check=False).returncode != 0:
                return False
        else:
            if subprocess.run(["xdg-open", file_path], check=False).returncode != 0:
                return False
        return True
    except Exception as exc:
        applog.get_logger("files").warning(
            "Не удалось открыть файл (%s): %s",
            os.path.splitext(str(file_path))[1].lower(), type(exc).__name__,
        )
        return False


def show_in_file_manager(file_path: str) -> bool:
    """Показ файла в проводнике без запуска; False — файла нет/сбой вызова.

    На Windows — explorer /select (файл выделяется), на macOS — open -R,
    на остальных — открытие папки файла через xdg-open.
    """
    if not file_path or not os.path.exists(file_path):
        applog.get_logger("files").warning(
            "Показ в папке отменён: файла нет на диске (%s)",
            os.path.splitext(str(file_path or ""))[1].lower(),
        )
        return False
    try:
        if platform.system() == "Windows":
            subprocess.run(
                ["explorer", "/select,", os.path.normpath(file_path)], check=False,
            )
        elif platform.system() == "Darwin":
            subprocess.run(["open", "-R", file_path], check=False)
        else:
            subprocess.run(["xdg-open", os.path.dirname(file_path)], check=False)
        return True
    except Exception as exc:
        applog.get_logger("files").warning(
            "Не удалось открыть папку: %s", type(exc).__name__,
        )
        return False
