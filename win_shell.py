"""Win32-обвязка окна (п. 9.3): безрамочный режим, DWM, мониторы, DPI.

Выделено из ui_macos.py и main.py: всё, что касается операционной системы
(ctypes/user32/shell32/dwmapi), живёт здесь — путь к кроссплатформенности
остального интерфейса. На Linux/macOS функции деградируют в безопасные
аналоги Tk или no-op.
"""
from __future__ import annotations

import ctypes
import os
import sys

import tkinter as tk

#: Сообщения окна, обрабатываемые единым подклассом WndProc (п. 6.3).
WM_DROPFILES = 0x0233
WM_DPICHANGED = 0x02E0

# Нативное перетаскивание/растягивание безрамочного окна (2.17.0):
# Windows сама ведёт модальный цикл переноса после синтетического
# WM_NCLBUTTONDOWN с HT-кодом «куда попали мышью». В этом цикле работают
# перенос на второй монитор (со сменой DPI) и Aero Snap — то, что
# недоступно ручному переносу через geometry().
WM_NCLBUTTONDOWN = 0x00A1
HTCAPTION = 0x2
#: HT-коды кромок для нативного растягивания (WM_NCLBUTTONDOWN).
HT_RESIZE_EDGES = {
    "w": 10,   # HTLEFT
    "e": 11,   # HTRIGHT
    "n": 12,   # HTTOP
    "nw": 13,  # HTTOPLEFT
    "ne": 14,  # HTTOPRIGHT
    "s": 15,   # HTBOTTOM
    "sw": 16,  # HTBOTTOMLEFT
    "se": 17,  # HTBOTTOMRIGHT
}
_SWP_FRAMECHANGED_FLAGS = 0x1 | 0x2 | 0x4 | 0x20  # NOMOVE|NOSIZE|NOZORDER|FRAMECHANGED


class _RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]
def _main_hwnd(window: tk.Misc) -> int:
    window.update_idletasks()
    hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
    return hwnd or window.winfo_id()
def enable_rounded_corners(window: tk.Misc) -> None:
    """Скругление углов силами DWM (Windows 11)."""
    if sys.platform != "win32":
        return
    try:
        hwnd = _main_hwnd(window)
        preference = ctypes.c_int(2)  # DWMWCP_ROUND
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 33, ctypes.byref(preference), ctypes.sizeof(preference)
        )
    except Exception:
        pass
def make_borderless(window: tk.Misc, taskbar_button: bool = True) -> None:
    """Убирает системную рамку, оставляя окно в панели задач и Alt+Tab.

    Стилевые биты WS_MAXIMIZEBOX/WS_MINIMIZEBOX/WS_SYSMENU (2.17.0) не
    рисуют ничего (рамки WS_CAPTION/WS_THICKFRAME не добавляются), но
    сообщают Windows, что окно можно разворачивать: без них Aero Snap
    считает окно неразворачиваемым и не прилипает к краям.
    """
    window.overrideredirect(True)
    if sys.platform != "win32":
        return
    try:
        hwnd = _main_hwnd(window)
        user32 = ctypes.windll.user32
        GWL_EXSTYLE = -20
        WS_EX_APPWINDOW = 0x00040000
        WS_EX_TOOLWINDOW = 0x00000080
        style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if taskbar_button:
            style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
        user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        # Разворачиваемость для Aero Snap: биты без визуального следа.
        GWL_STYLE = -16
        WS_MAXIMIZEBOX = 0x00010000
        WS_MINIMIZEBOX = 0x00020000
        WS_SYSMENU = 0x00080000
        base_style = user32.GetWindowLongW(hwnd, GWL_STYLE)
        user32.SetWindowLongW(
            hwnd, GWL_STYLE,
            base_style | WS_MAXIMIZEBOX | WS_MINIMIZEBOX | WS_SYSMENU,
        )
        user32.SetWindowPos(
            ctypes.c_void_p(hwnd), None, 0, 0, 0, 0,
            _SWP_FRAMECHANGED_FLAGS,
        )
        enable_rounded_corners(window)
    except Exception:
        pass
def focus_window(window: tk.Misc) -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.SetForegroundWindow(_main_hwnd(window))
            return
        except Exception:
            pass
    try:
        window.focus_force()
    except Exception:
        pass
def refresh_taskbar(window: tk.Misc) -> None:
    """Пересоздаёт кнопку в панели задач после смены стиля окна."""
    if sys.platform != "win32":
        return
    try:
        window.withdraw()
        window.after(12, lambda: (window.deiconify(), focus_window(window)))
    except Exception:
        pass
class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _RECT),
        ("rcWork", _RECT),
        ("dwFlags", ctypes.c_ulong),
    ]
def workarea(window: tk.Misc | None = None) -> tuple[int, int, int, int]:
    """Рабочая область (без панели задач): x, y, ширина, высота.

    Если передано окно — берётся рабочая область ТОГО монитора, где окно
    находится сейчас (п. 6.3), а не только основного монитора.
    """
    if sys.platform == "win32":
        if window is not None:
            try:
                hwnd = _main_hwnd(window)
                hmon = ctypes.windll.user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
                if hmon:
                    info = _MONITORINFO()
                    info.cbSize = ctypes.sizeof(info)
                    if ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
                        rc = info.rcWork
                        return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top
            except Exception:
                pass
        rect = _RECT()
        try:
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass
    else:
        # Не Windows: используем размеры экрана Tk (п. 6.3 — не «только 1920×1080»).
        try:
            source = window or getattr(tk, "_default_root", None)
            if source is not None:
                return 0, 0, source.winfo_screenwidth(), source.winfo_screenheight()
        except Exception:
            pass
    return 0, 0, 1920, 1080
def minimize_window(window: tk.Misc) -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.ShowWindow(_main_hwnd(window), 6)  # SW_MINIMIZE
            return
        except Exception:
            pass
    try:
        window.iconify()
    except Exception:
        pass
def maximize_window(window: tk.Misc) -> None:
    """Разворачивает окно ровно на рабочую область: панель задач остаётся видимой.

    Для безрамочного окна (overrideredirect) системная максимизация
    SW_MAXIMIZE накрывает панель задач, поэтому задаём геометрию сами.
    """
    if sys.platform == "win32":
        try:
            x, y, width, height = workarea(window)
            window.geometry(f"{width}x{height}+{x}+{y}")
            return
        except Exception:
            pass
    try:
        window.state("zoomed")
    except Exception:
        pass
def restore_window(window: tk.Misc, geometry: str = "") -> None:
    """Возврат из максимизации: применяет прежнюю геометрию, если она задана."""
    if sys.platform == "win32":
        try:
            if geometry:
                window.geometry(geometry)
            else:
                ctypes.windll.user32.ShowWindow(_main_hwnd(window), 9)  # SW_RESTORE
            return
        except Exception:
            pass
    try:
        window.state("normal")
    except Exception:
        pass


def _begin_native_ht(window: tk.Misc, ht_code: int) -> bool:
    """Запускает нативный модальный цикл Windows с HT-кодом точки.

    ReleaseCapture сбрасывает захват мыши Tk, после чего синтетический
    WM_NCLBUTTONDOWN передаёт Windows управление переносом/растягиванием:
    окно можно унести на второй монитор (со сменой DPI) и прилипнуть
    к краям (Aero Snap). Блокирует до отпускания кнопки мыши.
    False — не-Windows или сбой (вызывающий использует ручной fallback).
    """
    if sys.platform != "win32":
        return False
    try:
        hwnd = _main_hwnd(window)
        user32 = ctypes.windll.user32
        user32.ReleaseCapture()
        user32.SendMessageW.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_longlong,
        ]
        user32.SendMessageW.restype = ctypes.c_longlong
        user32.SendMessageW(
            ctypes.c_void_p(hwnd), WM_NCLBUTTONDOWN, ht_code, 0,
        )
        return True
    except Exception:
        return False


def begin_native_move(window: tk.Misc) -> bool:
    """Нативное перетаскивание окна за шапку (2.17.0).

    Вызывать из обработчика нажатия мыши на шапке. True — Windows взяла
    перенос на себя (второй монитор, Snap, смена DPI); False — не-Windows
    или сбой, нужен ручной перенос через geometry().
    """
    return _begin_native_ht(window, HTCAPTION)


def begin_native_resize(window: tk.Misc, edge: str) -> bool:
    """Нативное растягивание за кромку ``edge`` (2.17.0).

    ``edge`` — одна из 8 сторон: 'n', 's', 'w', 'e', 'nw', 'ne', 'sw',
    'se'. True — Windows ведёт цикл растягивания сама (плавно, с учётом
    монитора под курсором и minsize); False — ручной fallback.
    """
    ht_code = HT_RESIZE_EDGES.get(str(edge).lower())
    if not ht_code:
        return False
    return _begin_native_ht(window, ht_code)


# ----------------------------------------------------------------------- #
#  DPI и хуки окна (п. 6.3)
# ----------------------------------------------------------------------- #
def enable_dpi_awareness() -> None:
    """Per-Monitor DPI awareness: чёткий текст на экранах 125–150 %.

    Вызывается до создания первого окна Tk. Повторный вызов безопасен:
    при уже установленном режиме Windows возвращает ошибку, которая
    молча игнорируется. На не-Windows — no-op.
    """
    if sys.platform != "win32":
        return
    try:  # чёткий текст на экранах с масштабированием (125 %, 150 % и т. д.)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def dpi_scale_from_wparam(wparam: int) -> float:
    """Масштаб интерфейса из ``wparam`` сообщения WM_DPICHANGED.

    HIWORD(wparam) — новый DPI окна (96 = 100 %). Чистая функция без
    обращения к Win32 — используется и тестами.
    """
    dpi = (int(wparam) >> 16) & 0xFFFF
    return dpi / 96.0 if dpi else 1.0


def suggested_rect_from_lparam(lparam: int) -> tuple[int, int, int, int]:
    """Прямоугольник-подсказка Windows (LPRECT) из ``lparam`` WM_DPICHANGED.

    Окно обязано встать в этот прямоугольник: Windows уже учла новый
    масштаб монитора, иначе после переноса окно вылезает за границы.
    """
    rc = _RECT.from_address(int(lparam))
    return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top


def documents_known_folder() -> str:
    """Папка «Документы» пользователя (п. 6.2): видимое место вместо %APPDATA%.

    Windows: SHGetKnownFolderPath(FOLDERID_Documents). На не-Windows и при
    любом сбое — безопасные fallback-и прежнего кода main.py.
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
        except Exception:
            pass
        profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        return os.path.join(profile, "Documents")
    return os.path.join(os.path.expanduser("~"), "Documents")


def install_window_hooks(
    window: tk.Misc,
    on_files=None,
    on_dpi_changed=None,
) -> None:
    """Один подкласс WndProc на окно: перетаскивание файлов + смена DPI.

    ``on_files(list[str])`` вызывается на WM_DROPFILES (п. 4.5);
    ``on_dpi_changed(dpi, rect)`` — на WM_DPICHANGED при переносе окна
    на монитор с другим масштабом (п. 6.3). Оба колбэка уходят в цикл
    Tk через ``window.after(0, ...)`` — выполняются в главном потоке.
    При любой ошибке хуки не устанавливаются: документы по-прежнему
    доступны кнопкой «Добавить документы…», геометрию при смене DPI
    правит сам Tk.
    """
    if sys.platform != "win32":
        return
    try:
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        user32.GetParent.restype = ctypes.c_void_p
        user32.GetParent.argtypes = [ctypes.c_void_p]
        shell32.DragAcceptFiles.argtypes = [ctypes.c_void_p, ctypes.c_bool]
        hwnd = user32.GetParent(ctypes.c_void_p(window.winfo_id())) or window.winfo_id()
        if on_files is not None:
            shell32.DragAcceptFiles(ctypes.c_void_p(hwnd), True)

        GWLP_WNDPROC = -4
        user32.GetWindowLongPtrW.restype = ctypes.c_void_p
        user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p
        user32.SetWindowLongPtrW.argtypes = [
            ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p,
        ]
        user32.CallWindowProcW.restype = ctypes.c_longlong
        user32.CallWindowProcW.argtypes = [
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
            ctypes.c_ulonglong, ctypes.c_longlong,
        ]
        old_proc = user32.GetWindowLongPtrW(ctypes.c_void_p(hwnd), GWLP_WNDPROC)
        if not old_proc:
            return

        wndproc_type = ctypes.WINFUNCTYPE(
            ctypes.c_longlong, ctypes.c_void_p, ctypes.c_uint,
            ctypes.c_ulonglong, ctypes.c_longlong,
        )

        def handler(h, msg, wparam, lparam):
            if msg == WM_DROPFILES and on_files is not None:
                try:
                    count = shell32.DragQueryFileW(
                        ctypes.c_void_p(wparam), 0xFFFFFFFF, None, 0
                    )
                    files: list[str] = []
                    for index in range(count):
                        length = shell32.DragQueryFileW(
                            ctypes.c_void_p(wparam), index, None, 0
                        )
                        buffer = ctypes.create_unicode_buffer(length + 1)
                        shell32.DragQueryFileW(
                            ctypes.c_void_p(wparam), index, buffer, length + 1
                        )
                        if buffer.value:
                            files.append(buffer.value)
                    shell32.DragFinish(ctypes.c_void_p(wparam))
                    if files:
                        window.after(0, lambda f=list(files): on_files(f))
                    return 0
                except Exception:
                    pass
            if msg == WM_DPICHANGED and on_dpi_changed is not None:
                try:
                    dpi = (int(wparam) >> 16) & 0xFFFF
                    if dpi:
                        rect = suggested_rect_from_lparam(lparam)
                        window.after(0, lambda d=dpi, r=rect: on_dpi_changed(d, r))
                        return 0
                except Exception:
                    pass
            return user32.CallWindowProcW(old_proc, h, msg, wparam, lparam)

        hook = wndproc_type(handler)
        user32.SetWindowLongPtrW(
            ctypes.c_void_p(hwnd), GWLP_WNDPROC,
            ctypes.cast(hook, ctypes.c_void_p),
        )
        # Ссылки держим на окне: иначе сборщик мусора убьёт колбэк WndProc.
        window._ew_hook = hook
        window._ew_hook_old_proc = old_proc
        window._ew_hook_hwnd = hwnd
    except Exception:
        pass
