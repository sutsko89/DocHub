"""Win32-обвязка окна (п. 9.3): безрамочный режим, DWM, мониторы.

Выделено из ui_macos.py: всё, что касается операционной системы
(ctypes/user32/dwmapi), живёт здесь — путь к кроссплатформенности
остального интерфейса. На Linux/macOS функции деградируют в безопасные
аналоги Tk или no-op.
"""
from __future__ import annotations

import ctypes
import sys

import tkinter as tk

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
    """Убирает системную рамку, оставляя окно в панели задач и Alt+Tab."""
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
