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

import applog  # WINLOG-1: деградации Win32-вызовов видны в app.log

#: Сообщения окна, обрабатываемые единым подклассом WndProc (п. 6.3).
WM_DROPFILES = 0x0233
WM_DPICHANGED = 0x02E0

#: Поля отступа при зажиме позиции окна внутри виртуального экрана:
#: у края экрана остаётся видимая полоса окна, за которую его можно схватить.
DRAG_EDGE_MARGIN = 48
#: Системные индексы GetSystemMetrics для виртуального экрана (объединение
#: всех мониторов): x, y, ширина, высота.
_SM_XVIRTUALSCREEN = 76
_SM_YVIRTUALSCREEN = 77
_SM_CXVIRTUALSCREEN = 78
_SM_CYVIRTUALSCREEN = 79
_SWP_FRAMECHANGED_FLAGS = 0x1 | 0x2 | 0x4 | 0x20  # NOMOVE|NOSIZE|NOZORDER|FRAMECHANGED


def _log_win32_failure(operation: str, exc: Exception) -> None:
    """WINLOG-1: сбой Win32-вызова остаётся в app.log с именем операции.

    История вылетов 2.17.0–2.19.3 с пустым crash.log показала цену
    недиагностируемости: 24 молчаливых except в этом модуле съедали
    и ошибку, и след. Теперь каждый сбой пишет предупреждение с
    системным именем операции (без персональных данных: тип исключения
    и имя вызова, без текста с путями).
    """
    try:
        applog.get_logger("win_shell").warning(
            "Win32 «%s» не выполнен: %s", operation, type(exc).__name__,
        )
    except Exception:  # noqa: BLE001 — лог никогда не роняет вызывающий код
        pass


class _RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]
def _main_hwnd(window: tk.Misc) -> int:
    window.update_idletasks()
    if sys.platform != "win32":
        return window.winfo_id()
    user32 = ctypes.windll.user32
    # restype — 64-битный HWND: без него ctypes вернёт усечённый c_int
    # (user-хендлы Windows пока укладываются в 32 бита, но полагаться
    # на это не стоит).
    user32.GetParent.restype = ctypes.c_void_p
    hwnd = user32.GetParent(window.winfo_id())
    return int(hwnd) if hwnd else window.winfo_id()
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
    except Exception as exc:  # noqa: BLE001 — скругление не критично
        _log_win32_failure("DwmSetWindowAttribute", exc)
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
    except Exception as exc:  # noqa: BLE001 — рамка не критична
        _log_win32_failure("SetWindowLongW/SetWindowPos", exc)
def focus_window(window: tk.Misc) -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.SetForegroundWindow(_main_hwnd(window))
            return
        except Exception as exc:  # noqa: BLE001 — есть fallback focus_force
            _log_win32_failure("SetForegroundWindow", exc)
    try:
        window.focus_force()
    except Exception:
        pass  # tk-fallback: не Win32, тишина допустима
def refresh_taskbar(window: tk.Misc) -> None:
    """Пересоздаёт кнопку в панели задач после смены стиля окна."""
    if sys.platform != "win32":
        return
    try:
        window.withdraw()
        window.after(12, lambda: (window.deiconify(), focus_window(window)))
    except Exception as exc:  # noqa: BLE001 — кнопка уже на месте
        _log_win32_failure("refresh_taskbar", exc)
class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _RECT),
        ("rcWork", _RECT),
        ("dwFlags", ctypes.c_ulong),
    ]
class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
def workarea(window: tk.Misc | None = None) -> tuple[int, int, int, int]:
    """Рабочая область (без панели задач): x, y, ширина, высота.

    Если передано окно — берётся рабочая область ТОГО монитора, где окно
    находится сейчас (п. 6.3), а не только основного монитора.
    """
    if sys.platform == "win32":
        if window is not None:
            try:
                hwnd = _main_hwnd(window)
                user32 = ctypes.windll.user32
                # D6 (WINLOG-1): argtypes/restype обязательны — без них
                # ctypes вернёт усечённый c_int вместо 64-битного HMONITOR
                # (как в MonitorFromPoint/workarea_at и _main_hwnd).
                user32.MonitorFromWindow.argtypes = [ctypes.c_void_p, ctypes.c_ulong]
                user32.MonitorFromWindow.restype = ctypes.c_void_p
                user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
                user32.GetMonitorInfoW.restype = ctypes.c_bool
                hmon = user32.MonitorFromWindow(
                    ctypes.c_void_p(hwnd), 2,  # MONITOR_DEFAULTTONEAREST
                )
                if hmon:
                    info = _MONITORINFO()
                    info.cbSize = ctypes.sizeof(info)
                    if user32.GetMonitorInfoW(
                        ctypes.c_void_p(hmon), ctypes.byref(info),
                    ):
                        rc = info.rcWork
                        return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top
            except Exception as exc:  # noqa: BLE001 — есть fallback ниже
                _log_win32_failure("MonitorFromWindow/GetMonitorInfoW", exc)
        rect = _RECT()
        try:
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception as exc:  # noqa: BLE001 — крайний fallback 1920×1080
            _log_win32_failure("SystemParametersInfoW(SPI_GETWORKAREA)", exc)
    else:
        # Не Windows: используем размеры экрана Tk (п. 6.3 — не «только 1920×1080»).
        try:
            source = window or getattr(tk, "_default_root", None)
            if source is not None:
                return 0, 0, source.winfo_screenwidth(), source.winfo_screenheight()
        except Exception:
            pass  # tk-fallback: не Win32, тишина допустима
    return 0, 0, 1920, 1080
def workarea_at(root: tk.Misc | None, x: int, y: int) -> tuple[int, int, int, int]:
    """Рабочая область монитора, на котором лежит точка (x, y) (2.20.0).

    Прилипание к краям (2.20.0) ориентируется на монитор ПОД КУРСОРом:
    пользователь тянет окно к краю конкретного монитора (в том числе
    второго), а не к краю того, где окно «в среднем». Сбой Win32 и
    не-Windows — единая рабочая область из ``workarea(root)`` (экран Tk),
    крайний fallback 1920×1080.
    """
    if sys.platform == "win32":
        try:
            user32 = ctypes.windll.user32
            # MonitorFromPoint принимает POINT по значению — argtypes
            # обязательны, restype — 64-битный хендл (как в _main_hwnd).
            user32.MonitorFromPoint.argtypes = [_POINT, ctypes.c_ulong]
            user32.MonitorFromPoint.restype = ctypes.c_void_p
            user32.GetMonitorInfoW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
            user32.GetMonitorInfoW.restype = ctypes.c_bool
            hmon = user32.MonitorFromPoint(
                _POINT(int(x), int(y)), 2  # MONITOR_DEFAULTTONEAREST
            )
            if hmon:
                # R-06 (PERF-2): GetMonitorInfoW — через кэш по HMONITOR.
                area = _cached_workarea(
                    int(hmon), lambda: _monitor_workarea_from(user32, hmon),
                )
                if area is not None:
                    return area
        except Exception as exc:  # noqa: BLE001 — есть fallback workarea()
            _log_win32_failure("MonitorFromPoint/GetMonitorInfoW", exc)
    return workarea(root)


#: R-06 (PERF-2): кэш рабочих областей по хендлу монитора. Жесты
#: перетаскивания спрашивают рабочую область ~60 раз в секунду; вызов
#: GetMonitorInfoW на каждый тик — разовый хитч на медленных машинах.
#: Кэш безопасен, потому что рабочая область монитора не меняется посреди
#: жеста; ключ — сам HMONITOR, поэтому мониторы не смешиваются (курсор
#: может переехать на второй монитор посреди переноса). Переполнение
#: (смена топологии дисплеев — редкое событие) очищает кэш целиком.
_WORKAREA_CACHE: dict[int, tuple[int, int, int, int]] = {}
_WORKAREA_CACHE_LIMIT = 8


def _cached_workarea(monitor_key: int, fetch) -> tuple[int, int, int, int] | None:
    """Кэш рабочей области по ключу монитора (R-06, точка подмены в тестах).

    ``fetch`` опрашивает Win32 ровно один раз на монитор; ``None``
    (сбой GetMonitorInfoW) не кэшируется — следующий тик пробует снова.
    """
    cached = _WORKAREA_CACHE.get(monitor_key)
    if cached is not None:
        return cached
    area = fetch()
    if area is None:
        return None
    if len(_WORKAREA_CACHE) >= _WORKAREA_CACHE_LIMIT:
        _WORKAREA_CACHE.clear()
    _WORKAREA_CACHE[monitor_key] = area
    return area


def _monitor_workarea_from(user32, hmon) -> tuple[int, int, int, int] | None:
    """Рабочая область конкретного монитора через GetMonitorInfoW."""
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(info)
    if user32.GetMonitorInfoW(ctypes.c_void_p(hmon), ctypes.byref(info)):
        rc = info.rcWork
        return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top
    return None
def minimize_window(window: tk.Misc) -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.ShowWindow(_main_hwnd(window), 6)  # SW_MINIMIZE
            return
        except Exception as exc:  # noqa: BLE001 — есть fallback iconify()
            _log_win32_failure("ShowWindow(SW_MINIMIZE)", exc)
    try:
        window.iconify()
    except Exception:
        pass  # tk-fallback: не Win32, тишина допустима
def maximize_window(window: tk.Misc) -> None:
    """Разворачивает окно ровно на рабочую область: панель задач остаётся видимой.

    Для безрамочного окна (overrideredirect) системная максимизация
    SW_MAXIMIZE накрывает панель задач, поэтому задаём геометрию сами —
    на ВСЕХ платформах (2.20.0): прежняя ветка state('zoomed') на
    не-Windows бездействовала для overrideredirect-окна и не давала
    тестам проверять разворачивание. Не удалось посчитать рабочую
    область — системный аналог.
    """
    try:
        x, y, width, height = workarea(window)
        if width > 0 and height > 0:
            window.geometry(f"{width}x{height}+{x}+{y}")
            return
    except Exception as exc:  # noqa: BLE001 — есть системный fallback
        _log_win32_failure("maximize_window(workarea)", exc)
    try:
        window.state("zoomed")
    except Exception:
        pass  # tk-fallback: не Win32, тишина допустима
def restore_window(window: tk.Misc, geometry: str = "") -> None:
    """Возврат из максимизации: применяет прежнюю геометрию, если она задана.

    Явная геометрия применяется на ВСЕХ платформах (2.20.0): раньше
    не-Windows-ветка её игнорировала (только state('normal')) — возврат
    из разворота/прилипания не возвращал размер окна, что вскрыли
    smoke-проверки прилипания. Без геометрии — системный аналог
    (SW_RESTORE на Windows, state('normal') на не-Windows).
    """
    if geometry:
        try:
            window.geometry(geometry)
            return
        except Exception:
            pass  # tk-fallback: не Win32, тишина допустима
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.ShowWindow(_main_hwnd(window), 9)  # SW_RESTORE
            return
        except Exception as exc:  # noqa: BLE001 — есть fallback state('normal')
            _log_win32_failure("ShowWindow(SW_RESTORE)", exc)
    try:
        window.state("normal")
    except Exception:
        pass  # tk-fallback: не Win32, тишина допустима


def _get_system_metrics(index: int) -> int:
    """Точка подмены в тестах: sysmetrics без ctypes-магии в тесте."""
    return int(ctypes.windll.user32.GetSystemMetrics(index))


def virtual_screen_bounds() -> tuple[int, int, int, int]:
    """Границы виртуального экрана (объединение всех мониторов): x, y, w, h.

    Ручной перенос (2.19.3 — единственный путь, нативный модальный цикл
    убран как причина вылетов) зажимает позицию окна в эти границы:
    окно можно унести на ЛЮБОЙ монитор — левый, правый, верхний, нижний
    (см. 2.17.0: прежние зажимы x ≥ −40, y ≥ 0 резали левый/верхний),
    но окно не теряется за пределами экранов — у края остаётся полоса
    DRAG_EDGE_MARGIN, за которую окно можно схватить.

    Не-Windows и сбой Win32: размеры экрана Tk, крайний fallback 1920×1080.
    """
    if sys.platform == "win32":
        try:
            x = _get_system_metrics(_SM_XVIRTUALSCREEN)
            y = _get_system_metrics(_SM_YVIRTUALSCREEN)
            width = _get_system_metrics(_SM_CXVIRTUALSCREEN)
            height = _get_system_metrics(_SM_CYVIRTUALSCREEN)
            if width > 0 and height > 0:
                return x, y, width, height
        except Exception as exc:  # noqa: BLE001 — есть экран Tk и fallback
            _log_win32_failure("GetSystemMetrics(SM_*VIRTUALSCREEN)", exc)
    try:
        source = getattr(tk, "_default_root", None)
        if source is not None:
            return 0, 0, source.winfo_screenwidth(), source.winfo_screenheight()
    except Exception:
        pass  # tk-fallback: не Win32, тишина допустима
    return 0, 0, 1920, 1080


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
    except Exception as exc:  # noqa: BLE001 — легально: режим уже установлен
        _log_win32_failure("SetProcessDpiAwareness", exc)
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception as exc2:  # noqa: BLE001 — системный DPI по умолчанию
            _log_win32_failure("SetProcessDPIAware", exc2)


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
        except Exception as exc:  # noqa: BLE001 — есть USERPROFILE-fallback
            _log_win32_failure("SHGetKnownFolderPath", exc)
        profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        return os.path.join(profile, "Documents")
    return os.path.join(os.path.expanduser("~"), "Documents")


def _drag_files(shell32, wparam: int) -> list[str]:
    """Читает список файлов из HDROP; DragFinish — ВСЕГДА (D7, WINLOG-1).

    DragFinish освобождает системную память списка файлов: без него при
    каждом сбое DragQueryFileW HDROP утекал (D7). Раньше DragFinish
    стоял в середине try — при исключении чтения он не выполнялся.
    Здесь он в finally: освобождение гарантировано при любом исходе,
    включая сбой чтения имени файла посреди списка.
    """
    files: list[str] = []
    try:
        count = shell32.DragQueryFileW(ctypes.c_void_p(wparam), 0xFFFFFFFF, None, 0)
        for index in range(count):
            length = shell32.DragQueryFileW(ctypes.c_void_p(wparam), index, None, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            shell32.DragQueryFileW(ctypes.c_void_p(wparam), index, buffer, length + 1)
            if buffer.value:
                files.append(buffer.value)
        return files
    finally:
        try:
            shell32.DragFinish(ctypes.c_void_p(wparam))
        except Exception:  # noqa: BLE001 — освобождение не роняет хендлер
            pass


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
                    files = _drag_files(shell32, wparam)
                    if files:
                        window.after(0, lambda f=list(files): on_files(f))
                    return 0
                except Exception as exc:  # noqa: BLE001 — хук не роняет окно
                    _log_win32_failure("DragQueryFileW(WM_DROPFILES)", exc)
            if msg == WM_DPICHANGED and on_dpi_changed is not None:
                try:
                    dpi = (int(wparam) >> 16) & 0xFFFF
                    if dpi:
                        rect = suggested_rect_from_lparam(lparam)
                        window.after(0, lambda d=dpi, r=rect: on_dpi_changed(d, r))
                        return 0
                except Exception as exc:  # noqa: BLE001 — хук не роняет окно
                    _log_win32_failure("WM_DPICHANGED", exc)
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
    except Exception as exc:  # noqa: BLE001 — хуки не обязательны для работы
        _log_win32_failure("install_window_hooks", exc)
