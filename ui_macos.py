"""Элементы интерфейса в стиле macOS для Tkinter + customtkinter.

Никаких внешних ассетов: «светофор», иконки, карточки, шторки и алерты
рисуются кодом, поэтому всё это работает и внутри собранного PyInstaller .exe.
"""

from __future__ import annotations

import ctypes
import sys
import tkinter as tk
from typing import Callable, Optional, Sequence

import customtkinter as ctk

# --------------------------------------------------------------------------- #
#  Шрифты
# --------------------------------------------------------------------------- #
_FAMILIES: set[str] = set()
FAMILY_TEXT = "Segoe UI"
FAMILY_DISPLAY = "Segoe UI"

#: Коэффициент масштаба интерфейса (1.0 = 96 dpi). Задаётся приложением при старте.
SCALE = 1.0


def set_scale(value: float) -> None:
    global SCALE
    SCALE = max(1.0, float(value))


def px(value: float) -> int:
    """Логические пиксели -> физические."""
    return int(round(value * SCALE))


def _pick(candidates: Sequence[str]) -> str:
    for name in candidates:
        if name.lower() in _FAMILIES:
            return name
    return "Segoe UI"


def init_fonts() -> None:
    """Подбирает максимально близкие к SF Pro шрифты (вызывать после создания окна)."""
    global _FAMILIES, FAMILY_TEXT, FAMILY_DISPLAY
    try:
        import tkinter.font as tkfont

        _FAMILIES = {name.lower() for name in tkfont.families()}
    except Exception:  # pragma: no cover - окружение без шрифтов
        _FAMILIES = set()

    FAMILY_TEXT = _pick(
        ("SF Pro Text", "SF Pro Display", "Inter", "Segoe UI Variable Text", "Segoe UI")
    )
    FAMILY_DISPLAY = _pick(
        ("SF Pro Display", "SF Pro Text", "Inter", "Segoe UI Variable Display", "Segoe UI")
    )


def font(size: int = 13, weight: str = "normal", display: bool = False) -> tuple:
    """Шрифт в физических пикселях (Tk понимает отрицательный размер как px)."""
    family = FAMILY_DISPLAY if display else FAMILY_TEXT
    size_px = max(8, int(round(size * SCALE)))
    if weight == "normal":
        return (family, -size_px)
    return (family, -size_px, weight)


# --------------------------------------------------------------------------- #
#  Палитра (пара «светлая тема / тёмная тема»)
# --------------------------------------------------------------------------- #
COLORS: dict[str, tuple[str, str]] = {
    "window": ("#FFFFFF", "#1E1E1F"),
    "chrome": ("#F5F5F7", "#2B2B2E"),
    "chrome_text": ("#1D1D1F", "#F2F2F7"),
    "chrome_line": ("#DFDFE5", "#3A3A3E"),
    "sidebar": ("#EAEAEF", "#26262A"),
    "sidebar_hover": ("#E1E1E8", "#303035"),
    "sidebar_active": ("#D7D7E0", "#3C3C42"),
    "sidebar_text": ("#54545A", "#A2A2A9"),
    "sidebar_text_active": ("#1D1D1F", "#FFFFFF"),
    "content": ("#F6F6F8", "#1B1B1D"),
    "card": ("#FFFFFF", "#2A2A2D"),
    "card_border": ("#E6E6EC", "#39393D"),
    "text": ("#1D1D1F", "#F2F2F7"),
    "text_secondary": ("#6E6E73", "#A1A1A8"),
    "text_tertiary": ("#96969C", "#76767D"),
    "entry_bg": ("#FFFFFF", "#3A3A3E"),
    "entry_border": ("#D6D6DD", "#4A4A4F"),
    "button_secondary": ("#FFFFFF", "#323236"),
    "button_secondary_hover": ("#F1F1F5", "#3C3C41"),
    "button_secondary_border": ("#D8D8DF", "#4A4A50"),
    "button_ghost_hover": ("#E9E9EF", "#3A3A3F"),
    "segment_track": ("#E9E9EE", "#333337"),
    "segment_selected": ("#FFFFFF", "#4C4C51"),
    "segment_hover": ("#DEDEE5", "#3C3C41"),
    "hairline": ("#E4E4EA", "#343438"),
    "accent": ("#007AFF", "#0A84FF"),
    "accent_hover": ("#2E93FF", "#3B9BFF"),
    "on_accent": ("#FFFFFF", "#FFFFFF"),
    "scroll": ("#D6D6DC", "#4A4A4F"),
    "scroll_hover": ("#C2C2C9", "#5A5A60"),
    "caption_hover": ("#E6E6E6", "#333336"),
    "caption_pressed": ("#D9D9D9", "#3F3F41"),
    "error_border": ("#E03131", "#FF5C5C"),
    "error_text": ("#C92A2A", "#FF8787"),
}

LIGHT_RED = "#FF5F57"
LIGHT_YELLOW = "#FEBC2E"
LIGHT_GREEN = "#28C840"

_LISTENERS: list[Callable[[], None]] = []


def resolve(color):
    """Превращает пару (light, dark) в конкретный цвет текущей темы."""
    if isinstance(color, (tuple, list)):
        return color[1] if ctk.get_appearance_mode().lower() == "dark" else color[0]
    return color


def register_theme_listener(callback: Callable[[], None]) -> None:
    """Регистрирует перерисовку для виджетов, которые рисуют себя сами (canvas)."""
    _LISTENERS.append(callback)


def set_appearance_mode(mode: str) -> None:
    ctk.set_appearance_mode(mode)
    for callback in list(_LISTENERS):
        try:
            callback()
        except Exception:  # pragma: no cover - виджет мог быть уже уничтожен
            pass


def toggle_appearance() -> str:
    mode = "light" if ctk.get_appearance_mode().lower() == "dark" else "dark"
    set_appearance_mode(mode)
    return mode


# --------------------------------------------------------------------------- #
#  Win32: безрамочное окно, скруглённые углы, рабочая область
# --------------------------------------------------------------------------- #
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


def workarea() -> tuple[int, int, int, int]:
    """Рабочая область экрана (без панели задач): x, y, ширина, высота."""
    if sys.platform == "win32":
        rect = _RECT()
        try:
            if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
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
            x, y, width, height = workarea()
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


# --------------------------------------------------------------------------- #
#  Буфер обмена: Ctrl+C/V/X/Z/A на любой раскладке клавиатуры
# --------------------------------------------------------------------------- #
#: Коды клавиш Windows не зависят от раскладки, в отличие от keysym,
#: который на русской раскладке ломает стандартные привязки Tk.
_CLIPBOARD_KEYCODES = {
    67: "<<Copy>>",       # C
    86: "<<Paste>>",      # V
    88: "<<Cut>>",        # X
    65: "<<SelectAll>>",  # A
    90: "<<Undo>>",       # Z
    89: "<<Redo>>",       # Y
    45: "<<Copy>>",       # Ctrl+Insert
}


def fix_clipboard_shortcuts(widget: tk.Misc) -> None:
    """Чинит Ctrl+V (и соседние сочетания) на русской и других раскладках."""
    def _on_ctrl_key(event):
        action = _CLIPBOARD_KEYCODES.get(int(event.keycode))
        if action:
            # Важно: генерируем виртуальное событие на реальном виджете
            # (event.widget), а не на обёртке (CTkEntry/CTkTextbox) — классовые
            # обработчики <<Paste>>, <<Copy>> и т.д. висят на внутренних
            # tk.Entry/tk.Text, и только они умеют вставлять из буфера.
            try:
                event.widget.event_generate(action)
            except Exception:
                pass
            return "break"
        return None

    widget.bind("<Control-KeyPress>", _on_ctrl_key, add="+")


def enable_sanitized_paste(entry: tk.Widget, sanitize: Callable[[str], str],
                           max_len: Optional[int] = None) -> None:
    """Вставка с очисткой: из буфера берутся только допустимые символы.

    С validate="key" стандартный <<Paste>> молча отбрасывает вставку целиком,
    если в буфере есть хоть один недопустимый символ («р/с 40702810…», номер
    с пробелами). Вместо отклонения вставляем очищенное значение.
    """
    def _paste(_event=None) -> str:
        try:
            raw = entry.clipboard_get()
        except Exception:
            return "break"
        # Заменяем выделенный фрагмент (если он есть) вставкой.
        try:
            entry.delete("sel.first", "sel.last")
        except Exception:
            pass
        text = sanitize(str(raw))
        if max_len is not None:
            room = max_len - len(entry.get())
            if room <= 0:
                return "break"
            text = text[:room]
        entry.insert("insert", text)
        return "break"

    entry.bind("<<Paste>>", _paste, add="+")


# --------------------------------------------------------------------------- #
#  Примитивы рисования
# --------------------------------------------------------------------------- #
def rounded_rect(canvas: tk.Canvas, x0: float, y0: float, x1: float, y1: float, radius: float, **kwargs):
    r = max(0.0, min(radius, (x1 - x0) / 2, (y1 - y0) / 2))
    points = [
        x0 + r, y0, x1 - r, y0, x1, y0,
        x1, y0 + r, x1, y1 - r, x1, y1,
        x1 - r, y1, x0 + r, y1, x0, y1,
        x0, y1 - r, x0, y0 + r, x0, y0,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=18, **kwargs)


# --------------------------------------------------------------------------- #
#  Иконки (рисуются линиями, 18x18)
# --------------------------------------------------------------------------- #
def icon_person(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    k = size / 18.0
    canvas.create_oval(
        x + 6.2 * k, y + 3.0 * k, x + 11.8 * k, y + 8.6 * k, outline=color, width=width
    )
    canvas.create_arc(
        x + 3.0 * k, y + 8.4 * k, x + 15.0 * k, y + 20.0 * k,
        start=180, extent=180, style="arc", outline=color, width=width,
    )


def icon_doc(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    k = size / 18.0
    canvas.create_line(
        x + 4.5 * k, y + 2.5 * k, x + 13.5 * k, y + 2.5 * k, x + 13.5 * k, y + 15.5 * k,
        x + 4.5 * k, y + 15.5 * k, x + 4.5 * k, y + 2.5 * k,
        fill=color, width=width, joinstyle="round",
    )
    for offset in (6.6, 9.0, 11.4):
        canvas.create_line(
            x + 7.0 * k, y + offset * k, x + 11.0 * k, y + offset * k, fill=color, width=width
        )


def icon_bank(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    k = size / 18.0
    canvas.create_line(
        x + 2.8 * k, y + 6.4 * k, x + 9.0 * k, y + 2.6 * k, x + 15.2 * k, y + 6.4 * k,
        fill=color, width=width, joinstyle="round",
    )
    for xpos in (4.6, 9.0, 13.4):
        canvas.create_line(
            x + xpos * k, y + 7.6 * k, x + xpos * k, y + 13.2 * k, fill=color, width=width
        )
    canvas.create_line(
        x + 3.2 * k, y + 15.0 * k, x + 14.8 * k, y + 15.0 * k, fill=color, width=width
    )


def icon_pencil(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    k = size / 18.0
    canvas.create_line(
        x + 3.0 * k, y + 15.0 * k, x + 4.6 * k, y + 10.2 * k, x + 12.4 * k, y + 2.4 * k,
        x + 15.4 * k, y + 5.4 * k, x + 7.6 * k, y + 13.2 * k, x + 3.0 * k, y + 15.0 * k,
        fill=color, width=width, joinstyle="round",
    )


def icon_registry(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    """Реестр/список: три строки с маркерами слева."""
    k = size / 18.0
    for offset in (4.2, 9.0, 13.8):
        canvas.create_oval(
            x + 3.2 * k, y + (offset - 0.9) * k, x + 5.0 * k, y + (offset + 0.9) * k,
            outline=color, width=width,
        )
        canvas.create_line(
            x + 7.2 * k, y + offset * k, x + 14.8 * k, y + offset * k,
            fill=color, width=width,
        )


def icon_gear(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    """Настройки: шестерёнка — окружность с зубцами по кругу."""
    k = size / 18.0
    canvas.create_oval(
        x + 5.6 * k, y + 5.6 * k, x + 12.4 * k, y + 12.4 * k, outline=color, width=width
    )
    for dx, dy in (
        (0, -1), (0, 1), (-1, 0), (1, 0),
        (-0.7, -0.7), (0.7, -0.7), (-0.7, 0.7), (0.7, 0.7),
    ):
        canvas.create_line(
            x + 9.0 * k + dx * 6.2 * k, y + 9.0 * k + dy * 6.2 * k,
            x + 9.0 * k + dx * 8.6 * k, y + 9.0 * k + dy * 8.6 * k,
            fill=color, width=width, capstyle="round",
        )


def icon_folder(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    """Папка: контур с «язычком» слева."""
    k = size / 18.0
    canvas.create_line(
        x + 3.0 * k, y + 14.5 * k, x + 3.0 * k, y + 4.5 * k, x + 7.6 * k, y + 4.5 * k,
        x + 9.6 * k, y + 6.6 * k, x + 15.0 * k, y + 6.6 * k, x + 15.0 * k, y + 14.5 * k,
        x + 3.0 * k, y + 14.5 * k,
        fill=color, width=width, joinstyle="round",
    )


def icon_sun(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18, width: float = 1.6):
    """Солнце: круг с восемью лучами."""
    k = size / 18.0
    canvas.create_oval(
        x + 6.1 * k, y + 6.1 * k, x + 11.9 * k, y + 11.9 * k, outline=color, width=width
    )
    for dx, dy in (
        (0, -1), (0, 1), (-1, 0), (1, 0),
        (-0.71, -0.71), (0.71, -0.71), (-0.71, 0.71), (0.71, 0.71),
    ):
        canvas.create_line(
            x + 9.0 * k + dx * 5.2 * k, y + 9.0 * k + dy * 5.2 * k,
            x + 9.0 * k + dx * 7.6 * k, y + 9.0 * k + dy * 7.6 * k,
            fill=color, width=width, capstyle="round",
        )


def icon_moon(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18,
              width: float = 1.6, mask: Optional[str] = None):
    """Луна: filled-серп; «вырез» закрашивается цветом фона (mask)."""
    k = size / 18.0
    canvas.create_oval(x + 3.6 * k, y + 2.8 * k, x + 14.4 * k, y + 15.2 * k, fill=color, outline="")
    canvas.create_oval(
        x + 7.0 * k, y + 0.6 * k, x + 17.6 * k, y + 11.2 * k,
        fill=mask if mask is not None else canvas["bg"], outline="",
    )


def icon_plus(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18,
              width: float = 1.6, mask: Optional[str] = None):
    """Плюс: добавить строку."""
    k = size / 18.0
    caps = {"capstyle": "round"}
    canvas.create_line(x + 9 * k, y + 3.6 * k, x + 9 * k, y + 14.4 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 3.6 * k, y + 9 * k, x + 14.4 * k, y + 9 * k,
                       fill=color, width=width, **caps)


def icon_minus(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18,
               width: float = 1.6, mask: Optional[str] = None):
    """Минус: убрать строку."""
    k = size / 18.0
    canvas.create_line(x + 3.6 * k, y + 9 * k, x + 14.4 * k, y + 9 * k,
                       fill=color, width=width, capstyle="round")


def icon_trash(canvas: tk.Canvas, x: float, y: float, color: str, size: float = 18,
               width: float = 1.6, mask: Optional[str] = None):
    """Корзина: очистить таблицу."""
    k = size / 18.0
    caps = {"capstyle": "round"}
    # крышка
    canvas.create_line(x + 3 * k, y + 5 * k, x + 15 * k, y + 5 * k,
                       fill=color, width=width, **caps)
    # ручка крышки
    canvas.create_line(x + 7 * k, y + 5 * k, x + 7 * k, y + 3 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 7 * k, y + 3 * k, x + 11 * k, y + 3 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 11 * k, y + 3 * k, x + 11 * k, y + 5 * k,
                       fill=color, width=width, **caps)
    # корпус
    canvas.create_line(x + 4.8 * k, y + 5.5 * k, x + 5.6 * k, y + 14.6 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 5.6 * k, y + 14.6 * k, x + 12.4 * k, y + 14.6 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 12.4 * k, y + 14.6 * k, x + 13.2 * k, y + 5.5 * k,
                       fill=color, width=width, **caps)
    # «мусор» внутри
    canvas.create_line(x + 7.3 * k, y + 7.8 * k, x + 7.6 * k, y + 12.2 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 9 * k, y + 7.8 * k, x + 9 * k, y + 12.2 * k,
                       fill=color, width=width, **caps)
    canvas.create_line(x + 10.7 * k, y + 7.8 * k, x + 10.4 * k, y + 12.2 * k,
                       fill=color, width=width, **caps)


ICONS = {
    "person": icon_person,
    "doc": icon_doc,
    "bank": icon_bank,
    "pencil": icon_pencil,
    "list": icon_registry,
    "gear": icon_gear,
    "folder": icon_folder,
    "sun": icon_sun,
    "moon": icon_moon,
    "plus": icon_plus,
    "minus": icon_minus,
    "trash": icon_trash,
}


# --------------------------------------------------------------------------- #
#  «Светофор» macOS
# --------------------------------------------------------------------------- #
class TrafficLights(tk.Canvas):
    DIAMETER = 12
    GAP = 8

    def __init__(self, master, on_close=None, on_minimize=None, on_zoom=None):
        self._d = px(self.DIAMETER)
        self._gap = px(self.GAP)
        width = 3 * self._d + 2 * self._gap
        super().__init__(
            master,
            width=width,
            height=self._d + px(4),
            bg=resolve(COLORS["chrome"]),
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self._callbacks = (on_close, on_minimize, on_zoom)
        self._dots: list[tuple[float, float]] = []
        for index in range(3):
            x = index * (self._d + self._gap)
            y = px(2)
            self._dots.append((x, y))
            dot = self.create_oval(
                x, y, x + self._d, y + self._d,
                fill=(LIGHT_RED, LIGHT_YELLOW, LIGHT_GREEN)[index],
                outline="",
                tags=(f"dot{index}", "dot"),
            )
            self.tag_bind(dot, "<Button-1>", lambda _event, i=index: self._invoke(i))

        self._glyph_items: list[int] = []
        self.bind("<Enter>", lambda _event: self._show_glyphs(True))
        self.bind("<Leave>", lambda _event: self._show_glyphs(False))
        register_theme_listener(self._sync_theme)

    # -- внешний вид ------------------------------------------------------ #
    def _sync_theme(self) -> None:
        try:
            self.configure(bg=resolve(COLORS["chrome"]))
        except Exception:
            return
        if self._glyph_items:
            self._show_glyphs(False)
            self._show_glyphs(True)

    def _show_glyphs(self, visible: bool) -> None:
        for item in self._glyph_items:
            try:
                self.delete(item)
            except Exception:
                pass
        self._glyph_items.clear()
        if not visible:
            return

        glyph_color = "#3B3B3B"
        line = max(1.0, px(1.4))
        arm = px(2.6)
        for (x, y), index in zip(self._dots, range(3)):
            cx, cy = x + self._d / 2, y + self._d / 2
            if index == 0:
                self._glyph_items.append(
                    self.create_line(cx - arm, cy - arm, cx + arm, cy + arm, fill=glyph_color, width=line, capstyle="round")
                )
                self._glyph_items.append(
                    self.create_line(cx - arm, cy + arm, cx + arm, cy - arm, fill=glyph_color, width=line, capstyle="round")
                )
            elif index == 1:
                self._glyph_items.append(
                    self.create_line(cx - arm, cy, cx + arm, cy, fill=glyph_color, width=line, capstyle="round")
                )
            else:
                self._glyph_items.append(
                    self.create_line(cx - arm, cy, cx + arm, cy, fill=glyph_color, width=line, capstyle="round")
                )
                self._glyph_items.append(
                    self.create_line(cx, cy - arm, cx, cy + arm, fill=glyph_color, width=line, capstyle="round")
                )

    def _invoke(self, index: int) -> None:
        callback = self._callbacks[index]
        if callable(callback):
            callback()


# --------------------------------------------------------------------------- #
#  Кнопки окна в стиле Windows 11: свернуть / развернуть / закрыть
# --------------------------------------------------------------------------- #
class WindowControls(tk.Canvas):
    """Плоские caption-кнопки (35×24) в правом верхнем углу окна.

    Свернуть и развернуть при наведении подсвечиваются серым, закрыть —
    красным (#C42B1C) с белым крестиком, как в Windows 11.
    """

    BUTTON_W = 35
    BUTTON_H = 24
    #: Кнопки на четверть меньше типовых 46×32.
    SIZE_SCALE = 0.75
    CLOSE_RED = "#C42B1C"
    CLOSE_RED_PRESSED = "#B0271E"

    def __init__(self, master, on_close=None, on_minimize=None, on_zoom=None,
                 height: Optional[int] = None):
        self._bw = px(self.BUTTON_W)
        # Высота по умолчанию — «плоские» кнопки; если передана высота шапки
        # (уже в пикселях), кнопки занимают 3/4 её высоты и центрируются
        # по вертикали при упаковке.
        self._bh = int(int(height) * self.SIZE_SCALE) if height else px(self.BUTTON_H)
        super().__init__(
            master,
            width=3 * self._bw,
            height=self._bh,
            bg=resolve(COLORS["chrome"]),
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self._callbacks = (on_minimize, on_zoom, on_close)
        self._hover = -1
        self._pressed = -1
        self._zoomed = False
        self.bind("<Motion>", self._on_motion, add="+")
        self.bind("<Leave>", lambda _event: self._set_hover(-1), add="+")
        self.bind("<ButtonPress-1>", self._on_press, add="+")
        self.bind("<ButtonRelease-1>", self._on_release, add="+")
        register_theme_listener(self._sync_theme)
        self._draw()

    # -- состояния -------------------------------------------------------- #
    def _button_at(self, x: float) -> int:
        return min(2, max(0, int(x // self._bw)))

    def _set_hover(self, index: int) -> None:
        if index != self._hover:
            self._hover = index
            self._draw()

    def _on_motion(self, event) -> None:
        self._set_hover(self._button_at(event.x))

    def _on_press(self, event) -> None:
        self._pressed = self._button_at(event.x)
        self._draw()

    def _on_release(self, event) -> None:
        index = self._button_at(event.x)
        pressed, self._pressed = self._pressed, -1
        self._draw()
        if pressed == index and callable(self._callbacks[index]):
            self._callbacks[index]()

    def set_zoomed(self, zoomed: bool) -> None:
        """Меняет глиф средней кнопки: «развернуть» <-> «восстановить»."""
        if self._zoomed == bool(zoomed):
            return
        self._zoomed = bool(zoomed)
        self._draw()

    # -- отрисовка -------------------------------------------------------- #
    def _sync_theme(self) -> None:
        try:
            self.configure(bg=resolve(COLORS["chrome"]))
        except Exception:
            return
        self._draw()

    def _bg(self, index: int) -> str:
        if index == 2:  # закрыть
            if self._pressed == 2:
                return self.CLOSE_RED_PRESSED
            if self._hover == 2:
                return self.CLOSE_RED
        else:
            if self._pressed == index:
                return resolve(COLORS["caption_pressed"])
            if self._hover == index:
                return resolve(COLORS["caption_hover"])
        return resolve(COLORS["chrome"])

    def _glyph_color(self, index: int) -> str:
        if index == 2 and self._hover == 2:
            return "#FFFFFF"
        return resolve(COLORS["chrome_text"])

    def _draw(self) -> None:
        self.delete("all")
        stroke = max(1.0, px(1.0))
        arm = px(4)  # половина «размаха» глифа
        for index in range(3):
            x0 = index * self._bw
            self.create_rectangle(
                x0, 0, x0 + self._bw, self._bh, fill=self._bg(index), outline=""
            )
            cx, cy = x0 + self._bw / 2, self._bh / 2
            color = self._glyph_color(index)
            if index == 0:  # свернуть — горизонтальная черта
                self.create_line(cx - arm, cy, cx + arm, cy, fill=color, width=stroke)
            elif index == 1:  # развернуть — квадрат; восстановить — две рамки
                s = px(7)
                if self._zoomed:
                    off = px(2)
                    self.create_rectangle(
                        cx - s / 2 + off, cy - s / 2 - off, cx + s / 2 + off, cy + s / 2 - off,
                        outline=color, width=stroke,
                    )
                    self.create_rectangle(
                        cx - s / 2 - off, cy - s / 2 + off, cx + s / 2 - off, cy + s / 2 + off,
                        fill=self._bg(index), outline=color, width=stroke,
                    )
                else:
                    self.create_rectangle(
                        cx - s / 2, cy - s / 2, cx + s / 2, cy + s / 2,
                        outline=color, width=stroke,
                    )
            else:  # закрыть — крест
                self.create_line(cx - arm, cy - arm, cx + arm, cy + arm, fill=color, width=stroke)
                self.create_line(cx - arm, cy + arm, cx + arm, cy - arm, fill=color, width=stroke)


# --------------------------------------------------------------------------- #
#  Перетаскивание окна за шапку
# --------------------------------------------------------------------------- #
class WindowDrag:
    def __init__(self, window: tk.Misc, on_start: Optional[Callable[[], None]] = None):
        self._window = window
        self._on_start = on_start
        self._offset: Optional[tuple[int, int]] = None

    def attach(self, *widgets: tk.Misc) -> None:
        for widget in widgets:
            widget.bind("<ButtonPress-1>", self._press, add="+")
            widget.bind("<B1-Motion>", self._move, add="+")
            widget.bind("<ButtonRelease-1>", self._release, add="+")

    def _press(self, event) -> None:
        if self._on_start is not None:
            self._on_start(event)
        try:
            # Геометрия могла измениться в on_start (восстановление из
            # максимизации) — пересчитываем окно, чтобы оно не «убегало».
            self._window.update_idletasks()
            self._offset = (event.x_root - self._window.winfo_x(), event.y_root - self._window.winfo_y())
        except Exception:
            self._offset = None

    def _move(self, event) -> None:
        if self._offset is None:
            return
        x = event.x_root - self._offset[0]
        y = event.y_root - self._offset[1]
        self._window.geometry(f"+{max(x, -40)}+{max(y, 0)}")

    def _release(self, _event) -> None:
        self._offset = None


# --------------------------------------------------------------------------- #
#  Пункт бокового меню
# --------------------------------------------------------------------------- #
class NavItem(tk.Canvas):
    def __init__(self, master, text: str, icon: str, command: Callable[[], None], width: int = 186, height: int = 32):
        super().__init__(
            master, width=px(width), height=px(height), highlightthickness=0, bd=0,
            bg=resolve(COLORS["sidebar"]), cursor="hand2",
        )
        self._text = text
        self._icon = icon
        self._command = command
        self._selected = False
        self._hover = False
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda _event: self._command())
        self._draw()
        register_theme_listener(self._draw)

    def set_selected(self, selected: bool) -> None:
        if self._selected == selected:
            return
        self._selected = selected
        self._draw()

    def _on_enter(self, _event) -> None:
        self._hover = True
        self._draw()

    def _on_leave(self, _event) -> None:
        self._hover = False
        self._draw()

    def _draw(self) -> None:
        try:
            self.configure(bg=resolve(COLORS["sidebar"]))
        except Exception:
            return
        self.delete("all")
        width = int(self["width"])
        height = int(self["height"])

        if self._selected:
            rounded_rect(self, 1, 1, width - 1, height - 1, px(8), fill=resolve(COLORS["sidebar_active"]), outline="")
        elif self._hover:
            rounded_rect(self, 1, 1, width - 1, height - 1, px(8), fill=resolve(COLORS["sidebar_hover"]), outline="")

        if self._selected:
            icon_color = resolve(COLORS["accent"])
            text_color = resolve(COLORS["sidebar_text_active"])
        else:
            icon_color = resolve(COLORS["sidebar_text"])
            text_color = resolve(COLORS["sidebar_text"])

        ICONS[self._icon](self, px(13), height / 2 - px(9), icon_color, px(18), max(1.3, px(1.5)))
        self.create_text(
            px(41), height / 2, text=self._text, anchor="w", fill=text_color,
            font=font(13, "bold" if self._selected else "normal"),
        )


# --------------------------------------------------------------------------- #
#  Кнопка раздела в шапке окна
# --------------------------------------------------------------------------- #
class TopNavItem(tk.Canvas):
    """Кнопка верхней панели: иконка + подпись, единый стиль для всех разделов."""

    def __init__(self, master, text: str, icon: str, command: Callable[[], None], height: int = 30):
        import tkinter.font as tkfont

        self._text = text
        self._icon = icon
        self._command = command
        self._selected = False
        self._hover = False
        text_width = tkfont.Font(font=font(12)).measure(text)
        width = px(12) + px(18) + px(7) + text_width + px(13)
        super().__init__(
            master, width=width, height=px(height), highlightthickness=0, bd=0,
            bg=resolve(COLORS["chrome"]), cursor="hand2",
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda _event: self._command())
        self._draw()
        register_theme_listener(self._draw)

    def set_selected(self, selected: bool) -> None:
        if self._selected == selected:
            return
        self._selected = selected
        self._draw()

    def _on_enter(self, _event) -> None:
        self._hover = True
        self._draw()

    def _on_leave(self, _event) -> None:
        self._hover = False
        self._draw()

    def _draw(self) -> None:
        try:
            self.configure(bg=resolve(COLORS["chrome"]))
        except Exception:
            return
        self.delete("all")
        width = int(self["width"])
        height = int(self["height"])

        if self._selected:
            rounded_rect(
                self, 1, 1, width - 1, height - 1, px(8),
                fill=resolve(COLORS["segment_selected"]),
                outline=resolve(COLORS["card_border"]),
            )
        elif self._hover:
            rounded_rect(
                self, 1, 1, width - 1, height - 1, px(8),
                fill=resolve(COLORS["segment_hover"]), outline="",
            )

        icon_color = resolve(COLORS["accent"]) if self._selected else resolve(COLORS["chrome_text"])
        ICONS[self._icon](self, px(12), height / 2 - px(9), icon_color, px(18), max(1.3, px(1.5)))
        self.create_text(
            px(12) + px(18) + px(7), height / 2, text=self._text, anchor="w",
            fill=resolve(COLORS["chrome_text"]),
            font=font(12, "bold" if self._selected else "normal"),
        )


# --------------------------------------------------------------------------- #
#  Карточка с полями
# --------------------------------------------------------------------------- #
class Card(ctk.CTkFrame):
    def __init__(self, master, title: Optional[str] = None,
                 labels_above: bool = False, compact: bool = False, **kwargs):
        super().__init__(
            master,
            fg_color=COLORS["card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["card_border"],
            **kwargs,
        )
        #: Подписи над полями (для узких колонок, где поле перекрывает текст).
        self._labels_above = bool(labels_above)
        #: Компактный режим: меньше отступов между строками (текст слева).
        self._compact = bool(compact)
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=16, pady=(13, 15))
        self.body.grid_columnconfigure(0, weight=1)
        self.body.grid_columnconfigure(1, weight=0)
        self._row = 0

        if title:
            ctk.CTkLabel(
                self.body,
                text=title,
                font=font(12, "bold"),
                text_color=COLORS["text_secondary"],
                anchor="w",
                height=18,
            ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 7))
            self._row = 1

    def section(self, title: str) -> None:
        """Заголовок секции внутри общей карточки (с тонким разделителем)."""
        if self._row > 0:
            ctk.CTkFrame(self.body, fg_color=COLORS["hairline"], corner_radius=0, height=1).grid(
                row=self._row, column=0, columnspan=2, sticky="we", pady=(10, 0),
            )
            self._row += 1
        ctk.CTkLabel(
            self.body, text=title, font=font(12, "bold"),
            text_color=COLORS["text_secondary"], anchor="w", height=18,
        ).grid(row=self._row, column=0, columnspan=2, sticky="w", pady=(6, 7))
        self._row += 1

    def row(
        self,
        label: str,
        widget: tk.Misc,
        hint: Optional[str] = None,
        stretch: bool = False,
        required: bool = False,
        error_var: Optional[tk.StringVar] = None,
    ) -> ctk.CTkLabel:
        if required:
            label = f"{label} *"
        if self._labels_above:
            # Подпись над полем: поле занимает всю ширину и не перекрывает текст.
            label_widget = ctk.CTkLabel(
                self.body, text=label, font=font(12),
                text_color=COLORS["text_secondary"], anchor="w",
            )
            label_widget.grid(row=self._row, column=0, columnspan=2, sticky="w", pady=(5, 1))
            self._row += 1
            widget.grid(
                row=self._row, column=0, columnspan=2,
                sticky="ew" if stretch else "w", pady=(0, 4),
            )
            if stretch:
                self.body.grid_columnconfigure(0, weight=1)
            self._row += 1
        else:
            pady = 2 if self._compact else 4
            padx = (10, 0) if self._compact else (14, 0)
            label_widget = ctk.CTkLabel(
                self.body, text=label, font=font(13), text_color=COLORS["text"], anchor="w",
            )
            label_widget.grid(row=self._row, column=0, sticky="w", pady=pady)
            widget.grid(
                row=self._row, column=1, sticky="ew" if stretch else "w",
                pady=pady, padx=padx,
            )
            if stretch:
                # Поле растягивается вместе с карточкой при изменении окна.
                self.body.grid_columnconfigure(1, weight=1)
            self._row += 1

        if error_var is not None:
            # Строка под полем для inline-подсказки об ошибке: пустая строка
            # не занимает места, ошибка появляется красным текстом.
            hint_label = tk.Label(
                self.body, textvariable=error_var, font=font(11), height=0, bd=0,
                highlightthickness=0, anchor="w", justify="left",
                fg=resolve(COLORS["error_text"]), bg=resolve(COLORS["card"]),
            )
            hint_label.grid(row=self._row, column=0, columnspan=2, sticky="w", pady=(0, 3))
            register_theme_listener(lambda: _sync_error_hint(hint_label))
            self._row += 1

        if hint:
            ctk.CTkLabel(
                self.body, text=hint, font=font(11), text_color=COLORS["text_tertiary"], anchor="w",
            ).grid(row=self._row, column=0, columnspan=2, sticky="w", pady=(0, 4))
            self._row += 1
        return label_widget

    def info_row(self, label: str) -> ctk.CTkLabel:
        """Строка «подпись — значение» только для чтения; возвращает подпись значения."""
        ctk.CTkLabel(
            self.body, text=label, font=font(13), text_color=COLORS["text_secondary"], anchor="w",
        ).grid(row=self._row, column=0, sticky="w", pady=2)
        value = ctk.CTkLabel(
            self.body, text="—", font=font(13, "bold"), text_color=COLORS["text"], anchor="w",
        )
        value.grid(row=self._row, column=1, sticky="w", pady=2, padx=(14, 0))
        self._row += 1
        return value

    def wide(self, widget: tk.Misc, pady: tuple[int, int] = (2, 2)) -> tk.Misc:
        widget.grid(row=self._row, column=0, columnspan=2, sticky="we", pady=pady)
        self._row += 1
        return widget


def _sync_error_hint(label: tk.Label) -> None:
    """Обновляет цвета inline-подсказки об ошибке при смене темы."""
    try:
        label.configure(fg=resolve(COLORS["error_text"]), bg=resolve(COLORS["card"]))
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  Фабрики контролов
# --------------------------------------------------------------------------- #
def entry(master, width: int = 300, height: int = 30, validatecommand=None, textvariable=None) -> "TextField":
    return TextField(master, width=width, height=height, validatecommand=validatecommand, textvariable=textvariable)


def option_menu(master, values: Sequence[str], width: int = 300, variable=None) -> "ComboField":
    return ComboField(master, values, width=width, variable=variable)


class TextField(ctk.CTkFrame):
    """Поле ввода с рамкой и внутренними отступами, как у NSTextField в macOS."""

    def __init__(self, master, width: int = 300, height: int = 30, validatecommand=None, textvariable=None):
        super().__init__(
            master,
            width=width,
            height=height,
            corner_radius=8,
            border_width=1,
            border_color=COLORS["entry_border"],
            fg_color=COLORS["entry_bg"],
        )
        self.pack_propagate(False)
        self.grid_propagate(False)

        self._padding = px(9)
        padding = self._padding
        kwargs = dict(
            width=width - padding - px(6),
            height=height - 2,
            corner_radius=0,
            border_width=0,
            fg_color=COLORS["entry_bg"],
            text_color=COLORS["text"],
            font=font(13),
        )
        if textvariable is not None:
            kwargs["textvariable"] = textvariable
        if validatecommand is not None:
            kwargs["validate"] = "key"
            kwargs["validatecommand"] = validatecommand

        self.entry = ctk.CTkEntry(self, **kwargs)
        self.entry.place(x=padding, y=1)
        # Ctrl+C/V/X/Z/A должны работать и на русской раскладке.
        fix_clipboard_shortcuts(self.entry)

        self._error = False
        self.entry.bind("<FocusIn>", lambda _event: self.configure(border_color=COLORS["accent"]), add="+")
        self.entry.bind("<FocusOut>", self._on_focus_out, add="+")
        # Внешний кадр может растягиваться (sticky="ew") — подтягиваем ввод.
        self.bind("<Configure>", self._on_resize, add="+")

    def _on_resize(self, event) -> None:
        try:
            self.entry.configure(width=max(40, int(event.width) - self._padding * 2))
        except Exception:
            pass

    def _on_focus_out(self, _event=None) -> None:
        self.configure(
            border_color=COLORS["error_border"] if self._error else COLORS["entry_border"]
        )

    def set_error(self, state: bool) -> None:
        """Подсвечивает рамку поля красным при ошибке валидации."""
        self._error = bool(state)
        self._on_focus_out()

    # -- делегирование наружу ------------------------------------------- #
    def get(self) -> str:
        return self.entry.get()

    def insert(self, index, text):
        return self.entry.insert(index, text)

    def delete(self, first, last=None):
        return self.entry.delete(first, last)

    def focus_set(self):
        return self.entry.focus_set()

    def bind_entry(self, sequence, command, add="+"):
        return self.entry.bind(sequence, command, add=add)


class _ComboPopup(tk.Toplevel):
    """Всплывающий список для ComboField: собственное окно-«поповер».

    Пункты подсвечиваются при наведении, выбранный выделен акцентом,
    список прокручивается колесом и управляется стрелками/Enter/Esc.
    Закрывается кликом вне списка, Escape или выбором пункта.
    """

    ROW_H = 26
    MAX_VISIBLE = 8
    MAX_ITEMS = 100  # выше — список обрезается с подсказкой (п. 5.4)

    def __init__(self, combo: "ComboField", values: Sequence[str], on_pick: Callable[[str], None]):
        super().__init__(combo)
        self.withdraw()
        self.overrideredirect(True)
        self._combo = combo
        self._values = [str(v) for v in values]
        self._shown = self._values[:self.MAX_ITEMS]
        self._truncated = len(self._values) > len(self._shown)
        self._on_pick = on_pick
        self._index = -1

        selected = combo.get()
        self._selected_index = self._shown.index(selected) if selected in self._shown else -1

        self._frame = tk.Frame(
            self, bg=resolve(COLORS["card"]),
            highlightthickness=1, highlightbackground=resolve(COLORS["card_border"]),
            highlightcolor=resolve(COLORS["card_border"]),
        )
        self._frame.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(
            self._frame, bg=resolve(COLORS["card"]), highlightthickness=0, bd=0,
        )
        self._canvas.pack(side="left", fill="both", expand=True)
        self._inner = tk.Frame(self._canvas, bg=resolve(COLORS["card"]))
        self._window_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind(
            "<Configure>",
            lambda _e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfigure(self._window_id, width=e.width),
        )

        self._rows: list[tk.Label] = []
        for index, value in enumerate(self._shown):
            row = tk.Label(
                self._inner, text=value, anchor="w", justify="left",
                padx=px(11), pady=px(4), font=font(12),
                bg=resolve(COLORS["card"]), fg=resolve(COLORS["text"]),
            )
            row.pack(fill="x")
            row.bind("<Enter>", lambda _e, i=index: self._set_index(i))
            row.bind("<Leave>", lambda _e: self._set_index(self._selected_index))
            row.bind(
                "<Button-1>",
                lambda _e, i=index: self._confirm(i),
            )
            self._rows.append(row)

        if self._truncated:
            tk.Label(
                self._inner,
                text=f"Показаны первые {len(self._shown)} из {len(self._values)}. "
                     "Используйте поиск по ИНН или реестр.",
                anchor="w", justify="left", padx=px(11), pady=px(6),
                font=font(11), fg=resolve(COLORS["text_tertiary"]),
                bg=resolve(COLORS["card"]), wraplength=px(300),
            ).pack(fill="x")

        # Колесо мыши и клавиатура
        for widget in (self, self._frame, self._canvas, self._inner, *self._rows):
            widget.bind("<MouseWheel>", self._on_wheel, add="+")
        self.bind("<Escape>", lambda _e: self.close())
        self.bind("<Return>", lambda _e: self._confirm(self._index))
        self.bind("<KP_Enter>", lambda _e: self._confirm(self._index))
        self.bind("<Up>", lambda _e: self._move(-1))
        self.bind("<Down>", lambda _e: self._move(1))
        # Клик вне списка (при grab все «чужие» клики приходят в это окно)
        self.bind("<Button-1>", self._on_click, add="+")

        self._set_index(self._selected_index)

    # -- внешний вид ------------------------------------------------------ #
    def _set_index(self, index: int) -> None:
        """Подсвечивает только прежнюю и новую строку — O(1) на движение (п. 5.4)."""
        previous = self._index
        self._index = index
        if previous == index:
            return
        if 0 <= previous < len(self._rows):
            self._style_row(previous)
        if 0 <= index < len(self._rows):
            self._style_row(index)

    def _style_row(self, index: int) -> None:
        row = self._rows[index]
        if index == self._index:
            row.configure(
                bg=resolve(COLORS["accent"]), fg=resolve(COLORS["on_accent"]),
                font=font(12, "bold"),
            )
        else:
            row.configure(
                bg=resolve(COLORS["card"]), fg=resolve(COLORS["text"]),
                font=font(12, "bold" if index == self._selected_index else "normal"),
            )

    # -- события ---------------------------------------------------------- #
    def _on_wheel(self, event) -> None:
        try:
            self._canvas.yview_scroll(-1 if int(event.delta) > 0 else 1, "units")
        except Exception:
            pass
        return "break"

    def _on_click(self, event) -> None:
        try:
            x0, y0 = self.winfo_rootx(), self.winfo_rooty()
            x1 = x0 + self.winfo_width()
            y1 = y0 + self.winfo_height()
        except Exception:
            self.close()
            return "break"
        if not (x0 <= event.x_root <= x1 and y0 <= event.y_root <= y1):
            self.close()
        return "break"

    def _move(self, delta: int) -> None:
        if not self._rows:
            return
        new_index = min(len(self._rows) - 1, max(0, self._index + delta))
        if new_index != self._index:
            self._set_index(new_index)
            self._ensure_visible(new_index)
        return "break"

    def _confirm(self, index: int) -> None:
        if 0 <= index < len(self._shown):
            self._on_pick(self._shown[index])
        self.close()

    def _ensure_visible(self, index: int) -> None:
        try:
            row_h = self._rows[0].winfo_reqheight() if self._rows else px(self.ROW_H)
            top = index * row_h
            self._canvas.update_idletasks()
            view_top = self._canvas.canvasy(0)
            view_bottom = self._canvas.canvasy(self._canvas.winfo_height())
            if top < view_top:
                self._canvas.yview_moveto(max(0.0, top / max(1.0, self._total_height())))
            elif top + row_h > view_bottom:
                self._canvas.yview_moveto(
                    max(0.0, min(1.0, (top + row_h) / max(1.0, self._total_height())))
                )
        except Exception:
            pass

    def _total_height(self) -> float:
        row_h = self._rows[0].winfo_reqheight() if self._rows else px(self.ROW_H)
        return row_h * max(1, len(self._rows))

    # -- показ/закрытие ---------------------------------------------------- #
    def show(self) -> None:
        combo = self._combo
        combo.update_idletasks()
        self.update_idletasks()

        row_h = self._rows[0].winfo_reqheight() if self._rows else px(self.ROW_H)
        visible = min(len(self._rows), self.MAX_VISIBLE)
        height = visible * row_h + px(2)

        import tkinter.font as tkfont

        text_font = tkfont.Font(font=font(12))
        text_w = max((text_font.measure(v) for v in self._shown), default=0)
        wa = workarea()
        width = min(max(combo.winfo_width(), text_w + px(34)), px(520), wa[2])

        x = combo.winfo_rootx()
        y = combo.winfo_rooty() + combo.winfo_height() + px(2)
        if y + height > wa[1] + wa[3]:
            above = combo.winfo_rooty() - height - px(2)
            y = above if above >= wa[1] else max(wa[1], wa[1] + wa[3] - height)
        x = max(wa[0], min(x, wa[0] + wa[2] - width))

        self.geometry(f"{width}x{height}+{x}+{y}")
        self.deiconify()
        self.lift()
        try:
            self.grab_set()
            self.focus_set()
        except Exception:
            pass

        # Прокрутка к выбранному пункту
        if self._selected_index > 0:
            total = self._total_height()
            viewport = max(1.0, height - px(2))
            offset = max(0.0, self._selected_index * row_h - viewport / 3)
            self._canvas.yview_moveto(max(0.0, min(1.0, offset / max(1.0, total))))

    def close(self) -> None:
        try:
            self._combo._close_popup()
        except Exception:
            pass
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()


class ComboField(ctk.CTkFrame):
    """Выпадающий список в такой же рамке, как у обычного поля.

    Открывается кликом по всему полю; список рисуется собственным
    окном-«поповером» (_ComboPopup) в стиле macOS вместо системного меню.
    """

    def __init__(self, master, values: Sequence[str], width: int = 300, height: int = 30, variable=None):
        super().__init__(
            master,
            width=width,
            height=height,
            corner_radius=8,
            border_width=1,
            border_color=COLORS["entry_border"],
            fg_color=COLORS["entry_bg"],
        )
        self.pack_propagate(False)
        self.grid_propagate(False)

        self._values = [str(v) for v in values]
        self._variable = variable
        self._text = ""
        self._popup = None

        self._label = ctk.CTkLabel(
            self, text="", font=font(13), text_color=COLORS["text"],
            anchor="w", justify="left",
        )
        self._label.place(x=px(11), y=px(1))

        self._chevron = ctk.CTkLabel(
            self, text="▾", width=px(16), font=font(11, "bold"),
            text_color=COLORS["text_secondary"], cursor="hand2",
        )
        self._chevron.place(relx=1.0, rely=0.5, x=-px(22), anchor="w")

        # Список открывается кликом по всему полю.
        for widget in (self, self._label, self._chevron):
            widget.bind("<Button-1>", self._open_dropdown, add="+")
        self.bind("<Configure>", self._on_resize, add="+")

        if variable is not None:
            variable.trace_add("write", lambda *_: self._sync_from_variable())
        self._sync_from_variable()

    # -- выпадающий список ------------------------------------------------- #
    def _open_dropdown(self, _event=None) -> None:
        if self._popup is not None:
            return
        try:
            self._popup = _ComboPopup(self, self._values, self._pick)
            self._popup.show()
        except Exception:
            self._popup = None
        return "break"

    def _pick(self, value: str) -> None:
        self.set(value)

    def _close_popup(self) -> None:
        self._popup = None

    # -- внешний вид -------------------------------------------------------- #
    def _on_resize(self, event) -> None:
        self._label.place(x=px(11), y=px(1))
        try:
            self._label.configure(width=max(40, event.width - px(36)), height=max(1, event.height - 2))
        except Exception:
            pass

    def _sync_from_variable(self) -> None:
        if self._variable is not None:
            self._text = str(self._variable.get())
            try:
                self._label.configure(text=self._text)
            except Exception:
                pass

    def set_error(self, state: bool) -> None:
        """Подсвечивает рамку поля красным при ошибке валидации."""
        self.configure(
            border_color=COLORS["error_border"] if state else COLORS["entry_border"]
        )

    # -- публичный API ------------------------------------------------------ #
    def set_values(self, values: Sequence[str]) -> None:
        """Обновляет список пунктов (выбранное значение не сбрасывается)."""
        self._values = [str(v) for v in values]

    def get(self) -> str:
        if self._variable is not None:
            return self._variable.get()
        return self._text

    def set(self, value: str):
        if self._variable is not None:
            self._variable.set(str(value))
        else:
            self._text = str(value)
            self._label.configure(text=self._text)

    def bind_entry(self, sequence, command, add="+"):
        return self.bind(sequence, command, add=add)


def segmented(master, values: Sequence[str], command=None, width: int = 230, height: int = 30,
              font_size: int = 13) -> ctk.CTkSegmentedButton:
    return ctk.CTkSegmentedButton(
        master,
        values=list(values),
        width=width,
        height=height,
        corner_radius=8,
        border_width=3,
        fg_color=COLORS["segment_track"],
        selected_color=COLORS["segment_selected"],
        selected_hover_color=COLORS["segment_selected"],
        unselected_color=COLORS["segment_track"],
        unselected_hover_color=COLORS["segment_hover"],
        text_color=COLORS["text"],
        font=font(font_size),
        command=command,
        dynamic_resizing=False,
    )


def primary_button(master, text: str, command, width: int = 180, height: int = 32) -> ctk.CTkButton:
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        width=width,
        height=height,
        corner_radius=8,
        border_width=0,
        fg_color=COLORS["accent"],
        hover_color=COLORS["accent_hover"],
        text_color=COLORS["on_accent"],
        font=font(13, "bold"),
    )


def secondary_button(master, text: str, command, width: int = 170, height: int = 32) -> ctk.CTkButton:
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        width=width,
        height=height,
        corner_radius=8,
        border_width=1,
        fg_color=COLORS["button_secondary"],
        hover_color=COLORS["button_secondary_hover"],
        border_color=COLORS["button_secondary_border"],
        text_color=COLORS["text"],
        font=font(13),
    )


def ghost_button(master, text: str, command, width: int = 92, height: int = 28) -> ctk.CTkButton:
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        width=width,
        height=height,
        corner_radius=7,
        border_width=0,
        fg_color="transparent",
        hover_color=COLORS["button_ghost_hover"],
        text_color=COLORS["accent"],
        font=font(12, "bold"),
    )


def hairline(master) -> tk.Frame:
    return tk.Frame(master, height=1, bd=0, highlightthickness=0, bg=resolve(COLORS["hairline"]))


class IconButton(tk.Canvas):
    """Компактная кнопка-иконка (плюс, минус, корзина и т. п.) с подсветкой."""

    def __init__(self, master, icon: str, command: Callable[[], None],
                 size: int = 26, tooltip: str = ""):
        self._icon = icon
        self._command = command
        self._hover = False
        self._size = px(size)
        self._tooltip = tooltip
        self._tip_id = None
        self._tip_window = None
        super().__init__(
            master, width=self._size, height=self._size, highlightthickness=0, bd=0,
            bg=resolve(COLORS["card"]), cursor="hand2",
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda _event: self._command())
        self._draw()
        register_theme_listener(self._draw)

    def _on_enter(self, _event) -> None:
        self._hover = True
        self._draw()
        if self._tooltip:
            try:
                self._tip_id = self.after(600, self._show_tooltip)
            except Exception:
                pass

    def _on_leave(self, _event) -> None:
        self._hover = False
        self._draw()
        try:
            if self._tip_id:
                self.after_cancel(self._tip_id)
                self._tip_id = None
            if self._tip_window is not None:
                self._tip_window.destroy()
                self._tip_window = None
        except Exception:
            pass

    def _show_tooltip(self) -> None:
        try:
            x = self.winfo_rootx() + self._size + 6
            y = self.winfo_rooty() + self._size // 2
            tip = tk.Toplevel(self)
            tip.wm_overrideredirect(True)
            tip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(
                tip, text=self._tooltip, bg=resolve(COLORS["card"]),
                fg=resolve(COLORS["text"]), font=font(11),
                padx=6, pady=2, relief="solid", borderwidth=1,
                highlightthickness=1, highlightbackground=resolve(COLORS["card_border"]),
            )
            label.pack()
            self._tip_window = tip
        except Exception:
            self._tip_window = None

    def _draw(self) -> None:
        bg = resolve(COLORS["card"])
        try:
            self.configure(bg=bg)
        except Exception:
            return
        self.delete("all")
        size = self._size
        if self._hover:
            rounded_rect(
                self, 0, 0, size - 1, size - 1, px(7),
                fill=resolve(COLORS["button_ghost_hover"]),
                outline=resolve(COLORS["button_secondary_border"]),
            )
        icon_size = px(16)
        ox, oy = (size - icon_size) / 2, (size - icon_size) / 2
        # Обычное состояние — приглушённый цвет; при наведении — акцентный.
        color = (resolve(COLORS["accent"]) if self._hover
                 else resolve(COLORS["text_secondary"]))
        ICONS[self._icon](self, ox, oy, color, icon_size, max(1.5, px(1.7)))


# --------------------------------------------------------------------------- #
#  Переключатель темы: иконки «солнце» (светлая) и «луна» (тёмная)
# --------------------------------------------------------------------------- #
class ThemeIcon(tk.Canvas):
    """Кнопка-иконка: подсветка при наведении, выделение активной темы."""

    def __init__(self, master, icon: str, command: Callable[[], None], size: int = 32):
        self._icon = icon
        self._command = command
        self._selected = False
        self._hover = False
        self._size = px(size)
        super().__init__(
            master, width=self._size, height=self._size, highlightthickness=0, bd=0,
            bg=resolve(COLORS["card"]), cursor="hand2",
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda _event: self._command())
        self._draw()
        register_theme_listener(self._draw)

    def set_selected(self, selected: bool) -> None:
        if self._selected == selected:
            return
        self._selected = selected
        self._draw()

    def _on_enter(self, _event) -> None:
        self._hover = True
        self._draw()

    def _on_leave(self, _event) -> None:
        self._hover = False
        self._draw()

    def _draw(self) -> None:
        bg = resolve(COLORS["card"])
        try:
            self.configure(bg=bg)
        except Exception:
            return
        self.delete("all")
        size = self._size
        if self._selected:
            bg = resolve(COLORS["segment_selected"])
            rounded_rect(
                self, 1, 1, size - 1, size - 1, px(8),
                fill=bg, outline=resolve(COLORS["card_border"]),
            )
        elif self._hover:
            bg = resolve(COLORS["segment_hover"])
            rounded_rect(self, 1, 1, size - 1, size - 1, px(8), fill=bg, outline="")

        color = resolve(COLORS["accent"]) if self._selected else resolve(COLORS["chrome_text"])
        icon_size = px(18)
        ox, oy = (size - icon_size) / 2, (size - icon_size) / 2
        if self._icon == "moon":
            icon_moon(self, ox, oy, color, icon_size, mask=bg)
        else:
            ICONS[self._icon](self, ox, oy, color, icon_size, max(1.3, px(1.5)))


class ThemeToggle(tk.Frame):
    """Переключатель темы: солнышко (светлая) и луна (тёмная).

    Активная тема подсвечивается, command получает строку "light" / "dark".
    """

    def __init__(self, master, command: Callable[[str], None], mode: str = "light"):
        super().__init__(master, bg=resolve(COLORS["card"]), highlightthickness=0, bd=0)
        self._command = command
        self._mode = mode
        self._buttons: dict[str, ThemeIcon] = {}
        for key, icon in (("light", "sun"), ("dark", "moon")):
            btn = ThemeIcon(self, icon, lambda m=key: self._select(m))
            btn.pack(side="left", padx=(0, px(6)))
            self._buttons[key] = btn
        self._sync()
        register_theme_listener(self._sync)

    def set(self, mode: str) -> None:
        self._mode = mode
        self._sync()

    def _select(self, mode: str) -> None:
        self._mode = mode
        self._sync()
        self._command(mode)

    def _sync(self) -> None:
        try:
            self.configure(bg=resolve(COLORS["card"]))
        except Exception:
            return
        for key, btn in self._buttons.items():
            btn.set_selected(key == self._mode)


# --------------------------------------------------------------------------- #
#  Шторка (sheet) — модальное окно в стиле macOS
# --------------------------------------------------------------------------- #
class Sheet(ctk.CTkToplevel):
    HEADER_HEIGHT = 46
    FOOTER_HEIGHT = 58

    def __init__(self, master, title: str, width: int, height: int):
        super().__init__(master)
        self.withdraw()
        self.title(title)
        self.configure(fg_color=COLORS["window"])
        self._width, self._height = width, height
        make_borderless(self, taskbar_button=False)

        shell = ctk.CTkFrame(
            self, fg_color=COLORS["window"], corner_radius=px(14), border_width=1,
            border_color=COLORS["card_border"],
        )
        shell.pack(fill="both", expand=True)
        self._shell = shell

        self._header = tk.Frame(shell, bg=resolve(COLORS["chrome"]), height=px(self.HEADER_HEIGHT))
        self._header.pack(fill="x", side="top")
        self._header.pack_propagate(False)
        self._title_label = tk.Label(
            self._header, text=title, bg=resolve(COLORS["chrome"]),
            fg=resolve(COLORS["chrome_text"]), font=font(13, "bold"), anchor="w",
        )
        self._title_label.pack(side="left", padx=px(20))

        bottom = ctk.CTkFrame(shell, fg_color="transparent")
        bottom.pack(fill="x", side="bottom")
        self._footer_line = ctk.CTkFrame(bottom, fg_color=COLORS["hairline"], corner_radius=0, height=1)
        self._footer_line.pack(fill="x")
        self._footer = ctk.CTkFrame(bottom, fg_color="transparent", height=px(self.FOOTER_HEIGHT))
        self._footer.pack(fill="x")
        self._footer.pack_propagate(False)

        self.body = ctk.CTkFrame(shell, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=px(20), pady=px(18))
        self.footer = self._footer

        self._drag = WindowDrag(self)
        self._drag.attach(self._header, self._title_label)
        self.geometry(f"{width}x{height}")
        self.bind("<Escape>", lambda _event: self._cancel())
        register_theme_listener(self._sync_theme)

    def _sync_theme(self) -> None:
        try:
            self._header.configure(bg=resolve(COLORS["chrome"]))
            self._title_label.configure(bg=resolve(COLORS["chrome"]), fg=resolve(COLORS["chrome_text"]))
        except Exception:
            pass

    def _cancel(self) -> None:
        self.destroy()

    def fit_content(self, minimum: int = 240) -> None:
        """Подгоняет высоту шторки под содержимое (как это делает AppKit)."""
        self.update_idletasks()
        height = max(px(minimum), self._shell.winfo_reqheight())
        self._height = height
        # CTkToplevel по умолчанию запирает размер в min=max, снимаем этот замок
        self.minsize(px(200), px(160))
        self.maxsize(px(2000), px(1200))
        self.geometry(f"{self._width}x{height}")

    def show(self) -> None:
        parent = self.master
        parent.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self._width) // 2)
        y = parent.winfo_rooty() + px(52)
        self.geometry(f"{self._width}x{self._height}+{x}+{y}")
        self.deiconify()
        self.lift()
        enable_rounded_corners(self)
        focus_window(self)
        try:
            self.grab_set()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
#  Алерт в стиле macOS
# --------------------------------------------------------------------------- #
class Alert(ctk.CTkToplevel):
    WIDTH = 420
    ICON_COLORS = {
        "error": ("#FF3B30", "#FF6961"),
        "success": ("#34C759", "#30D158"),
        "info": ("#007AFF", "#0A84FF"),
        "question": ("#8E8E93", "#98989D"),
    }
    GLYPHS = {"error": "!", "success": None, "info": "i", "question": "?"}

    def __init__(self, master, title: str, message: str, kind: str = "info",
                 ok_text: str = "ОК", cancel_text: Optional[str] = None):
        super().__init__(master)
        self._result = False
        self.withdraw()
        self.configure(fg_color=COLORS["window"])
        make_borderless(self, taskbar_button=False)

        shell = ctk.CTkFrame(
            self, fg_color=COLORS["window"], corner_radius=px(16), border_width=1,
            border_color=COLORS["card_border"],
        )
        shell.pack(fill="both", expand=True)

        inner = ctk.CTkFrame(shell, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=px(24), pady=(px(26), px(20)))

        self._icon = tk.Canvas(
            inner, width=px(48), height=px(48), bg=resolve(COLORS["window"]), highlightthickness=0, bd=0,
        )
        self._icon.pack(pady=(0, px(12)))
        self._kind = kind
        self._icon_color = resolve(self.ICON_COLORS.get(kind, self.ICON_COLORS["info"]))
        self._draw_icon()

        ctk.CTkLabel(
            inner, text=title, font=font(14, "bold"), text_color=COLORS["text"],
        ).pack(pady=(0, px(6)))

        ctk.CTkLabel(
            inner,
            text=message,
            font=font(12),
            text_color=COLORS["text_secondary"],
            wraplength=px(330),
            justify="center",
        ).pack(pady=(0, px(20)))

        buttons = ctk.CTkFrame(inner, fg_color="transparent")
        buttons.pack()
        if cancel_text:
            ghost_button(
                buttons, cancel_text, self._close, width=px(110), height=px(30),
            ).pack(side="left", padx=(0, px(10)))
            primary_button(buttons, ok_text, self._ok, width=px(110), height=px(30)).pack(side="left")
        else:
            primary_button(buttons, ok_text, self._close, width=px(110), height=px(30)).pack()

        self.update_idletasks()
        self._width = px(self.WIDTH)
        self._height = max(px(180), shell.winfo_reqheight())
        self.geometry(f"{self._width}x{self._height}")

        self.bind("<Escape>", lambda _event: self._close())
        self.bind("<Return>", lambda _event: self._close())
        register_theme_listener(self._sync_theme)

    def _draw_icon(self) -> None:
        self._icon.delete("all")
        self._icon.configure(bg=resolve(COLORS["window"]))
        size = px(48)
        self._icon.create_oval(2, 2, size - 2, size - 2, fill=self._icon_color, outline="")

        glyph = self.GLYPHS.get(self._kind, "i")
        if glyph is None:  # галочка для успеха
            self._icon.create_line(
                size * 0.28, size * 0.52, size * 0.44, size * 0.68, size * 0.73, size * 0.34,
                fill="#FFFFFF", width=px(3), capstyle="round", joinstyle="round",
            )
        else:
            self._icon.create_text(
                size / 2, size / 2, text=glyph, fill="#FFFFFF",
                font=(FAMILY_TEXT, -px(24), "bold"),
            )

    def _sync_theme(self) -> None:
        try:
            self._draw_icon()
        except Exception:
            pass

    def _ok(self) -> None:
        self._result = True
        self._close()

    def _close(self) -> None:
        try:
            self.grab_release()
        except Exception:
            pass
        self.destroy()

    def show(self, parent: tk.Misc) -> None:
        parent.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self._width) // 2)
        y = parent.winfo_rooty() + max(px(60), (parent.winfo_height() - self._height) // 3)
        self.geometry(f"{self._width}x{self._height}+{x}+{y}")
        self.deiconify()
        self.lift()
        enable_rounded_corners(self)
        focus_window(self)
        try:
            self.grab_set()
        except Exception:
            pass
        self.wait_window()


def alert(parent: tk.Misc, title: str, message: str, kind: str = "info", ok_text: str = "ОК") -> None:
    """Показывает модальный алерт в стиле macOS и ждёт закрытия."""
    if parent is None:
        return
    Alert(parent, title, message, kind, ok_text).show(parent)


def confirm(parent: tk.Misc, title: str, message: str,
            ok_text: str = "ОК", cancel_text: str = "Отмена") -> bool:
    """Модальный вопрос с двумя кнопками; True — если нажата главная."""
    if parent is None:
        return False
    dialog = Alert(parent, title, message, "question", ok_text, cancel_text)
    dialog.show(parent)
    return bool(dialog._result)
