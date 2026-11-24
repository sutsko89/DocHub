"""Темы Qt-каркаса: светлая/тёмная через QPalette + QSS (этап M2 миграции).

Ключ и значения настроек — прежние, профиль совместим с Tk-сборкой:
``settings.json`` → ``"theme": "light" | "dark"`` (Tk пишет то же самое,
main.py::_persist_settings). Значение «system» в Qt-сборке M2 не
поддерживается: нормализация приводит его к ``light`` — как Tk-сборка
интерпретирует неизвестные значения (customtkinter сам ставит дефолт).

Пары цветов скопированы из ``ui_macos.COLORS`` (светлая/тёмная пара на
ключ) — визуальный язык каркаса тот же. Копия сознательна: импортировать
ui_macos из Qt-каркаса нельзя (там tkinter/customtkinter). При изменении
палитры в ui_macos.py синхронизировать эту копию (до M8).

Ничего из запретов раздела З плана не нарушается: рамка окна остаётся
стандартной системной, рисуются только внутренние поверхности.
"""

from __future__ import annotations

from PySide6.QtGui import QColor, QPalette

DEFAULT_THEME = "light"
VALID_MODES = ("light", "dark")

# Пары цветов зеркалят ui_macos.COLORS: (светлая, тёмная).
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
    "hairline": ("#E4E4EA", "#343438"),
    "accent": ("#007AFF", "#0A84FF"),
    "accent_hover": ("#2E93FF", "#3B9BFF"),
    "on_accent": ("#FFFFFF", "#FFFFFF"),
}

_current_mode: str = DEFAULT_THEME


def normalize_mode(value) -> str:
    """Приводит сохранённое значение темы к "light"/"dark" (неизвестное — light)."""
    text = str(value or "").strip().lower()
    return text if text in VALID_MODES else DEFAULT_THEME


def current_mode() -> str:
    """Текущий применённый режим (после apply_theme; до него — дефолт)."""
    return _current_mode


def color(mode: str, key: str) -> str:
    """Hex-цвет пары ``key`` для режима ``mode`` (неизвестный ключ — красный маркер)."""
    pair = COLORS.get(key)
    if pair is None:
        return "#FF00FF"  # заметный маркер забытого ключа, не тихий чёрный
    index = VALID_MODES.index(normalize_mode(mode))
    return pair[index]


def apply_theme(app, mode) -> str:
    """Применяет тему к QApplication: QPalette + глобальный QSS.

    Возвращает нормализованный режим. Вызов идемпотентен: QPalette и QSS
    устанавливаются целиком, накапливания стилей нет.
    """
    normalized = normalize_mode(mode)
    global _current_mode
    _current_mode = normalized
    app.setPalette(palette_for(normalized))
    app.setStyleSheet(qss_for(normalized))
    return normalized


def palette_for(mode: str) -> QPalette:
    """QPalette режима: базовые роли Qt для виджетов без QSS."""
    mode = normalize_mode(mode)
    palette = QPalette()
    set_palette_color(palette, QPalette.ColorRole.Window, mode, "window")
    set_palette_color(palette, QPalette.ColorRole.WindowText, mode, "text")
    set_palette_color(palette, QPalette.ColorRole.Base, mode, "entry_bg")
    set_palette_color(palette, QPalette.ColorRole.AlternateBase, mode, "content")
    set_palette_color(palette, QPalette.ColorRole.Text, mode, "text")
    set_palette_color(palette, QPalette.ColorRole.Button, mode, "card")
    set_palette_color(palette, QPalette.ColorRole.ButtonText, mode, "text")
    set_palette_color(palette, QPalette.ColorRole.ToolTipBase, mode, "card")
    set_palette_color(palette, QPalette.ColorRole.ToolTipText, mode, "text")
    set_palette_color(palette, QPalette.ColorRole.Highlight, mode, "accent")
    set_palette_color(palette, QPalette.ColorRole.HighlightedText, mode, "on_accent")
    set_palette_color(palette, QPalette.ColorRole.PlaceholderText, mode, "text_tertiary")
    # Отключённые виджеты — третичный текст, а не унаследованный тёмный.
    disabled_text = QColor(color(mode, "text_tertiary"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, disabled_text)
    palette.setColor(
        QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, disabled_text,
    )
    palette.setColor(
        QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, disabled_text,
    )
    return palette


def set_palette_color(palette: QPalette, role, mode: str, key: str) -> None:
    palette.setColor(role, QColor(color(mode, key)))


def qss_for(mode: str) -> str:
    """Глобальный QSS каркаса: сайдбар, футер, кнопки, карточки, тост.

    Держится компактным: детальная стилизация страниц появится вместе с
    самими страницами (M3+), здесь — только каркасные поверхности.
    """
    c = {key: color(mode, key) for key in COLORS}
    return f"""
QMainWindow, QDialog {{
    background: {c["window"]};
}}
QWidget#centralBody {{
    background: {c["content"]};
}}
QWidget {{
    font-family: "Segoe UI";
}}
QLabel {{
    color: {c["text"]};
    background: transparent;
}}
QLabel#statusLabel {{
    color: {c["text_secondary"]};
    font-size: 12px;
}}
QLabel#pagePlaceholder {{
    color: {c["text_secondary"]};
    font-size: 15px;
}}
QLabel#pagePlaceholderHint {{
    color: {c["text_tertiary"]};
    font-size: 12px;
}}
QLabel#pageTitle {{
    color: {c["text"]};
    font-size: 20px;
    font-weight: bold;
}}
QLabel#pageSubtitle {{
    color: {c["text_secondary"]};
    font-size: 12px;
}}
QLabel#currentPartyLabel {{
    color: {c["text"]};
    font-weight: bold;
    font-size: 12px;
}}
QLabel#cardCaption {{
    color: {c["text_secondary"]};
    font-size: 12px;
    font-weight: bold;
    letter-spacing: 0.3px;
}}
QLabel#fieldHint {{
    color: {c["text_tertiary"]};
    font-size: 11px;
}}
QLabel#fieldError {{
    color: transparent;
    font-size: 11px;
}}
QLabel#fieldErrorActive {{
    color: #D03050;
    font-size: 11px;
}}
QLabel#specTotals {{
    color: {c["text_secondary"]};
    font-size: 12px;
}}
QLabel#dadataStatus {{
    color: {c["accent"]};
    font-size: 12px;
}}
QLabel#previewTitle {{
    color: {c["text_secondary"]};
    font-size: 12px;
    font-weight: bold;
}}
QFrame#card {{
    background: {c["card"]};
    border: 1px solid {c["card_border"]};
    border-radius: 10px;
}}
QFrame#partyBar {{
    background: {c["sidebar_hover"]};
    border-radius: 8px;
}}
QLineEdit, QComboBox {{
    background: {c["entry_bg"]};
    color: {c["text"]};
    border: 1px solid {c["entry_border"]};
    border-radius: 6px;
    padding: 4px 8px;
}}
QLineEdit:focus, QComboBox:focus {{
    border: 1px solid {c["accent"]};
}}
QLineEdit[errorState="true"], QComboBox[errorState="true"] {{
    border: 1px solid #D03050;
}}
QPlainTextEdit#previewText {{
    background: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["card_border"]};
    border-radius: 8px;
}}
QTableWidget {{
    background: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["card_border"]};
    gridline-color: {c["hairline"]};
    selection-background-color: {c["sidebar_active"]};
    selection-color: {c["text"]};
}}
QHeaderView::section {{
    background: {c["chrome"]};
    color: {c["text_secondary"]};
    border: none;
    border-bottom: 1px solid {c["chrome_line"]};
    padding: 6px;
    font-weight: bold;
}}
QFrame#footer {{
    background: {c["window"]};
    border-top: 1px solid {c["hairline"]};
}}
QListWidget#navPanel {{
    background: {c["sidebar"]};
    border: none;
    outline: none;
    color: {c["sidebar_text"]};
    font-size: 13px;
}}
QListWidget#navPanel::item {{
    padding: 9px 14px;
    border-radius: 8px;
    margin: 2px 8px;
}}
QListWidget#navPanel::item:hover {{
    background: {c["sidebar_hover"]};
    color: {c["chrome_text"]};
}}
QListWidget#navPanel::item:selected {{
    background: {c["sidebar_active"]};
    color: {c["sidebar_text_active"]};
}}
QPushButton {{
    background: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["entry_border"]};
    border-radius: 6px;
    padding: 6px 14px;
}}
QPushButton:hover {{
    background: {c["sidebar_hover"]};
}}
QPushButton[variant="primary"] {{
    background: {c["accent"]};
    color: {c["on_accent"]};
    border: none;
}}
QPushButton[variant="primary"]:hover {{
    background: {c["accent_hover"]};
}}
QPushButton[variant="segment"] {{
    padding: 5px 18px;
    border-top-right-radius: 0;
    border-bottom-right-radius: 0;
    margin-right: -1px;
}}
QPushButton[variant="segment"]:checked {{
    background: {c["sidebar_active"]};
    color: {c["sidebar_text_active"]};
    font-weight: bold;
}}
QPushButton:disabled {{
    color: {c["text_tertiary"]};
    border-color: {c["hairline"]};
}}
QFrame#toast {{
    background: {c["card"]};
    border: 1px solid {c["card_border"]};
    border-radius: 10px;
}}
QToolTip {{
    background: {c["card"]};
    color: {c["text"]};
    border: 1px solid {c["card_border"]};
    padding: 4px;
}}
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background: {c["chrome_line"]};
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {c["text_tertiary"]};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}
"""
