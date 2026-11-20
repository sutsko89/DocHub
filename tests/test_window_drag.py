"""WindowDrag — ручное перетаскивание окна за шапку (2.19.3).

Нативный модальный цикл Windows (2.17.0 — синхронный SendMessageW из
обработчика нажатия, 2.19.1 — отложенный via after_idle) ронял процесс
на ряде машин с Windows 11 без traceback и с пустым crash.log: вложенный
message loop перекачивает сообщения окна через Tk-обработчики, Tcl не
рассчитан на такую ре-entrant обработку — повреждение памяти завершается
fast-fail, который не ловит даже faulthandler. В 2.19.3 нативный путь
удалён полностью: перенос — только через geometry() (работал без вылетов
во всех версиях до 2.17.0), позиции зажимаются в виртуальный экран
(объединение мониторов — перенос на второй монитор сохранён, Snap — нет).

Тесты headless: duck-typed окно вместо Tk. Проверяются контракты
WindowDrag: нажатие вооружает offset, движение двигает окно в границах
виртуального экрана, привязки добавляются к существующим (add="+").
"""
from __future__ import annotations

from types import SimpleNamespace

import ui_macos as ui


class FakeDragWindow:
    """Duck-typed аналог Tk-окна: запоминает вызовы, двигает позицию."""

    def __init__(self, width=900, height=600):
        self.calls: list[tuple] = []
        self._x, self._y = 100, 80
        self._width, self._height = width, height

    def update_idletasks(self):
        self.calls.append(("update_idletasks",))

    def winfo_x(self):
        return self._x

    def winfo_y(self):
        return self._y

    def winfo_width(self):
        return self._width

    def winfo_exists(self):
        return True

    def geometry(self, spec=""):
        if spec:
            self.calls.append(("geometry", spec))
            pos = spec.lstrip("+")
            x, _, y = pos.partition("+")
            self._x, self._y = int(x), int(y)

    def after_idle(self, fn):
        self.calls.append(("after_idle", fn))


def event(x, y):
    return SimpleNamespace(x_root=x, y_root=y)


# --------------------------------------------------------------------------- #
#  Привязки
# --------------------------------------------------------------------------- #
def test_attach_binds_press_motion_release():
    bound: dict[str, list] = {}
    widget = SimpleNamespace(
        bind=lambda seq, fn, add="": bound.setdefault(seq, []).append((fn, add)),
    )
    drag = ui.WindowDrag(SimpleNamespace())
    drag.attach(widget)
    assert set(bound) == {"<ButtonPress-1>", "<B1-Motion>", "<ButtonRelease-1>"}
    assert all(add == "+" for pairs in bound.values() for _, add in pairs), \
        "привязки добавляются к существующим (add='+')"


# --------------------------------------------------------------------------- #
#  Ручной перенос: нажатие → движение → отпускание
# --------------------------------------------------------------------------- #
def test_press_arms_offset_without_blocking_calls():
    """Нажатие считает offset и не делает ничего блокирующего."""
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    drag._press(event(150, 120))
    # offset = (x_root - winfo_x, y_root - winfo_y) = (50, 40)
    assert drag._offset == (50, 40)
    assert any(call[0] == "update_idletasks" for call in window.calls), \
        "геометрия пересчитывается на нажатии (on_start мог её сменить)"
    assert not any(call[0] == "geometry" for call in window.calls), \
        "в колбэке нажатия нет переноса окна"


def test_manual_drag_moves_window():
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    drag._press(event(150, 120))
    drag._move(event(160, 140))
    assert ("geometry", "+110+100") in window.calls
    drag._release(event(160, 140))


def test_move_is_noop_before_press_and_after_release():
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    drag._move(event(160, 140))
    assert not any(kind == "geometry" for kind, _ in window.calls)
    drag._press(event(150, 120))
    drag._release(event(160, 140))
    drag._move(event(200, 200))
    assert not any(
        call[1] == "+150+160" for call in window.calls if call[0] == "geometry"
    ), "после отпускания кнопки окно не двигается"


def test_on_start_receives_press_event():
    window = FakeDragWindow()
    started = []
    drag = ui.WindowDrag(window, on_start=lambda ev: started.append(ev))
    drag._press(event(150, 120))
    assert len(started) == 1 and started[0].x_root == 150


# --------------------------------------------------------------------------- #
#  Зажим в виртуальный экран (2.19.3)
# --------------------------------------------------------------------------- #
def test_move_clamps_to_virtual_screen(monkeypatch):
    """Окно не теряется за краями: остаётся полоса DRAG_EDGE_MARGIN."""
    window = FakeDragWindow(width=900)
    drag = ui.WindowDrag(window)
    # Виртуальный экран: два монитора 1920×1080, второй слева.
    monkeypatch.setattr(
        ui, "virtual_screen_bounds", lambda: (-1920, 0, 3840, 1080))
    drag._press(event(150, 120))  # offset (50, 40)
    drag._move(event(0, 5))       # x = -50, y = -35
    # x: зажим в [-1920 - 900 + 48, -1920 + 3840 - 48] → -50 внутри
    # y: зажим в [0, 1080 - 48] → 0
    assert ("geometry", "+-50+0") in window.calls


def test_move_keeps_window_grabbable_at_left_edge(monkeypatch):
    """Далеко за левым краем окно не улетает — margin остаётся видимым."""
    window = FakeDragWindow(width=900)
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(
        ui, "virtual_screen_bounds", lambda: (-1920, 0, 3840, 1080))
    drag._press(event(150, 120))          # offset (50, 40)
    drag._move(event(-3300, 500))         # x = -3350, y = 460
    # нижняя граница x: bounds_x - win_w + margin = -1920 - 900 + 48 = -2772
    assert ("geometry", "+-2772+460") in window.calls


def test_move_keeps_titlebar_below_top_of_virtual_screen(monkeypatch):
    """Шапка не прячется выше верхнего монитора: y ≥ bounds_y."""
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(
        ui, "virtual_screen_bounds", lambda: (0, -200, 1920, 1280))
    drag._press(event(150, 120))  # offset (50, 40)
    drag._move(event(160, -500))  # y = -540
    assert ("geometry", "+110+-200") in window.calls


def test_move_survives_bounds_failure(monkeypatch):
    """Сбой расчёта границ не прерывает перенос (позиция без зажима)."""
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)

    def boom():
        raise OSError("win32 сбой")

    monkeypatch.setattr(ui, "virtual_screen_bounds", boom)
    drag._press(event(150, 120))
    drag._move(event(160, 140))
    assert ("geometry", "+110+100") in window.calls, \
        "перенос работает и без зажима"
