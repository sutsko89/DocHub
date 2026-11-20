"""WindowDrag — диспетчеризация перетаскивания окна (фикс вылета 2.19.1).

Причина падения процесса на Windows (релиз 2.19.0): обработчик нажатия
шапки запускал SendMessageW(WM_NCLBUTTONDOWN, HTCAPTION) синхронно —
вызов блокирует до отпускания кнопки и раскручивает вложенный message
loop поверх незавершённого Tk-колбэка; события Tcl обрабатываются
ре-entrantly и процесс умирает без traceback. Фикс: отложенный старт
(begin_native_move_deferred → after_idle), ручной путь остаётся запасным.

Тесты headless: duck-typed окно вместо Tk. Проверяются контракты
WindowDrag: нажатие не выполняет блокирующей работы в колбэке, ручной
fallback считает геометрию и зажимает позиции, привязки добавляются
к существующим (add="+").
"""
from __future__ import annotations

from types import SimpleNamespace

import ui_macos as ui


class FakeDragWindow:
    """Duck-typed аналог Tk-окна: запоминает вызовы, двигает позицию."""

    def __init__(self):
        self.calls: list[tuple] = []
        self._x, self._y = 100, 80

    def update_idletasks(self):
        self.calls.append(("update_idletasks",))

    def winfo_x(self):
        return self._x

    def winfo_y(self):
        return self._y

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
#  Windows-путь: блокирующий вызов не живёт в колбэке нажатия
# --------------------------------------------------------------------------- #
def test_press_with_native_path_defers_and_returns_fast(monkeypatch):
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(ui, "begin_native_move_deferred", lambda w: True)
    drag._press(event(150, 120))
    # Нативный путь: offset сброшен, никакой ручной геометрии в колбэке.
    assert drag._offset is None
    assert not any(kind == "geometry" for kind, _ in window.calls)
    assert not any(kind == "update_idletasks" for kind, _ in window.calls)


def test_press_falls_back_manually_when_native_unavailable(monkeypatch):
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(ui, "begin_native_move_deferred", lambda w: False)
    drag._press(event(150, 120))
    # offset = (x_root - winfo_x, y_root - winfo_y) = (50, 40)
    assert drag._offset == (50, 40)


def test_press_survives_native_dispatch_failure(monkeypatch):
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)

    def boom(window_arg):
        raise RuntimeError("сбой планирования")

    monkeypatch.setattr(ui, "begin_native_move_deferred", boom)
    drag._press(event(150, 120))
    assert drag._offset == (50, 40), "сбой нативной диспетчеризации → ручной путь"


# --------------------------------------------------------------------------- #
#  Ручной fallback: перенос окна по событиям
# --------------------------------------------------------------------------- #
def test_manual_drag_moves_window(monkeypatch):
    # Ручной путь закрепляем явно: на Windows настоящий
    # begin_native_move_deferred уходит в after_idle (нативный цикл) —
    # тест же проверяет именно fallback.
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(ui, "begin_native_move_deferred", lambda w: False)
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


def test_manual_drag_clamps_positions(monkeypatch):
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(ui, "begin_native_move_deferred", lambda w: False)
    drag._press(event(150, 120))  # offset (50, 40)
    drag._move(event(20, 5))      # x = -30, y = -35 → зажим (-30, 0)
    assert ("geometry", "+-30+0") in window.calls


def test_on_start_receives_press_event():
    window = FakeDragWindow()
    started = []
    drag = ui.WindowDrag(window, on_start=lambda ev: started.append(ev))
    drag._press(event(150, 120))
    assert len(started) == 1 and started[0].x_root == 150
