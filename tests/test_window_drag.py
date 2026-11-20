"""WindowDrag — ручное перетаскивание окна за шапку + прилипание (2.20.0).

Нативный модальный цикл Windows (2.17.0 — синхронный SendMessageW из
обработчика нажатия, 2.19.1 — отложенный via after_idle) ронял процесс
на ряде машин с Windows 11 без traceback и с пустым crash.log: вложенный
message loop перекачивает сообщения окна через Tk-обработчики, Tcl не
рассчитан на такую ре-entrant обработку — повреждение памяти завершается
fast-fail, который не ловит даже faulthandler. В 2.19.3 нативный путь
удалён полностью: перенос — только через geometry() (работал без вылетов
во всех версиях до 2.17.0), позиции зажимаются в виртуальный экран
(объединение мониторов — перенос на второй монитор сохранён).

2.20.0: перенос и растягивание получают постоянный темп через пулер
курсора (PointerPoller, ~60 Гц вместо рваных <B1-Motion>), а Aero Snap
эмулируется вручную: snap_zone/snap_rect считают зону у краёв рабочей
области монитора под курсором, SnapPreview показывает предпросмотр,
snap_apply применяет прилипание при отпускании кнопки.

Тесты headless: duck-typed окно вместо Tk. Проверяются контракты
WindowDrag: нажатие вооружает offset, движение двигает окно в границах
виртуального экрана, привязки добавляются к существующим (add="+"),
зоны прилипания активируют предпросмотр и применяются на отпускании.
"""
from __future__ import annotations

from types import SimpleNamespace

import ui_macos as ui


class FakeDragWindow:
    """Duck-typed аналог Tk-окна: запоминает вызовы, двигает позицию."""

    def __init__(self, width=900, height=600, pointer=(0, 0)):
        self.calls: list[tuple] = []
        self._x, self._y = 100, 80
        self._width, self._height = width, height
        self.pointer = pointer
        self.timers: list[tuple[int, object]] = []
        self.cancelled: list[int] = []

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

    def winfo_pointerx(self):
        return self.pointer[0]

    def winfo_pointery(self):
        return self.pointer[1]

    def after(self, ms, fn):
        self.timers.append((ms, fn))
        return len(self.timers)

    def after_cancel(self, timer_id):
        self.cancelled.append(timer_id)

    def geometry(self, spec=""):
        if spec:
            self.calls.append(("geometry", spec))
            size, _, pos = spec.partition("+")
            if size:
                w, _, h = size.partition("x")
                self._width, self._height = int(w), int(h)
            if pos:
                x, _, y = pos.partition("+")
                self._x, self._y = int(x), int(y)
        return f"{self._width}x{self._height}+{self._x}+{self._y}"

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


# --------------------------------------------------------------------------- #
#  Прилипание: чистые функции (2.20.0)
# --------------------------------------------------------------------------- #
def test_snap_zone_top_edge_maximizes():
    """Верхняя кромка окна у верха рабочей области — разворачивание."""
    area = (0, 0, 1920, 1080)
    assert ui.snap_zone(510, 10, 900, area, edge=16) == "max"
    # Полоса симметрична: окно чуть выше области тоже считается у края.
    assert ui.snap_zone(510, -10, 900, area, edge=16) == "max"


def test_snap_zone_sides_and_center():
    area = (0, 0, 1920, 1080)
    assert ui.snap_zone(0, 300, 900, area, edge=16) == "left"
    assert ui.snap_zone(10, 300, 900, area, edge=16) == "left"
    assert ui.snap_zone(1036, 300, 900, area, edge=16) == "right"
    assert ui.snap_zone(600, 300, 900, area, edge=16) is None


def test_snap_zone_corner_is_half_not_quarter():
    """Верхний угол — половина по горизонтали (стиль Windows 10)."""
    area = (0, 0, 1920, 1080)
    assert ui.snap_zone(0, 0, 900, area, edge=16) == "left"
    assert ui.snap_zone(1036, 0, 900, area, edge=16) == "right"


def test_snap_zone_pushed_through_edge_cancels():
    """Окно вдвинуто дальше полосы за край — прилипание отменяется
    (пользователь перетаскивает окно на соседний монитор)."""
    area = (0, 0, 1920, 1080)
    assert ui.snap_zone(-40, 300, 900, area, edge=16) is None
    assert ui.snap_zone(1060, 300, 900, area, edge=16) is None
    assert ui.snap_zone(510, -40, 900, area, edge=16) is None


def test_snap_zone_degenerate_input():
    assert ui.snap_zone(0, 0, 900, (0, 0, 0, 0), edge=16) is None
    assert ui.snap_zone(0, 0, 900, (0, 0, 1920, 1080), edge=0) is None
    assert ui.snap_zone(0, 0, 0, (0, 0, 1920, 1080), edge=16) is None


def test_snap_rect_halves_divide_workarea():
    """Половины делят рабочую область без щелей и перекрытий."""
    area = (0, 0, 1921, 900)
    assert ui.snap_rect("max", area) == (0, 0, 1921, 900)
    left = ui.snap_rect("left", area)
    right = ui.snap_rect("right", area)
    assert left == (0, 0, 960, 900)
    assert right == (960, 0, 961, 900)
    assert left[0] + left[2] == right[0], "правая половина начинается вплотную"
    assert right[0] + right[2] == area[0] + area[2], "правая половина до края области"
    assert ui.snap_rect("bogus", area) is None


def test_snap_rect_keeps_monitor_offset():
    """Прилипание работает и на втором мониторе (сдвиг области учитывается)."""
    area = (-1920, 0, 1920, 1080)  # монитор слева от основного
    assert ui.snap_rect("right", area) == (-960, 0, 960, 1080)
    assert ui.snap_rect("left", area) == (-1920, 0, 960, 1080)


# --------------------------------------------------------------------------- #
#  Прилипание: контракт WindowDrag (2.20.0)
# --------------------------------------------------------------------------- #
class FakePreview:
    def __init__(self):
        self.shown: list[tuple] = []
        self.hidden = 0

    def show_rect(self, rect):
        self.shown.append(rect)

    def hide(self):
        self.hidden += 1


def make_drag(window, resolver, applied, preview_box):
    return ui.WindowDrag(
        window,
        snap_resolver=resolver,
        snap_apply=lambda zone, rect, start: applied.append((zone, rect, start)),
        preview_factory=lambda: preview_box.setdefault("preview", FakePreview()),
    )


def test_drag_left_edge_shows_preview_and_applies_on_release(monkeypatch):
    window = FakeDragWindow()
    applied: list[tuple] = []
    preview_box: dict = {}
    resolver_calls: list[tuple] = []

    def resolver(cursor_x, cursor_y, win_x, win_y, win_w):
        resolver_calls.append((cursor_x, cursor_y, win_x, win_y, win_w))
        return "left", (0, 0, 960, 1080)

    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag = make_drag(window, resolver, applied, preview_box)
    drag._press(event(150, 120))  # offset (50, 40)
    drag._move(event(60, 500))    # x = 10 — левая кромка у края области
    assert resolver_calls[-1] == (60, 500, 10, 460, 900), \
        "резольвер получает курсор и целевую позицию окна"
    preview = preview_box["preview"]
    assert preview.shown == [(0, 0, 960, 1080)], "предпросмотр показывает половину"
    drag._release(event(60, 500))
    assert applied == [("left", (0, 0, 960, 1080), "900x600+100+80")], \
        "на отпускании прилипание получает зону, прямоугольник и прежнюю геометрию"
    assert preview.hidden >= 1, "предпросмотр скрыт после отпускания"


def test_drag_top_edge_resolves_maximize(monkeypatch):
    window = FakeDragWindow()
    applied: list[tuple] = []
    preview_box: dict = {}
    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag = make_drag(window, lambda *args: ("max", (0, 0, 1920, 1080)), applied, preview_box)
    drag._press(event(150, 120))
    drag._move(event(600, 40))    # y = 0 — верхняя кромка у верха области
    assert preview_box["preview"].shown == [(0, 0, 1920, 1080)]
    drag._release(event(600, 40))
    assert applied == [("max", (0, 0, 1920, 1080), "900x600+100+80")]


def test_drag_center_does_not_snap(monkeypatch):
    window = FakeDragWindow()
    applied: list[tuple] = []
    preview_box: dict = {}
    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag = make_drag(window, lambda *args: None, applied, preview_box)
    drag._press(event(150, 120))
    drag._move(event(600, 500))
    drag._release(event(600, 500))
    assert applied == [], "в центре экрана прилипания нет"
    assert "preview" not in preview_box, "подложка не создаётся без зоны"


def test_drag_zone_leaving_hides_preview(monkeypatch):
    """Выход из зоны (уход к соседнему монитору) прячет предпросмотр."""
    window = FakeDragWindow()
    applied: list[tuple] = []
    preview_box: dict = {}
    zones = iter([("left", (0, 0, 960, 1080)), None])
    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag = make_drag(window, lambda *args: next(zones, None), applied, preview_box)
    drag._press(event(150, 120))
    drag._move(event(60, 500))
    preview = preview_box["preview"]
    assert preview.shown and preview.hidden == 0
    drag._move(event(700, 500))   # резольвер вернул None — окно ушло из зоны
    assert preview.hidden >= 1, "предпросмотр скрыт при выходе из зоны"
    drag._release(event(700, 500))
    assert applied == [], "прилипание не применяется вне зоны"


def test_drag_without_resolver_keeps_plain_move(monkeypatch):
    """Без snap-колбэков перенос работает как в 2.19.3 (обратная совместимость)."""
    window = FakeDragWindow()
    drag = ui.WindowDrag(window)
    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag._press(event(150, 120))
    drag._move(event(160, 140))
    drag._release(event(160, 140))
    assert ("geometry", "+110+100") in window.calls


def test_drag_resolver_failure_does_not_break_move(monkeypatch):
    """Сбой резольвера не прерывает перенос: окно двигается, прилипания нет."""
    window = FakeDragWindow()
    applied: list[tuple] = []

    def boom(*args):
        raise OSError("win32 сбой")

    monkeypatch.setattr(ui, "virtual_screen_bounds", lambda: (0, 0, 1920, 1080))
    drag = ui.WindowDrag(
        window, snap_resolver=boom,
        snap_apply=lambda zone, rect, start: applied.append((zone, rect, start)),
    )
    drag._press(event(150, 120))
    drag._move(event(160, 140))
    drag._release(event(160, 140))
    assert ("geometry", "+110+100") in window.calls and applied == []


# --------------------------------------------------------------------------- #
#  PointerPoller — постоянный темп переноса (2.20.0)
# --------------------------------------------------------------------------- #
def test_press_starts_poller_and_tick_moves_window():
    """Пулер читает курсор мимо очереди событий и двигает окно."""
    window = FakeDragWindow(pointer=(160, 140))
    drag = ui.WindowDrag(window)
    drag._press(event(150, 120))
    assert window.timers and window.timers[0][0] == 16, \
        "после нажатия заведён опрос курсора (~60 Гц)"
    tick = window.timers[0][1]
    tick()
    assert ("geometry", "+110+100") in window.calls
    assert len(window.timers) == 2, "тик перевешивается — опрос продолжается"
    drag._release(event(160, 140))
    tick()  # тик после отпускания не двигает окно
    assert sum(1 for call in window.calls
               if call[0] == "geometry" and call[1] == "+110+100") == 1


def test_poller_repeated_same_position_skips_geometry():
    """Одинаковая позиция не порождает лишних geometry() (без дёрганья)."""
    window = FakeDragWindow(pointer=(160, 140))
    drag = ui.WindowDrag(window)
    drag._press(event(150, 120))
    tick = window.timers[0][1]
    tick()
    tick()
    assert sum(1 for call in window.calls if call[0] == "geometry") == 1
    drag._release(event(160, 140))


def test_poller_stop_cancels_timer():
    win = SimpleNamespace(
        after=lambda ms, fn: 1, after_cancel=lambda timer_id: None,
        winfo_pointerx=lambda: 0, winfo_pointery=lambda: 0,
    )
    cancelled: list[int] = []
    win.after_cancel = lambda timer_id: cancelled.append(timer_id)
    poller = ui.PointerPoller(win, lambda x, y: None)
    poller.start()
    poller.stop()
    assert cancelled == [1], "таймер опроса отменён"
    assert poller._active is False


def test_poller_tick_without_pointer_api_degrades():
    """Окно без winfo_pointerx (duck-typed тесты): тик гасит пулер без падения."""
    win = SimpleNamespace(after=lambda ms, fn: 1, after_cancel=lambda timer_id: None)
    steps: list[tuple] = []
    poller = ui.PointerPoller(win, lambda x, y: steps.append((x, y)))
    poller.start()
    poller._tick()
    assert steps == [] and poller._active is False


def test_poller_start_without_after_is_safe():
    """Duck-typed окно без after (unit-тесты WindowDrag): пулер не заводится."""
    poller = ui.PointerPoller(SimpleNamespace(), lambda x, y: None)
    poller.start()
    assert poller._active is False


def test_poller_callback_failure_stops_loop():
    """Сбой callback гасит пулер: окно не остаётся с бесконечным таймером."""
    def boom(x, y):
        raise RuntimeError("тестовый сбой")

    win = SimpleNamespace(
        after=lambda ms, fn: 1, after_cancel=lambda timer_id: None,
        winfo_pointerx=lambda: 10, winfo_pointery=lambda: 20,
    )
    poller = ui.PointerPoller(win, boom)
    poller.start()
    poller._tick()
    assert poller._active is False
