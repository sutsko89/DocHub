"""win_shell — Win32-обвязка окна: DPI (п. 6.3), папка «Документы», хуки.

На не-Windows функции обязаны деградировать в безопасные аналоги Tk или
no-op: здесь это проверяется на фейковых окнах без настоящего Tk. Чистые
хелперы WM_DPICHANGED (масштаб из wparam, прямоугольник-подсказка из
lparam) тестируются напрямую — они не обращаются к Win32.

Windows-раннер CI (2.15.0): тесты секции «не-Windows» помечены
``skipif(win32)`` — на Windows эти функции идут по ветке реального Win32
(SHGetKnownFolderPath, GetMonitorInfo, ShowWindow), и подмена фейковым
окном на них не действует: сами функции при этом покрыты smoke-тестом на
реальном приложении.
"""
from __future__ import annotations

import ctypes
import os
import sys

import pytest

import win_shell

not_win32 = pytest.mark.skipif(
    sys.platform == "win32",
    reason="тест закрепляет не-Windows ветку (фейк не подменяет Win32-API)",
)


class FakeWindow:
    """Duck-typed окно Tk: записывает вызовы, не создаёт виджетов."""

    def __init__(self, screen=(1366, 768)):
        self.screen = screen
        self.calls: list[tuple] = []
        self._winfo_id = 12345

    def winfo_id(self):
        return self._winfo_id

    def winfo_screenwidth(self):
        return self.screen[0]

    def winfo_screenheight(self):
        return self.screen[1]

    def update_idletasks(self):
        pass

    def overrideredirect(self, value):
        self.calls.append(("overrideredirect", value))

    def iconify(self):
        self.calls.append(("iconify",))

    def state(self, value):
        self.calls.append(("state", value))

    def geometry(self, spec=""):
        if spec:
            self.calls.append(("geometry", spec))

    def focus_force(self):
        self.calls.append(("focus_force",))

    def after(self, _ms, fn):
        self.calls.append(("after", fn))

    def after_idle(self, fn):
        self.calls.append(("after_idle", fn))

    def winfo_exists(self):
        return True


# --------------------------------------------------------------------------- #
#  DPI (п. 6.3)
# --------------------------------------------------------------------------- #
def test_enable_dpi_awareness_is_safe_noop_on_linux():
    win_shell.enable_dpi_awareness()  # не бросает исключений вне Windows


def test_dpi_scale_from_wparam():
    assert win_shell.dpi_scale_from_wparam((96 << 16) | 96) == 1.0
    assert win_shell.dpi_scale_from_wparam((120 << 16) | 120) == 1.25
    assert win_shell.dpi_scale_from_wparam((144 << 16) | 144) == 1.5
    assert win_shell.dpi_scale_from_wparam((192 << 16) | 192) == 2.0


def test_dpi_scale_from_wparam_zero_falls_back_to_100():
    assert win_shell.dpi_scale_from_wparam(0) == 1.0


def test_suggested_rect_from_lparam():
    rc = win_shell._RECT(100, 50, 1700, 950)
    rect = win_shell.suggested_rect_from_lparam(ctypes.addressof(rc))
    assert rect == (100, 50, 1600, 900)


def test_message_constants():
    assert win_shell.WM_DROPFILES == 0x0233
    assert win_shell.WM_DPICHANGED == 0x02E0
    assert win_shell.WM_NCLBUTTONDOWN == 0x00A1
    assert win_shell.HTCAPTION == 0x2


# --------------------------------------------------------------------------- #
#  Нативное перетаскивание/растягивание (2.17.0)
# --------------------------------------------------------------------------- #
def test_ht_resize_edges_cover_all_eight_sides():
    """Все 8 кромок имеют HT-коды, значения — системные константы Win32."""
    assert set(win_shell.HT_RESIZE_EDGES) == {"n", "s", "w", "e", "nw", "ne", "sw", "se"}
    assert [win_shell.HT_RESIZE_EDGES[e] for e in ("w", "e", "n", "nw", "ne", "s", "sw", "se")] == [
        10, 11, 12, 13, 14, 15, 16, 17,
    ]


@not_win32
def test_begin_native_move_returns_false_without_windows():
    """Не-Windows: нативный перенос недоступен — вызывающий использует ручной путь."""
    fake = FakeWindow()
    assert win_shell.begin_native_move(fake) is False


@not_win32
def test_begin_native_resize_returns_false_without_windows():
    fake = FakeWindow()
    for edge in ("n", "s", "w", "e", "nw", "ne", "sw", "se"):
        assert win_shell.begin_native_resize(fake, edge) is False


def test_begin_native_resize_unknown_edge_returns_false_anywhere():
    """Неизвестная кромка отклоняется мапой HT-кодов до любых Win32-вызовов."""
    fake = FakeWindow()
    assert win_shell.begin_native_resize(fake, "diagonal-fake") is False
    assert win_shell.begin_native_resize(fake, "") is False


# --------------------------------------------------------------------------- #
#  Отложенный нативный перенос/растягивание (2.19.1 — фикс вылета
#  при перетаскивании): модальный цикл стартует из after_idle, а не
#  изнутри Tk-обработчика нажатия.
# --------------------------------------------------------------------------- #
@not_win32
def test_deferred_native_move_returns_false_without_windows():
    """Не-Windows: отложенного переноса нет — вызывающий идёт ручным путём."""
    fake = FakeWindow()
    assert win_shell.begin_native_move_deferred(fake) is False
    assert not any(call[0] == "after_idle" for call in fake.calls)


def test_deferred_native_move_schedules_after_idle(monkeypatch):
    """Windows: цикл переносится в after_idle — стек Tk успевает раскрутиться."""
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    recorded = []
    monkeypatch.setattr(
        win_shell, "_run_native_ht", lambda w, ht: recorded.append((w, ht)),
    )
    fake = FakeWindow()
    assert win_shell.begin_native_move_deferred(fake) is True
    scheduled = [fn for kind, fn in fake.calls if kind == "after_idle"]
    assert len(scheduled) == 1, "ровно один отложенный старт переноса"
    scheduled[0]()  # after_idle срабатывает: кнопка «нажата» (фейк → True)
    assert recorded == [(fake, win_shell.HTCAPTION)]


def test_deferred_native_resize_schedules_after_idle(monkeypatch):
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    recorded = []
    monkeypatch.setattr(
        win_shell, "_run_native_ht", lambda w, ht: recorded.append((w, ht)),
    )
    fake = FakeWindow()
    assert win_shell.begin_native_resize_deferred(fake, "ne") is True
    scheduled = [fn for kind, fn in fake.calls if kind == "after_idle"]
    scheduled[0]()
    assert recorded == [(fake, win_shell.HT_RESIZE_EDGES["ne"])]


def test_deferred_native_resize_unknown_edge_rejected(monkeypatch):
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    fake = FakeWindow()
    assert win_shell.begin_native_resize_deferred(fake, "diagonal-fake") is False
    assert win_shell.begin_native_resize_deferred(fake, "") is False
    assert not any(call[0] == "after_idle" for call in fake.calls)


def test_native_drag_env_switch_disables_all_native_paths(monkeypatch):
    """EWDOCHUB_NO_NATIVE_DRAG=1 глушит и отложенные, и синхронные вызовы."""
    monkeypatch.setenv(win_shell.NO_NATIVE_DRAG_ENV, "1")
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    fake = FakeWindow()
    assert win_shell.begin_native_move_deferred(fake) is False
    assert win_shell.begin_native_resize_deferred(fake, "se") is False
    assert win_shell.begin_native_move(fake) is False
    assert win_shell.begin_native_resize(fake, "w") is False
    assert not any(call[0] == "after_idle" for call in fake.calls)


def test_run_native_ht_skips_when_button_already_released(monkeypatch):
    """Клик без движения не входит в модальный цикл переноса."""
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    monkeypatch.setattr(win_shell, "left_button_pressed", lambda: False)
    called = []
    monkeypatch.setattr(win_shell, "_begin_native_ht", lambda w, ht: called.append(ht))
    win_shell._run_native_ht(FakeWindow(), win_shell.HTCAPTION)
    assert called == []


def test_run_native_ht_starts_loop_when_button_held(monkeypatch):
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    monkeypatch.setattr(win_shell, "left_button_pressed", lambda: True)
    called = []
    monkeypatch.setattr(
        win_shell, "_begin_native_ht", lambda w, ht: called.append((w, ht)),
    )
    fake = FakeWindow()
    win_shell._run_native_ht(fake, win_shell.HTCAPTION)
    assert called == [(fake, win_shell.HTCAPTION)]


def test_run_native_ht_ignores_dead_window_and_failures(monkeypatch):
    """Уничтоженное окно и любые сбои внутри отложенного старта — тихо."""
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    called = []
    monkeypatch.setattr(win_shell, "_begin_native_ht", lambda w, ht: called.append(ht))
    fake = FakeWindow()
    fake.winfo_exists = lambda: False
    win_shell._run_native_ht(fake, win_shell.HTCAPTION)
    assert called == []

    def boom(window, ht_code):
        raise RuntimeError("win32 сбой")

    monkeypatch.setattr(win_shell, "_begin_native_ht", boom)
    win_shell._run_native_ht(FakeWindow(), win_shell.HTCAPTION)  # не бросает


# --------------------------------------------------------------------------- #
#  Папка «Документы» (п. 6.2)
# --------------------------------------------------------------------------- #
@not_win32
def test_documents_known_folder_returns_path(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.delenv("USERPROFILE", raising=False)
    folder = win_shell.documents_known_folder()
    assert folder == os.path.join(str(tmp_path), "Documents")


# --------------------------------------------------------------------------- #
#  Не-Windows: безопасные аналоги Tk
# --------------------------------------------------------------------------- #
@not_win32
def test_workarea_uses_window_screen_size():
    assert win_shell.workarea(FakeWindow()) == (0, 0, 1366, 768)


@not_win32
def test_workarea_without_window_falls_back_to_defaults():
    assert win_shell.workarea(None) == (0, 0, 1920, 1080)


def test_make_borderless_sets_overrideredirect_and_keeps_taskbar_flag():
    fake = FakeWindow()
    win_shell.make_borderless(fake)
    assert ("overrideredirect", True) in fake.calls


@not_win32
def test_minimize_maximize_restore_on_linux_use_tk_analogs():
    fake = FakeWindow()
    win_shell.minimize_window(fake)
    assert ("iconify",) in fake.calls
    win_shell.maximize_window(fake)
    assert ("state", "zoomed") in fake.calls
    win_shell.restore_window(fake)
    # На Linux Tk сам помнит геометрию до разворота: явная геометрия
    # (как в ветке Windows из-за overrideredirect) здесь не нужна.
    assert ("state", "normal") in fake.calls


@not_win32
def test_focus_window_on_linux_forces_focus():
    fake = FakeWindow()
    win_shell.focus_window(fake)
    assert ("focus_force",) in fake.calls


def test_install_window_hooks_is_noop_on_linux():
    """Вне Windows хуки не ставятся и колбэки не требуются."""
    fake = FakeWindow()
    win_shell.install_window_hooks(
        fake,
        on_files=lambda files: None,
        on_dpi_changed=lambda dpi, rect: None,
    )
    assert not any(call[0] == "after" for call in fake.calls)
    assert not hasattr(fake, "_ew_hook")
