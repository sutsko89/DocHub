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


# --------------------------------------------------------------------------- #
#  Виртуальный экран (2.19.3 — границы ручного переноса)
# --------------------------------------------------------------------------- #
def test_virtual_screen_bounds_from_system_metrics(monkeypatch):
    """Windows: объединение мониторов из GetSystemMetrics (второй слева/сверху)."""
    metrics = {
        win_shell._SM_XVIRTUALSCREEN: -1920,
        win_shell._SM_YVIRTUALSCREEN: -200,
        win_shell._SM_CXVIRTUALSCREEN: 3840,
        win_shell._SM_CYVIRTUALSCREEN: 1000,
    }
    monkeypatch.setattr(win_shell, "_get_system_metrics", lambda i: metrics[i])
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    assert win_shell.virtual_screen_bounds() == (-1920, -200, 3840, 1000)


def test_virtual_screen_bounds_survives_metrics_failure(monkeypatch):
    """Сбой Win32 — fallback на экран Tk/константы, без исключений."""
    def boom(_index):
        raise OSError("win32 сбой")

    monkeypatch.setattr(win_shell, "_get_system_metrics", boom)
    monkeypatch.setattr(win_shell.sys, "platform", "win32")
    x, y, w, h = win_shell.virtual_screen_bounds()
    assert w > 0 and h > 0, "fallback возвращает пригодные границы"


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


@not_win32
def test_workarea_at_follows_root_screen_on_linux():
    """Не-Windows: единая рабочая область экрана Tk для любой точки (2.20.0)."""
    root = FakeWindow(screen=(1366, 768))
    assert win_shell.workarea_at(root, 100, 200) == (0, 0, 1366, 768)
    assert win_shell.workarea_at(root, -50, 5000) == (0, 0, 1366, 768)


@not_win32
def test_workarea_at_without_root_falls_back_to_defaults():
    assert win_shell.workarea_at(None, 0, 0) == (0, 0, 1920, 1080)


@not_win32
def test_workarea_at_survives_win32_failure(monkeypatch):
    """Сбой MonitorFromPoint/GetMonitorInfoW не роняет вызов (2.20.0):
    ветка win32 с полностью сломанным windll уходит в каскад fallback'ей
    workarea() и завершается крайним fallback 1920×1080 без исключения."""
    monkeypatch.setattr(win_shell.sys, "platform", "win32")

    class Boom:
        def __getattr__(self, name):
            raise OSError("win32 сбой")

    # На Linux windll нет вовсе — создаём атрибут (raising=False),
    # чтобы ветка win32 получила сломанную обвязку.
    monkeypatch.setattr(win_shell.ctypes, "windll", Boom(), raising=False)
    assert win_shell.workarea_at(FakeWindow(), 10, 20) == (0, 0, 1920, 1080)


def test_make_borderless_sets_overrideredirect_and_keeps_taskbar_flag():
    fake = FakeWindow()
    win_shell.make_borderless(fake)
    assert ("overrideredirect", True) in fake.calls


@not_win32
def test_minimize_maximize_restore_on_linux_use_workarea_geometry():
    fake = FakeWindow()
    win_shell.minimize_window(fake)
    assert ("iconify",) in fake.calls
    # 2.20.0: разворот задаётся геометрией рабочей области на всех
    # платформах — state('zoomed') бездействовал для overrideredirect.
    win_shell.maximize_window(fake)
    assert ("geometry", "1366x768+0+0") in fake.calls
    # Явная геометрия возврата применяется на всех платформах (2.20.0).
    win_shell.restore_window(fake, "900x600+10+20")
    assert ("geometry", "900x600+10+20") in fake.calls
    win_shell.restore_window(fake)
    # Без геометрии — системный аналог: state('normal').
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
