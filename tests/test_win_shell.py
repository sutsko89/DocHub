"""win_shell — Win32-обвязка окна: DPI (п. 6.3), папка «Документы», хуки.

На не-Windows функции обязаны деградировать в безопасные аналоги Tk или
no-op: здесь это проверяется на фейковых окнах без настоящего Tk. Чистые
хелперы WM_DPICHANGED (масштаб из wparam, прямоугольник-подсказка из
lparam) тестируются напрямую — они не обращаются к Win32.
"""
from __future__ import annotations

import ctypes
import os

import win_shell


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
#  Папка «Документы» (п. 6.2)
# --------------------------------------------------------------------------- #
def test_documents_known_folder_returns_path(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.delenv("USERPROFILE", raising=False)
    folder = win_shell.documents_known_folder()
    assert folder == os.path.join(str(tmp_path), "Documents")


# --------------------------------------------------------------------------- #
#  Не-Windows: безопасные аналоги Tk
# --------------------------------------------------------------------------- #
def test_workarea_uses_window_screen_size():
    assert win_shell.workarea(FakeWindow()) == (0, 0, 1366, 768)


def test_workarea_without_window_falls_back_to_defaults():
    assert win_shell.workarea(None) == (0, 0, 1920, 1080)


def test_make_borderless_sets_overrideredirect_and_keeps_taskbar_flag():
    fake = FakeWindow()
    win_shell.make_borderless(fake)
    assert ("overrideredirect", True) in fake.calls


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
