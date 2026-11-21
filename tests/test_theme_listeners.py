"""LEAK-1 (D8) — слушатели темы: unregister и автоотписка по владельцу.

Раньше список был append-only: каждое диалоговое окно (Sheet/Alert —
Toplevel на каждое открытие), каждая подсказка ошибки в окне
редактирования и каждая подложка импорта оставляли своё замыкание
навсегда — слушатели накапливались, держали уничтоженные виджеты и
роняли исключения на смене темы. Контракты:

  * register_theme_listener(callback, owner=виджет) — на <Destroy>
    владельца подписка снимается сама;
  * unregister_theme_listener снимает подписку по тождеству;
  * все точки регистрации в коде передают владельца там, где виджет
    может быть уничтожен (диалоги, подсказки, хэндлы, таблица).

Тесты идут на duck-typed владельцах без Tk.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import ui_macos as ui


class FakeOwner:
    """Duck-typed виджет: bind запоминает обработчики Destroy."""

    def __init__(self):
        self.handlers: list = []

    def bind(self, sequence, fn, add=""):
        if sequence == "<Destroy>":
            self.handlers.append(fn)


def isolate_listeners(monkeypatch):
    """Изолирует глобальный список слушателей на время теста."""
    saved = list(ui._LISTENERS)
    ui._LISTENERS[:] = []
    yield
    ui._LISTENERS[:] = saved


@pytest.fixture()
def listeners(monkeypatch):
    yield from isolate_listeners(monkeypatch)


# --------------------------------------------------------------------------- #
#  Регистрация / снятие
# --------------------------------------------------------------------------- #
def test_register_appends_and_unregister_removes(listeners):
    calls: list[str] = []
    callback = lambda: calls.append("x")  # noqa: E731
    ui.register_theme_listener(callback)
    assert ui._LISTENERS == [callback], "подписка добавлена"
    ui.unregister_theme_listener(callback)
    assert ui._LISTENERS == [], "подписка снята по тождеству"


def test_unregister_only_target_callback(listeners):
    first, second = lambda: None, lambda: None  # noqa: E731
    ui.register_theme_listener(first)
    ui.register_theme_listener(second)
    ui.unregister_theme_listener(first)
    assert ui._LISTENERS == [second], "чужие подписки не тронуты"


def test_unregister_unknown_callback_is_noop(listeners):
    ui.register_theme_listener(lambda: None)
    ui.unregister_theme_listener(lambda: None)  # нет такой подписки
    assert len(ui._LISTENERS) == 1


# --------------------------------------------------------------------------- #
#  Автоотписка по владельцу
# --------------------------------------------------------------------------- #
def test_owner_destroy_unregisters_automatically(listeners):
    owner = FakeOwner()
    callback = lambda: None  # noqa: E731
    ui.register_theme_listener(callback, owner=owner)
    assert ui._LISTENERS == [callback], "подписка жива при живом владельце"
    for handler in owner.handlers:
        handler(SimpleNamespace())  # владелец уничтожен (Tk даёт event)
    assert ui._LISTENERS == [], "подписка снята вместе с владельцем (D8)"


def test_owner_without_bind_degrades_to_persistent_subscription(listeners):
    """Duck-typed владелец без bind (тесты): подписка ведёт себя как раньше."""
    callback = lambda: None  # noqa: E731
    ui.register_theme_listener(callback, owner=SimpleNamespace())
    assert ui._LISTENERS == [callback]


def test_stale_callback_does_not_break_theme_switch(listeners):
    """Смена темы переживает сбой в чужом колбэке (контракт set_appearance_mode)."""
    def broken():
        raise RuntimeError("виджет уничтожен")

    healthy_calls: list[str] = []
    ui.register_theme_listener(broken)
    ui.register_theme_listener(lambda: healthy_calls.append("ok"))  # noqa: E731
    ui.set_appearance_mode("dark")
    assert healthy_calls == ["ok"], "остальные слушатели сработали"


# --------------------------------------------------------------------------- #
#  Все не-главнооконные точки регистрации передают владельца
# --------------------------------------------------------------------------- #
def test_dialog_and_transient_registrations_pass_owner():
    """Sheet/Alert/подсказки Card/хэндлы/SpecTable подписываются с owner —
    иначе каждое открытие диалога навсегда оставляет замыкание (D8)."""
    ui_source = Path(ui.__file__).read_text(encoding="utf-8")
    for marker in (
        "register_theme_listener(self._sync_theme, owner=self)",   # TrafficLights,
        "register_theme_listener(self._draw, owner=self)",          # NavItem и др.
        "register_theme_listener(self._sync, owner=self)",          # ThemeToggle
        "register_theme_listener(hint_sync, owner=hint_label)",     # подсказка Card
    ):
        assert marker in ui_source, f"нет автоотписки: {marker}"
    widgets_source = Path(
        Path(ui.__file__).parent / "widgets.py"
    ).read_text(encoding="utf-8")
    assert "register_theme_listener(self._apply_theme, owner=self)" in widgets_source
    main_source = (Path(ui.__file__).parent / "main.py").read_text(encoding="utf-8")
    assert "ui.register_theme_listener(handle_sync, owner=handle)" in main_source
    # Классы диалогов — Toplevel на каждое открытие, у них owner обязателен.
    sheet_block = ui_source[ui_source.index("class Sheet(ctk.CTkToplevel)"):]
    sheet_block = sheet_block[:sheet_block.index("class Alert(ctk.CTkToplevel)")]
    assert "owner=self" in sheet_block, "Sheet подписывается с владельцем"
    alert_block = ui_source[ui_source.index("class Alert(ctk.CTkToplevel)"):]
    assert "register_theme_listener(self._sync_theme, owner=self)" in alert_block, \
        "Alert подписывается с владельцем"
