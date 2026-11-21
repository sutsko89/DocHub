"""UX-A — живые виджеты (требуют дисплей/Xvfb, как test_tk_page_build).

Покрывает Tk-часть карточки UX-A:

  * U4 — подсказка формата в поле: надпись-подложка показывается в пустом
    ненаходящемся в фокусе поле, скрывается при фокусе/вводе и никогда
    не попадает в значение (get() пустого поля — пустая строка);
  * U5 — scroll_to_widget/containing_scroller: прокрутка к ошибочному
    полю внутри CTkScrollableFrame;
  * U7 — Alert: в нейтральном подтверждении Return нажимает главную
    кнопку, в деструктивном остаётся «Отменой».

Страницы документов с блоком «Поиск по ИНН» (U3) проверяются в
test_tk_page_build.py на собранном целиком ContractApp.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

tk = pytest.importorskip("tkinter")

try:
    _probe = tk.Tk()
    _probe.withdraw()
    _probe.destroy()
    TK_AVAILABLE = True
except tk.TclError:  # pragma: no cover — нет дисплея в CI без Xvfb
    TK_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not TK_AVAILABLE, reason="нет графического дисплея (нужен Xvfb)")

import customtkinter as ctk

import ui_macos as ui


@pytest.fixture()
def root():
    root = tk.Tk()
    # Без withdraw: скрытое окно не получает геометрию (winfo_height() == 1,
    # виджеты не мапятся) — прокрутка и place не поддаются проверке.
    root.geometry("480x360+0+0")
    yield root
    try:
        root.update_idletasks()
        root.destroy()
    except tk.TclError:
        pass


# --------------------------------------------------------------------------- #
#  U4: подсказка формата в поле
# --------------------------------------------------------------------------- #
def test_u4_placeholder_label_shown_on_empty_unfocused_field(root):
    field = ui.entry(root, placeholder="10 или 12 цифр")
    field.pack()
    root.update()
    assert field._placeholder_label is not None
    assert field._placeholder_label.cget("text") == "10 или 12 цифр"
    assert field.get() == "", "подсказка попала в значение поля"


def test_u4_placeholder_hidden_while_focused(root):
    field = ui.entry(root, placeholder="9 цифр")
    field.pack()
    root.update()
    field.entry.focus_force()
    root.update()
    field._update_placeholder()
    root.update()
    # В фокусе подложка спрятана, даже когда поле пусто.
    assert not field._placeholder_label.winfo_ismapped()


def test_u4_placeholder_hidden_after_typing(root):
    field = ui.entry(root, placeholder="20 цифр")
    field.pack()
    root.update()
    field.entry.focus_force()
    field.insert(0, "40702810123456789012")
    field._update_placeholder()  # KeyRelease не срабатывает от programmatic insert
    root.update()
    assert field.get() == "40702810123456789012"
    assert not field._placeholder_label.winfo_ismapped()


def test_u4_placeholder_returns_after_clearing(root):
    field = ui.entry(root, placeholder="13 цифр")
    field.pack()
    root.update()
    field.insert(0, "1037739010221")
    field._update_placeholder()
    field.delete(0, tk.END)
    root.focus_set()  # снять фокус с поля
    root.update()
    field._update_placeholder()
    root.update()
    assert field.get() == "", "после очистки поле не содержит подсказку"
    assert field._placeholder_label.winfo_ismapped(), \
        "подложка не вернулась после очистки поля"


def test_u4_placeholder_removal_destroys_label(root):
    field = ui.entry(root, placeholder="9 цифр")
    field.pack()
    root.update()
    label = field._placeholder_label
    field.set_placeholder(None)
    root.update()
    assert field._placeholder_label is None
    assert not label.winfo_exists(), "подложка не уничтожена после снятия подсказки"


# --------------------------------------------------------------------------- #
#  U5: прокрутка к ошибочному полю
# --------------------------------------------------------------------------- #
def test_u5_containing_scroller_finds_frame(root):
    scroller = ctk.CTkScrollableFrame(root, height=120)
    scroller.pack(fill="both", expand=True)
    card = ui.Card(scroller, compact=True)
    card.pack(fill="x")
    field = ui.entry(card.body, placeholder="9 цифр")
    card.row("БИК", field)
    root.update()
    found = ui.containing_scroller(field)
    assert found is scroller, "хелпер не нашёл ближайший прокручиваемый кадр"


def test_u5_containing_scroller_without_scroller(root):
    frame = ctk.CTkFrame(root)
    label = ctk.CTkLabel(frame, text="—")
    label.pack()
    assert ui.containing_scroller(label) is None
    assert ui.scroll_to_widget(None, label) is False
    assert ui.scroll_to_widget(frame, None) is False


def test_u5_scroll_moves_view_to_widget(root):
    scroller = ctk.CTkScrollableFrame(root, height=140)
    scroller.pack(fill="both", expand=True)
    fields = []
    for i in range(20):
        field = ui.entry(scroller)
        field.pack(pady=8)
        fields.append(field)
    root.update()
    scroller._parent_canvas.yview_moveto(0.0)
    root.update()
    target = fields[15]  # далеко за нижней границей видимой области
    moved = ui.scroll_to_widget(scroller, target)
    root.update()
    assert moved, "scroll_to_widget не выполнил прокрутку"
    top = scroller._parent_canvas.yview()[0]
    assert top > 0.4, f"вид не сдвинулся к ошибочному полю (yview={top:.2f})"


def test_u5_scroll_is_noop_for_widget_at_top(root):
    scroller = ctk.CTkScrollableFrame(root, height=140)
    scroller.pack(fill="both", expand=True)
    field = ui.entry(scroller)
    field.pack(pady=8)
    root.update()
    scroller._parent_canvas.yview_moveto(0.5)
    root.update()
    assert ui.scroll_to_widget(scroller, field)
    root.update()
    top = scroller._parent_canvas.yview()[0]
    assert top == pytest.approx(0.0, abs=0.05), \
        "поле у верха содержимого должно прокручиваться к началу"


# --------------------------------------------------------------------------- #
#  U7: Enter в подтверждениях
# --------------------------------------------------------------------------- #
def test_u7_neutral_confirm_return_presses_ok(root):
    dialog = ui.Alert(root, "Повтор номера", "Создать ещё один документ?",
                      "question", "Создать ещё", "Отмена")
    dialog.deiconify()
    dialog.focus_force()
    dialog.update()
    assert dialog._destructive is False
    # Return-привязка нажимает главную кнопку.
    dialog.event_generate("<Return>")
    dialog.update()
    assert dialog._result is True, "Enter в нейтральном подтверждении не нажал «ОК»"


def test_u7_destructive_confirm_return_stays_cancel(root):
    dialog = ui.Alert(root, "Удалить навсегда", "Удалить запись окончательно?",
                      "question", "Удалить навсегда", "Отмена", destructive=True)
    dialog.deiconify()
    dialog.focus_force()
    dialog.update()
    assert dialog._destructive is True
    dialog.event_generate("<Return>")
    dialog.update()
    assert dialog._result is False, "Enter нажал деструктивную главную кнопку"
    dialog._ok()
    assert dialog._result is True, "кнопка «Удалить навсегда» продолжает работать"


def test_u7_escape_always_cancels(root):
    dialog = ui.Alert(root, "Повтор номера", "Создать ещё один документ?",
                      "question", "Создать ещё", "Отмена")
    dialog.deiconify()
    dialog.focus_force()
    dialog.update()
    dialog.event_generate("<Escape>")
    dialog.update()
    assert dialog._result is False, "Escape перестал отменять"
