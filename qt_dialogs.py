"""Модальные диалоги Qt-каркаса — замены ui_macos.Alert/Sheet (этап M2).

Контракты поведения переносятся с Tk-версии дословно:

* ``alert`` — модальное информирование с одной кнопкой; kind задаёт
  иконку (info/warning/error/question);
* ``confirm`` — вопрос «да/нет»; ``destructive=True`` означает, что
  Enter по умолчанию нажимает «Отмена» (U7, UX-A): подтверждение,
  ведущее к потере данных, не срабатывает от случайного Enter;
* ``choice`` — диалог с N вариантами; возвращает ключ выбранного
  варианта, закрытие окна/Escape — ``None``.

Функции не владеют логикой — это замена только «оболочки» диалогов;
smoke-подмены сохраняются на уровне main_qt (маршруты контроллеров —
раздел «Модальные диалоги и smoke-подмены», риски плана).
"""

from __future__ import annotations

from typing import Optional, Sequence

from PySide6.QtWidgets import QMessageBox, QWidget

# kind → иконка QMessageBox (неизвестный kind — информационная).
_ICONS = {
    "info": QMessageBox.Icon.Information,
    "warning": QMessageBox.Icon.Warning,
    "error": QMessageBox.Icon.Critical,
    "question": QMessageBox.Icon.Question,
}


def alert(parent: Optional[QWidget], title: str, message: str,
          kind: str = "info", ok_text: str = "ОК") -> None:
    """Модальный алерт с одной кнопкой; ждёт закрытия (exec)."""
    box = QMessageBox(parent)
    box.setWindowTitle(title or "EW.DocHub")
    box.setIcon(_ICONS.get(kind, QMessageBox.Icon.Information))
    box.setText(message)
    ok = box.addButton(ok_text, QMessageBox.ButtonRole.AcceptRole)
    box.setDefaultButton(ok)
    box.exec()


def confirm(parent: Optional[QWidget], title: str, message: str,
            ok_text: str = "ОК", cancel_text: str = "Отмена",
            destructive: bool = False) -> bool:
    """Модальный вопрос с двумя кнопками; True — нажата главная.

    ``destructive=True``: Enter означает «Отмена» (U7, UX-A) — как в
    Tk-версии. Закрытие окна/Escape — False.
    """
    box = QMessageBox(parent)
    box.setWindowTitle(title or "EW.DocHub")
    box.setIcon(_ICONS["question"])
    box.setText(message)
    ok = box.addButton(ok_text, QMessageBox.ButtonRole.AcceptRole)
    cancel = box.addButton(cancel_text, QMessageBox.ButtonRole.RejectRole)
    box.setDefaultButton(cancel if destructive else ok)
    box.exec()
    return box.clickedButton() is ok


def choice(parent: Optional[QWidget], title: str, message: str,
           options: Sequence[tuple[str, str, bool]]) -> Optional[str]:
    """Модальный диалог с N кнопками (замена ui_macos.choice_dialog).

    ``options``: список ``(ключ, подпись, главная_кнопка)`` слева направо.
    Возвращает ключ выбранного варианта; закрытие окна/Escape — ``None``.
    """
    box = QMessageBox(parent)
    box.setWindowTitle(title or "EW.DocHub")
    box.setIcon(_ICONS["question"])
    box.setText(message)
    buttons: dict[QMessageBox, str] = {}
    default = None
    for key, label, primary in options:
        role = (
            QMessageBox.ButtonRole.AcceptRole
            if primary else QMessageBox.ButtonRole.ActionRole
        )
        button = box.addButton(label, role)
        buttons[button] = key
        if primary and default is None:
            default = button
    if default is not None:
        box.setDefaultButton(default)
    box.exec()
    clicked = box.clickedButton()
    return buttons.get(clicked)
