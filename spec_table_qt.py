"""Таблица позиций спецификации на Qt (этап M3 миграции, CP1).

Qt-перенос Tk-виджета widgets.SpecTable с сохранением контрактов:

* **Вставка блока из Excel** (PERF-1, D2/R-08): Ctrl+V в ячейке вставляет
  TSV-блок начиная с этой ячейки; лимит ``MAX_ROWS = 500`` с учётом
  стартовой строки — один ``spec_table_core.paste_plan`` с Tk-таблицей;
  обрезка не молчит: сигнал ``notice(str)`` несёт «Вставлено N,
  отброшено K …», вставка ровно до лимита — без сообщения.
* **Правка ячеек с валидацией**: правка по двойному клику (EditTrigger),
  «Кол-во»/«Цена» валидируются по правилам ядра; вставка идёт мимо
  валидатора ввода, поэтому каждое значение прогоняется через
  ``spec_table_core.sanitize_value`` (аудит 2.0.1, P2) — мусор в числовые
  колонки не попадает.
* **Отмена и подтверждения** (п. 4.4): снапшот перед деструктивным
  действием (удаление/очистка/уменьшение числа строк/вставка поверх
  заполненной), undo по Ctrl+Z, подтверждение через callback
  ``confirm_destructive`` (задаёт каркас: qt-диалог с семантикой U7).
* **Итог** (U12): виджет только хранит строки — суммы считает
  ``document_builder.spec_total`` на стороне контроллера; сигнал
  ``changed`` срабатывает один раз на массовую операцию (``_batch``).

Никаких Tk-импортов; зависимость от интерфейса — только сигналы Qt.
"""

from __future__ import annotations

import re

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeySequence, QRegularExpressionValidator, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QHeaderView, QStyledItemDelegate,
    QTableWidget, QTableWidgetItem, QWidget,
)

import spec_table_core as core

# Ширины колонок повторяют Tk COL_WIDTHS (№, Кол-во, Ед. изм., Цена).
_COLUMN_WIDTHS = {
    0: core.COL_WIDTHS[0], 2: core.COL_WIDTHS[2],
    3: core.COL_WIDTHS[3], 4: core.COL_WIDTHS[4],
}
_QTY_RE = re.compile(r"^[\d,.]*$")
_PRICE_RE = re.compile(r"^[\d .,]*$")
_DATA_COLUMNS = (1, 2, 3, 4)  # name, qty, unit, price


class SpecCellDelegate(QStyledItemDelegate):
    """Редактор ячейки с валидатором «Кол-во»/«Цена» (правила Tk-таблицы)."""

    def createEditor(self, parent, option, index):  # noqa: N802 — имя Qt
        editor = super().createEditor(parent, option, index)
        column = index.column()
        if column == 2:  # Кол-во
            editor.setValidator(QRegularExpressionValidator(_QTY_RE, editor))
        elif column == 4:  # Цена
            editor.setValidator(QRegularExpressionValidator(_PRICE_RE, editor))
        return editor


class SpecTableQt(QTableWidget):
    """Таблица позиций: строки 10..500, вставка Excel, undo, подтверждения."""

    #: любое пользовательское изменение содержимого (не при batch-операции)
    changed = Signal()
    #: немодальное сообщение (PERF-1 «Вставлено N, отброшено K»)
    notice = Signal(str)

    def __init__(self, parent: QWidget | None = None,
                 confirm_destructive=None):
        super().__init__(0, len(core.HEADERS), parent)
        # Подтверждение деструктивных действий задаёт каркас; None —
        # фолбэк Tk-виджета (сбой диалога не должен молча терять данные).
        self._confirm_destructive = confirm_destructive
        self._batch = False
        self._undo_stack: list[tuple[tuple[str, ...], ...]] = []
        self._undo_limit = 30

        self.setHorizontalHeaderLabels(list(core.HEADERS))
        self.verticalHeader().setVisible(False)
        self.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        # Правка — по двойному клику; одинарный клик только выделяет строку.
        self.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.setItemDelegate(SpecCellDelegate(self))
        header = self.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for column, width in _COLUMN_WIDTHS.items():
            self.setColumnWidth(column, width)

        # Пользовательская правка ячейки → changed (программные изменения
        # внутри массовых операций подавлены флагом _batch).
        self.itemChanged.connect(self._on_item_changed)

        # Стартовые строки — как Tk (MIN_ROWS).
        self._batch = True
        try:
            for _ in range(core.MIN_ROWS):
                self._append_row()
        finally:
            self._batch = False

        self._undo_shortcut = QShortcut(QKeySequence.StandardKey.Undo, self)
        self._undo_shortcut.activated.connect(self.undo)

    # ------------------------------------------------------------------ #
    #  Строки и данные
    # ------------------------------------------------------------------ #
    def row_values(self) -> list[dict]:
        """Значения всех строк: словари «name», «qty», «unit», «price»."""
        out = []
        for row in range(self.rowCount()):
            out.append({
                "name": self._text(row, 1),
                "qty": self._text(row, 2),
                "unit": self._text(row, 3),
                "price": self._text(row, 4),
            })
        return out

    def get_items(self) -> list[dict]:
        """Позиции с заполненным наименованием (Decimal qty/price) — ядро."""
        return core.items_from_values(self.row_values())

    def snapshot(self) -> tuple[tuple[str, ...], ...]:
        """Снимок строк для подписи формы (п. 4.4) и undo."""
        return tuple(
            (row["name"], row["qty"], row["unit"], row["price"])
            for row in self.row_values()
        )

    def _text(self, row: int, column: int) -> str:
        item = self.item(row, column)
        return item.text() if item is not None else ""

    def _set_text(self, row: int, column: int, value: str) -> None:
        item = QTableWidgetItem(value)
        if column == 0:
            item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setItem(row, column, item)

    def _on_item_changed(self, _item) -> None:
        if not self._batch:
            self.changed.emit()

    def _append_row(self, values: dict | None = None) -> None:
        row = self.rowCount()
        self.insertRow(row)
        values = values or {}
        self._set_text(row, 0, str(row + 1))
        for column, key in enumerate(core.KEYS, start=1):
            self._set_text(row, column, str(values.get(key, "") or ""))

    def add_row(self, values: dict | None = None) -> None:
        """Добавляет строку (кнопка «+», расширение под вставку)."""
        self._append_row(values)
        self._renumber()
        if not self._batch:
            self.changed.emit()

    # ------------------------------------------------------------------ #
    #  Массовые операции (паритет Tk-методов)
    # ------------------------------------------------------------------ #
    def set_row_count(self, count: int) -> None:
        """Задаёт точное количество строк (1..MAX_ROWS)."""
        count = max(1, min(int(count), core.MAX_ROWS))
        if count == self.rowCount():
            return
        if count < self.rowCount():
            lost = [
                index for index in range(count, self.rowCount())
                if any(self._text(index, col).strip() for col in _DATA_COLUMNS)
            ]
            if lost and not self._confirm(
                "Уменьшение числа строк удалит "
                f"заполненные строки: {len(lost)}.\n\nПродолжить?"
            ):
                return
            self._push_undo()
        self._batch = True
        try:
            while self.rowCount() < count:
                self._append_row()
            while self.rowCount() > count:
                self.removeRow(self.rowCount() - 1)
            self._renumber()
        finally:
            self._batch = False
        self.changed.emit()

    def remove_last(self) -> None:
        """«−»: удалить последнюю строку (с подтверждением, если заполнена)."""
        if self.rowCount() <= 1:
            return
        last = self.rowCount() - 1
        if any(self._text(last, col).strip() for col in _DATA_COLUMNS):
            if not self._confirm("Удалить последнюю строку с данными?"):
                return
        self._push_undo()
        self._batch = True
        try:
            self.removeRow(last)
            self._renumber()
        finally:
            self._batch = False
        self.changed.emit()

    def clear(self) -> None:
        """Очистить таблицу (с подтверждением, если есть заполненные строки)."""
        filled = sum(
            1 for row in range(self.rowCount())
            if any(self._text(row, col).strip() for col in _DATA_COLUMNS)
        )
        if filled and not self._confirm(
            f"Очистить таблицу? Заполненных строк: {filled}."
        ):
            return
        if filled:
            self._push_undo()
        self._reset_rows()
        self.changed.emit()

    def _reset_rows(self) -> None:
        """Очистка до MIN_ROWS без уведомления (перед массовой вставкой).

        Флаг ``_batch`` внешней операции сохраняется: undo/вставка
        продолжают подавлять сигнал до своего завершения.
        """
        outer = self._batch
        self._batch = True
        try:
            while self.rowCount() > core.MIN_ROWS:
                self.removeRow(self.rowCount() - 1)
            while self.rowCount() < core.MIN_ROWS:
                self._append_row()
            for row in range(self.rowCount()):
                for column in _DATA_COLUMNS:
                    self._set_text(row, column, "")
            self._renumber()
        finally:
            self._batch = outer

    def _delete_row(self, row: int) -> None:
        """Delete на ячейке: удалить строку (с подтверждением, если заполнена)."""
        if self.rowCount() <= 1:
            return
        if any(self._text(row, col).strip() for col in _DATA_COLUMNS):
            if not self._confirm("Удалить строку с данными?"):
                return
        self._push_undo()
        self._batch = True
        try:
            self.removeRow(row)
            self._renumber()
        finally:
            self._batch = False
        self.changed.emit()

    def _renumber(self) -> None:
        for row in range(self.rowCount()):
            item = self.item(row, 0)
            if item is not None and item.text() != str(row + 1):
                item.setText(str(row + 1))

    # ------------------------------------------------------------------ #
    #  Отмена и подтверждения (п. 4.4)
    # ------------------------------------------------------------------ #
    def _confirm(self, message: str) -> bool:
        if self._confirm_destructive is None:
            return True
        try:
            return bool(self._confirm_destructive(message))
        except Exception:
            return True  # фолбэк Tk-виджета: сбой диалога не блокирует

    def _push_undo(self) -> None:
        self._undo_stack.append(self.snapshot())
        if len(self._undo_stack) > self._undo_limit:
            self._undo_stack.pop(0)

    def undo(self) -> bool:
        """Возвращает таблицу к состоянию перед последним деструктивным действием."""
        if not self._undo_stack:
            return False
        snapshot = self._undo_stack.pop()
        self._batch = True
        try:
            self._reset_rows()
            while self.rowCount() < len(snapshot):
                self._append_row()
            while self.rowCount() > len(snapshot) and self.rowCount() > core.MIN_ROWS:
                self.removeRow(self.rowCount() - 1)
            for row, values in enumerate(snapshot):
                self._set_text(row, 1, values[0])
                self._set_text(row, 2, values[1])
                self._set_text(row, 3, values[2])
                self._set_text(row, 4, values[3])
            self._renumber()
        finally:
            self._batch = False
        self.changed.emit()
        return True

    # ------------------------------------------------------------------ #
    #  Вставка из буфера (PERF-1)
    # ------------------------------------------------------------------ #
    @staticmethod
    def _clipboard_text() -> str | None:
        text = QApplication.clipboard().text()
        return text if text else None

    def paste_clipboard(self) -> bool:
        """Заменить содержимое таблицы блоком из буфера (целиком)."""
        if self._has_data():
            self._push_undo()
        self._reset_rows()
        return self._paste_block((0, 0))

    def paste_block_at(self, row: int, column: int) -> bool:
        """Вставить блок из буфера начиная с ячейки (Ctrl+V в таблице).

        ``column`` — виджет-колонка (1..4, как current.column());
        внутри — индекс по core.KEYS (0..3), как в Tk _cell_of.
        """
        if self._has_data():
            self._push_undo()
        return self._paste_block((row, max(0, min(column, 4) - 1)))

    def _has_data(self) -> bool:
        return any(
            self._text(row, col).strip()
            for row in range(self.rowCount()) for col in _DATA_COLUMNS
        )

    def _paste_block(self, start_cell: tuple[int, int]) -> bool:
        """Вставляет блок TSV начиная с ячейки; лимит и сообщение — ядро.

        PERF-1 (D2/R-08): вставка ограничена MAX_ROWS с учётом стартовой
        строки r0 — как поле «Строк:». Обрезанные строки не теряются
        незаметно: notice называет оба числа («Вставлено 500, отброшено
        100»), данные остаются в источнике. Ровно до лимита — без сообщения.
        """
        raw = self._clipboard_text()
        if raw is None:
            return False
        raw = raw.replace("\r\n", "\n").replace("\r", "\n")
        matrix = [line.split("\t") for line in raw.split("\n")]
        while matrix and all(not cell.strip() for cell in matrix[-1]):
            matrix.pop()
        if not matrix:
            return False

        r0, c0 = start_cell
        usable, dropped = core.paste_plan(matrix, r0, core.MAX_ROWS)
        need = r0 + len(usable) - self.rowCount()
        self._batch = True
        try:
            for _ in range(max(0, need)):
                self._append_row()
            for dr, cells in enumerate(usable):
                mapped = core.map_row(cells)
                for dc, value in enumerate(mapped):
                    col = c0 + dc
                    if col >= len(core.KEYS):
                        break
                    key = core.KEYS[col]
                    # Вставка идёт мимо валидатора ввода, поэтому каждое
                    # значение дополнительно прогоняется через правила
                    # ячейки (аудит 2.0.1, P2): мусор в «Кол-во»/«Цена»
                    # не попадает.
                    self._set_text(r0 + dr, col + 1, core.sanitize_value(key, value))
            self._renumber()
        finally:
            self._batch = False
        if dropped:
            self.notice.emit(
                core.paste_notice_text(len(usable), dropped, core.MAX_ROWS)
            )
        self.changed.emit()
        return True

    # ------------------------------------------------------------------ #
    #  Клавиатура: Delete — удаление строки, Ctrl+V — блок из буфера
    # ------------------------------------------------------------------ #
    def keyPressEvent(self, event) -> None:  # noqa: N802 — имя Qt
        # Правка ячейки: редактор обрабатывает Delete/Ctrl+V штатно
        # (удаление символа, вставка в ячейку).
        if self.state() == QAbstractItemView.State.EditingState:
            super().keyPressEvent(event)
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            current = self.currentItem()
            row = current.row() if current is not None else 0
            column = current.column() if current is not None else 0
            self.paste_block_at(row, column)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Delete:
            current = self.currentItem()
            if current is not None:
                self._delete_row(current.row())
                event.accept()
                return
        super().keyPressEvent(event)
