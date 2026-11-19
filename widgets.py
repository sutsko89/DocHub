"""Составные виджеты интерфейса (п. 9.4).

SpecTable — таблица позиций спецификации: строки с подсветкой, валидный ввод
количества/цены, итоги. Зависит только от примитивов ui_macos и поля F из
document_builder — приложение не тянется.
"""
from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont

import customtkinter as ctk

import ui_macos as ui
from decimal import Decimal
from document_builder import digits_only, parse_decimal


class SpecTable(ctk.CTkFrame):
    """Excel-подобная таблица позиций с поддержкой вставки из буфера.

    Столбцы: Наименование, Кол-во, Ед. изм., Цена за ед. с НДС (руб.).
    Ctrl+V в любой ячейке вставляет прямоугольный блок данных (из Excel —
    столбцы, разделённые табуляцией); при нехватке строки добавляются сами.
    Одинарный клик выделяет строку, Delete удаляет выделенную строку,
    двойной клик по ячейке включает правку.
    """

    HEADERS = ("№", "Наименование", "Кол-во", "Ед. изм.", "Цена за ед. с НДС (руб.)")
    KEYS = ("name", "qty", "unit", "price")
    MIN_ROWS = 10
    MAX_ROWS = 500
    # Ширина фиксированных колонок в логических пикселях (№, Кол-во, Ед. изм., Цена).
    COL_WIDTHS = {0: 34, 2: 64, 3: 84, 4: 150}

    def __init__(self, master, scale: float = 1.0, on_change=None, **kwargs):
        super().__init__(master, fg_color="transparent", corner_radius=0, **kwargs)
        self.scale = scale
        self.on_change = on_change
        self._rows: list[dict] = []
        self._batch = False  # подавление on_change при массовой вставке
        self._selected_row: int | None = None
        self._digits_vcmd = self.register(self._validate_digits)
        self._price_vcmd = self.register(self._validate_price)
        self._entry_font = tkfont.Font(font=ui.font(12))
        # Отмена разрушительных действий с таблицей (п. 4.4).
        self._undo_stack: list[list[dict]] = []
        self._undo_limit = 30

        # -- Панель инструментов ---------------------------------------- #
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, int(4 * self.scale)))
        ui.IconButton(
            toolbar, "plus", self.add_row, size=26, tooltip="Добавить строку",
        ).pack(side="left")
        ui.IconButton(
            toolbar, "minus", self.remove_last, size=26, tooltip="Удалить последнюю строку",
        ).pack(side="left", padx=(int(4 * self.scale), 0))
        ui.IconButton(
            toolbar, "trash", self.clear, size=26, tooltip="Очистить таблицу",
        ).pack(side="left", padx=(int(4 * self.scale), 0))

        # Кол-во строк: задаётся в поле, применяется по Enter.
        self._count_var = tk.StringVar(value=str(self.MIN_ROWS))
        ctk.CTkLabel(
            toolbar, text="Строк:", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"],
        ).pack(side="right")
        count_entry = ui.entry(
            toolbar, width=int(56 * self.scale),
            textvariable=self._count_var,
            validatecommand=(self._digits_vcmd, "%P"),
        )
        count_entry.pack(side="right", padx=(int(8 * self.scale), 0))
        self._count_entry = count_entry
        try:  # фокус внутри внутреннего tk.Entry — биндим именно его
            count_entry.bind_entry("<Return>", lambda _e: self._apply_row_count())
            count_entry.bind_entry("<KP_Enter>", lambda _e: self._apply_row_count())
            count_entry.bind_entry("<Control-z>", self._on_undo)
            count_entry.bind_entry("<Control-Z>", self._on_undo)
        except Exception:
            pass

        # -- Шапка таблицы: ячейки фиксированной ширины, подписи по центру -- #
        header = tk.Frame(self)
        header.pack(fill="x", padx=1)
        header.grid_columnconfigure(1, weight=1)
        for col in (0, 2, 3, 4):
            header.grid_columnconfigure(col, weight=0)
        self._header_frame = header
        self._header_cells: list[tk.Frame] = []
        self._header_labels: list[tk.Label] = []
        for col, title in enumerate(self.HEADERS):
            cell = tk.Frame(header, height=int(30 * self.scale))
            if col in self.COL_WIDTHS:
                cell.configure(width=int(self.COL_WIDTHS[col] * self.scale))
            cell.grid(row=0, column=col, sticky="ew",
                      padx=(2 if col == 0 else 3, 6 if col == 0 else 3))
            cell.pack_propagate(False)
            cell.grid_propagate(False)
            label = tk.Label(
                cell, text=title, font=ui.font(11, "bold"),
                anchor="center", justify="center",
                wraplength=int(150 * self.scale),
            )
            label.pack(fill="both", expand=True)
            self._header_cells.append(cell)
            self._header_labels.append(label)

        # -- Тело таблицы ------------------------------------------------ #
        self.body = ctk.CTkScrollableFrame(
            self, height=int(216 * self.scale),
            fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        self.body.pack(fill="x")
        self.body.grid_columnconfigure(0, weight=0)
        self._schedule_height_update()
        self.body.grid_columnconfigure(1, weight=1)
        for col in (2, 3, 4):
            self.body.grid_columnconfigure(col, weight=0)

        ui.register_theme_listener(self._apply_theme)
        self._apply_theme()

        for _ in range(self.MIN_ROWS):
            self._add_row()

    # -- Строки ----------------------------------------------------------- #
    @property
    def rows(self) -> list[dict]:
        """Строки таблицы: словари со StringVar «name», «qty», «unit», «price»."""
        return self._rows

    def add_row(self, values: dict | None = None) -> dict:
        return self._add_row(values)

    def _entry_width_chars(self, px: int) -> int:
        """Ширина ячейки в символах, соответствующая ширине шапки в пикселях."""
        digit = max(1, self._entry_font.measure("0"))
        return max(3, int((px * self.scale - 10) // digit))

    def _add_row(self, values: dict | None = None) -> dict:
        index = len(self._rows)
        values = values or {}

        num = tk.Label(self.body, text=str(index + 1), font=ui.font(11), anchor="center")
        num.grid(row=index, column=0, sticky="ew",
                 padx=(2, 6), ipady=int(3 * self.scale))
        num.bind("<Button-1>", self._on_cell_click)

        row: dict = {"num": num, "cells": {}}
        for col, key in enumerate(self.KEYS, start=1):
            var = tk.StringVar(value=str(values.get(key, "") or ""))
            # Валидатор «Кол-во»/«Цена» включён флагом validate="key": без него
            # Tk создаёт validatecommand, но НИКОГДА его не вызывает (аудит
            # 2.0.1, P2 «Мёртвые валидаторы») — в поля можно было ввести буквы.
            entry_kwargs = dict(
                textvariable=var, font=ui.font(12),
                relief="flat", justify="left" if key == "name" else "center",
                highlightthickness=1, state="readonly",  # правка — по двойному клику
            )
            if key == "qty":
                entry_kwargs["validate"] = "key"
                entry_kwargs["validatecommand"] = (self._digits_vcmd, "%P")
            elif key == "price":
                entry_kwargs["validate"] = "key"
                entry_kwargs["validatecommand"] = (self._price_vcmd, "%P")
            entry = tk.Entry(self.body, **entry_kwargs)
            if col in self.COL_WIDTHS:  # ширина синхронизирована со шапкой
                entry.configure(width=self._entry_width_chars(self.COL_WIDTHS[col]))
            entry.grid(row=index, column=col, sticky="ew",
                       padx=(0 if col == 1 else 3, 3), ipady=int(3 * self.scale))

            entry.bind("<Button-1>", self._on_cell_click)
            entry.bind("<Double-Button-1>", self._begin_edit)
            entry.bind("<FocusOut>", self._end_edit)
            entry.bind("<Delete>", self._on_delete_key)
            entry.bind("<Control-z>", self._on_undo)
            entry.bind("<Control-Z>", self._on_undo)
            entry.bind("<Return>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<KP_Enter>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<Up>", lambda e: self._move_row(e.widget, -1))
            entry.bind("<Down>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<Tab>", lambda e: self._move_cell(e.widget, +1))
            entry.bind("<ISO_Left_Tab>", lambda e: self._move_cell(e.widget, -1))
            entry.bind("<<Paste>>", self._on_paste)
            try:
                ui.fix_clipboard_shortcuts(entry)
            except Exception:
                pass

            var.trace_add("write", lambda *_a: self._cell_changed())
            row[key] = var
            row["cells"][key] = entry

        self._rows.append(row)
        # П. 5.4: стилизуем только новую строку. Раньше _apply_theme()
        # перестилизовывала ВСЕ строки на каждое добавление — задание
        # 500 строк было O(n²) и блокировало интерфейс.
        if not self._batch:
            self._style_row(index)
            self._schedule_height_update()
        return row

    # -- Высота тела таблицы ------------------------------------------------ #
    def _schedule_height_update(self):
        """Пересчитать высоту таблицы, когда виджеты будут отображены."""
        try:
            self.after_idle(self._update_body_height)
        except Exception:
            pass

    def _update_body_height(self):
        """Высота тела — ровно под строки, без пустого места снизу.

        Если строк больше, чем помещается на экране, высота ограничивается
        максимумом и включается прокрутка.
        """
        try:
            if not self._rows:
                return
            row_h = self._rows[0]["cells"]["name"].winfo_reqheight()
            if row_h <= 1:  # виджеты ещё не отображены — оценка
                row_h = int(24 * self.scale)
            row_h += int(6 * self.scale)  # ipady вокруг ячейки
            total = row_h * len(self._rows) + int(4 * self.scale)
            max_h = int(360 * self.scale)
            self.body.configure(height=max(int(140 * self.scale), min(total, max_h)))
        except Exception:
            pass

    # -- Выделение строки и правка ячеек ------------------------------------- #
    def _row_of_widget(self, widget) -> int | None:
        for i, row in enumerate(self._rows):
            if widget is row["num"] or widget in row["cells"].values():
                return i
        return None

    def _on_cell_click(self, event=None):
        """Одинарный клик: выделить строку (правка — только по двойному клику)."""
        if event is not None:
            index = self._row_of_widget(event.widget)
            if index is not None and index != self._selected_row:
                previous = self._selected_row
                self._selected_row = index
                # Стилизуем только прежнюю и новую строку, а не всю таблицу (п. 5.4).
                if previous is not None:
                    self._style_row(previous)
                self._style_row(index)
        return None

    def _begin_edit(self, event):
        """Двойной клик: разрешить правку ячейки."""
        entry = event.widget
        try:
            entry.configure(state="normal")
            entry.focus_set()
            entry.icursor("end")
        except Exception:
            pass
        return None

    def _end_edit(self, event=None):
        """После выхода из ячейки возвращаем режим «только чтение»."""
        if event is None:
            return None
        entry = event.widget
        try:
            if entry.winfo_exists() and str(entry.cget("state")) == "normal":
                entry.configure(state="readonly")
        except Exception:
            pass
        return None

    def _on_delete_key(self, event):
        """Delete: если ячейка не редактируется — удалить выделенную строку."""
        entry = event.widget
        if str(entry.cget("state")) == "normal":
            return None  # идёт правка — обычное удаление символа
        index = (self._selected_row if self._selected_row is not None
                 else self._row_of_widget(entry))
        if index is not None and 0 <= index < len(self._rows):
            self._delete_row(index)
            return "break"
        return None

    def _delete_row(self, index: int):
        if len(self._rows) <= 1:
            return
        row = self._rows[index]
        if any(row[key].get().strip() for key in self.KEYS):
            if not self._confirm_destructive("Удалить строку с данными?"):
                return
        self._push_undo()
        self._destroy_row(index)
        self._selected_row = None
        self._renumber()
        self._notify()

    def set_row_count(self, count: int):
        """Задаёт точное количество строк (добавляет или удаляет лишние)."""
        count = max(1, min(int(count), self.MAX_ROWS))
        if count == len(self._rows):
            return
        if count < len(self._rows):
            lost = [
                index for index in range(count, len(self._rows))
                if any(self._rows[index][key].get().strip() for key in self.KEYS)
            ]
            if lost and not self._confirm_destructive(
                f"Уменьшение числа строк удалит заполненные строки: {len(lost)}.\n\n"
                "Продолжить?"
            ):
                self._count_var.set(str(len(self._rows)))
                return
            self._push_undo()
        self._batch = True
        try:
            while len(self._rows) < count:
                self._add_row()
            while len(self._rows) > count:
                self._destroy_row(len(self._rows) - 1)
        finally:
            self._batch = False
        self._apply_theme()
        self._schedule_height_update()
        self._selected_row = None
        self._renumber()
        self._count_var.set(str(count))
        self._notify()

    def _apply_row_count(self):
        value = digits_only(self._count_var.get())
        self.set_row_count(int(value or self.MIN_ROWS))

    def _renumber(self):
        for i, row in enumerate(self._rows):
            try:
                row["num"].configure(text=str(i + 1))
            except Exception:
                pass

    def _destroy_row(self, index: int):
        row = self._rows.pop(index)
        for widget in (row["num"], *row["cells"].values()):
            try:
                widget.destroy()
            except Exception:
                pass
        self._schedule_height_update()

    # -- Отмена и подтверждения (п. 4.4) --------------------------------- #
    def _snapshot(self) -> list[dict]:
        return [{key: row[key].get() for key in self.KEYS} for row in self._rows]

    def _push_undo(self) -> None:
        try:
            self._undo_stack.append(self._snapshot())
            if len(self._undo_stack) > self._undo_limit:
                self._undo_stack.pop(0)
        except Exception:
            pass

    def _confirm_destructive(self, message: str) -> bool:
        try:
            return ui.confirm(
                self.winfo_toplevel(), "Таблица позиций", message,
                ok_text="Продолжить", cancel_text="Отмена",
            )
        except Exception:
            return True

    def _on_undo(self, _event=None):
        self.undo()
        return "break"

    def undo(self) -> bool:
        """Возвращает таблицу к состоянию перед последним разрушительным действием."""
        if not self._undo_stack:
            return False
        snapshot = self._undo_stack.pop()
        self._batch = True
        try:
            self._reset_rows()
            while len(self._rows) < len(snapshot):
                self._add_row()
            while len(self._rows) > len(snapshot) and len(self._rows) > self.MIN_ROWS:
                self._destroy_row(len(self._rows) - 1)
            for row, values in zip(self._rows, snapshot):
                for key in self.KEYS:
                    row[key].set(values.get(key, ""))
            self._count_var.set(str(len(self._rows)))
        finally:
            self._batch = False
        self._selected_row = None
        self._renumber()
        self._apply_theme()
        self._schedule_height_update()
        self._notify()
        return True

    def remove_last(self):
        if len(self._rows) <= 1:
            return
        last = self._rows[-1]
        if any(last[key].get().strip() for key in self.KEYS):
            if not self._confirm_destructive("Удалить последнюю строку с данными?"):
                return
        self._push_undo()
        self._destroy_row(len(self._rows) - 1)
        self._selected_row = None
        self._notify()

    def clear(self):
        filled = sum(
            1 for row in self._rows
            if any(row[key].get().strip() for key in self.KEYS)
        )
        if filled and not self._confirm_destructive(
            f"Очистить таблицу? Заполненных строк: {filled}."
        ):
            return
        if filled:
            self._push_undo()
        self._reset_rows()
        self._selected_row = None
        self._notify()

    def _reset_rows(self):
        """Очистка без уведомления (используется перед массовой вставкой)."""
        while self._rows:
            self._destroy_row(len(self._rows) - 1)
        for _ in range(self.MIN_ROWS):
            self._add_row()
        self._apply_theme()
        self._schedule_height_update()

    # -- Данные ------------------------------------------------------------ #
    def get_items(self) -> list[dict]:
        """Позиции с заполненным наименованием; суммы считаются на лету.

        Кол-во и цена — Decimal: цена может содержать копейки,
        кол-во может быть дробным (например, 2,5 м).
        """
        items = []
        for row in self._rows:
            name = row["name"].get().strip()
            if not name:
                continue
            items.append({
                "name": name,
                "qty": parse_decimal(row["qty"].get()),
                "unit": row["unit"].get().strip() or "шт.",
                "price": parse_decimal(row["price"].get()),
            })
        return items

    # Итог таблицы считает document_builder.spec_total (п. 9.8) — один
    # расчёт для предпросмотра, реестра и печатной формы; собственный
    # метод total() удалён как дублирующий (п. 9.5).

    # -- Вставка из буфера -------------------------------------------------- #
    def paste_clipboard(self):
        """Кнопка «Вставить из Excel»: заменить содержимое таблицы буфером."""
        if any(
            row[key].get().strip() for row in self._rows for key in self.KEYS
        ):
            self._push_undo()
        self._reset_rows()
        self._paste_block((0, 0))

    def _on_paste(self, event):
        cell = self._cell_of(event.widget)
        if cell is None:
            return None
        if any(
            row[key].get().strip() for row in self._rows for key in self.KEYS
        ):
            self._push_undo()
        self._paste_block(cell)
        return "break"

    def _paste_block(self, start_cell: tuple[int, int]) -> bool:
        """Вставляет блок текста (TSV) начиная с ячейки start_cell."""
        try:
            raw = self.clipboard_get()
        except Exception:
            return False
        raw = raw.replace("\r\n", "\n").replace("\r", "\n")
        matrix = [line.split("\t") for line in raw.split("\n")]
        while matrix and all(not cell.strip() for cell in matrix[-1]):
            matrix.pop()
        if not matrix:
            return False

        r0, c0 = start_cell
        need = r0 + len(matrix) - len(self._rows)
        for _ in range(max(0, need)):
            self._add_row()

        self._batch = True
        try:
            for dr, values in enumerate(matrix):
                mapped = self._map_row(values)
                for dc, value in enumerate(mapped):
                    col = c0 + dc
                    if col >= len(self.KEYS):
                        break
                    key = self.KEYS[col]
                    # Вставка идёт мимо validate="key" (var.set), поэтому каждое
                    # значение дополнительно прогоняется через правила ячейки
                    # (аудит 2.0.1, P2): мусор в «Кол-во»/«Цена» не попадает.
                    self._rows[r0 + dr][key].set(self._sanitize_value(key, value))
        finally:
            self._batch = False
        self._notify()
        return True

    def _map_row(self, cells: list[str]) -> list[str]:
        """Эвристика раскладки вставленных столбцов по колонкам таблицы."""
        cells = [cell.strip() for cell in cells]
        while cells and not cells[-1]:
            cells.pop()
        n = len(cells)
        if n == 0:
            return ["", "", "", ""]
        if n == 1:
            return [cells[0], "", "", ""]
        if n == 2:  # Наименование, Кол-во
            return [cells[0], self._normalize_number(cells[1]), "", ""]
        if n == 3:
            # «Наименование, Кол-во, Цена» либо «Наименование, Кол-во, Ед. изм.»
            if self._normalize_number(cells[2]):
                return [cells[0], self._normalize_number(cells[1]), "",
                        self._normalize_number(cells[2])]
            return [cells[0], self._normalize_number(cells[1]), cells[2], ""]
        return [cells[0], self._normalize_number(cells[1]), cells[2],
                self._normalize_number(cells[3])]

    @staticmethod
    def _normalize_number(text: str) -> str:
        """«1 000,50» -> «1000.50»: убирает разрядные разделители,
        сохраняет дробную часть (копейки).

        Правила: если встречаются «,» и «.» — правый из них десятичный;
        одиночная запятая считается десятичным разделителем; повторные
        запятые/точки считаются разрядными и удаляются.
        """
        text = str(text).replace("\xa0", "").replace(" ", "")
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "")   # «2.052,59» -> «2052,59»
            else:
                text = text.replace(",", "")   # «2,052.59» -> «2052.59»
        if text.count(",") > 1:
            text = text.replace(",", "")       # «1,000,000» — разрядные
        if text.count(".") > 1:
            text = text.replace(".", "")       # «1.000.000» — разрядные
        text = text.replace(",", ".")
        cleaned = "".join(ch for ch in text if ch.isdigit() or ch == ".")
        return cleaned.rstrip(".")

    # -- Навигация ----------------------------------------------------------- #
    def _cell_of(self, entry) -> tuple[int, int] | None:
        for r, row in enumerate(self._rows):
            for key, widget in row["cells"].items():
                if widget is entry:
                    return (r, self.KEYS.index(key))
        return None

    def _move_row(self, entry, delta: int):
        cell = self._cell_of(entry)
        if cell is None:
            return None
        r, c = cell
        r = min(max(r + delta, 0), len(self._rows) - 1)
        return self._focus_cell((r, c))

    def _move_cell(self, entry, delta: int):
        cell = self._cell_of(entry)
        if cell is None:
            return None
        r, c = cell
        c += delta
        if c >= len(self.KEYS):
            if r + 1 < len(self._rows):
                return self._focus_cell((r + 1, 0))
            self._add_row()
            return self._focus_cell((r + 1, 0))
        if c < 0:
            if r > 0:
                return self._focus_cell((r - 1, len(self.KEYS) - 1))
            c = 0
        return self._focus_cell((r, c))

    def _focus_cell(self, cell: tuple[int, int]):
        r, c = cell
        widget = self._rows[r]["cells"][self.KEYS[c]]
        widget.focus_set()
        widget.icursor("end")
        return "break"

    # -- Служебное ------------------------------------------------------------ #
    @staticmethod
    def _validate_digits(new_value: str) -> bool:
        """Кол-во: цифры и запятая/точка (может быть дробным)."""
        return new_value == "" or all(ch.isdigit() or ch in ",." for ch in new_value)

    @staticmethod
    def _validate_price(new_value: str) -> bool:
        """Цена: цифры, разряды и запятая/точка (копейки)."""
        return new_value == "" or all(
            ch.isdigit() or ch in " .," for ch in new_value
        )

    @classmethod
    def _sanitize_value(cls, key: str, value: str) -> str:
        """Приводит значение к правилам ячейки (вставка мимо validate="key")."""
        if key in ("qty", "price"):
            value = cls._normalize_number(value)
            validator = cls._validate_digits if key == "qty" else cls._validate_price
            return value if validator(value) else ""
        return value

    def _cell_changed(self):
        self._notify()

    def _notify(self):
        if self._batch:
            return
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass

    def _style_row(self, index: int) -> None:
        """Стилизует одну строку (фон, рамки, цвета) с учётом выделения."""
        border = ui.resolve(ui.COLORS["entry_border"])
        accent = ui.resolve(ui.COLORS["accent"])
        base_bg = ui.resolve(ui.COLORS["entry_bg"])
        selected_bg = ui.resolve(ui.COLORS["sidebar_active"])
        bg = selected_bg if index == self._selected_row else base_bg
        row = self._rows[index]
        try:
            row["num"].configure(
                bg=bg,
                fg=ui.resolve(ui.COLORS["text_tertiary"]),
            )
        except Exception:
            pass
        for entry in row["cells"].values():
            try:
                entry.configure(
                    bg=bg,
                    readonlybackground=bg,
                    fg=ui.resolve(ui.COLORS["text"]),
                    insertbackground=ui.resolve(ui.COLORS["text"]),
                    highlightbackground=border,
                    highlightcolor=accent,
                )
            except Exception:
                pass

    def _apply_theme(self):
        header_bg = ui.resolve(ui.COLORS["chrome"])
        for cell in self._header_cells:
            try:
                cell.configure(bg=header_bg)
            except Exception:
                pass
        for label in self._header_labels:
            try:
                label.configure(bg=header_bg, fg=ui.resolve(ui.COLORS["text_secondary"]))
            except Exception:
                pass
        for index in range(len(self._rows)):
            self._style_row(index)
