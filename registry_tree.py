"""Список реестра на ttk.Treeview (этап 4 roadmap, TREE-1 + UX-REG).

Замена перерисовке строк из CTk-виджетов (П-аудит P-02: ~10 виджетов на
строку, 8–9 мс/строка, 10k — минуты): один ttk.Treeview со стилизацией
под светлую/тёмную тему (R-03), карта id->iid и точечные операции
insert/delete/item по дифу модели строк (R-04).

Контракты карточки TREE-1:

* поиск и сортировка всегда выполняются по всей базе (в контроллере) —
  виджет получает готовую упорядоченную модель строк и не знает про
  фильтры; лимит отображения (R-05b) применяется в контроллере;
* полная популяция — только при смене фильтра, сортировки и
  drill-переходах; изменение состава и правки записей идут дифом:
  строки, которых изменение не касается, сохраняют iid (не пересоздаются);
* теги сроков «expired»/«soon» красят строку документа красным/амберным
  (те же цвета, что были у CTk-строк);
* выделение — подложка через ttk.Style (V2); точка типа записи в
  нулевой колонке — PhotoImage-точка цвета типа (V4-пилот: ttk.Treeview
  не красит отдельную ячейку, изображение в #0 — дешёвая альтернатива);
* прокрутка — нативная для Treeview, полоса скрывается, когда содержимое
  помещается целиком (как у CTkScrollableFrame в macOS-стиле).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Iterable

import tkinter as tk
from tkinter import ttk

import ui_macos as ui


#: Мягкий лимит отображения (R-05b): порция «Показать ещё». Значение из
#: прототип-замера (tools/bench_tree_prototype.py): популяция 10k — 86 мс
#: при пороге раздела Е 300 мс; 50k — 470 мс, поэтому объёмы до 10k
#: показываются целиком, большие — порциями по 10k с «Показать ещё».
DISPLAY_SOFT_LIMIT = 10_000

#: iid служебной строки «Показать ещё».
SHOW_MORE_IID = "show-more"

#: Вертикальный отступ строки (логические пиксели).
ROW_HEIGHT = 30


@dataclass
class RowSpec:
    """Строка модели реестра: iid, значения колонок, теги, точка типа."""

    iid: str
    values: tuple[str, ...]
    tags: tuple[str, ...] = ()
    dot: str | None = None  # цвет точки типа (V4); None — точки нет
    data: dict = field(default_factory=dict, repr=False)


@dataclass
class DiffPlan:
    """План диф-обновления (чистая функция plan_diff)."""

    to_delete: list[str]
    ops: list[tuple[int, str, str]]  # (позиция, iid, "insert"|"move")
    updates: list[RowSpec]


def plan_display_count(total: int, shown: int) -> tuple[int, bool]:
    """Сколько строк показывать и есть ли ещё (R-05b).

    ``shown`` — сколько строк уже разрешено показывать (копится
    кнопкой «Показать ещё»). Возвращает (количество строк, есть_ли_скрытые).
    Чистая функция — тестируется без Tk.
    """
    shown = max(shown, DISPLAY_SOFT_LIMIT)
    shown = min(shown, total) if total else 0
    return shown, total > shown


def plan_diff(current: Iterable[str], desired: list[RowSpec]) -> DiffPlan:
    """Диф модели строк без Tk: (удалить, операции, обновить).

    ``current`` — iid видимых строк в порядке отображения. Алгоритм:
    префикс, совпадающий с желаемой моделью, не трогается вовсе (критерий
    5.1 — «выбор/изменение не пересоздаёт строки»); правится только
    суффикс после первой расходимости — вставки и перемещения применяются
    в порядке возрастания позиции, поэтому каждая операция видит все
    строки до неё уже стоящими правильно.
    """
    current = list(current)
    # ВАЖНО: множество желаемых iid строится ОДИН раз. Внутри comprehension
    # оно пересобиралось бы на каждой итерации — O(n²) (на 10k это 5,8 с
    # вместо 2 мс, найдено профилированием прототипа).
    desired_iids = {row.iid for row in desired}
    to_delete = [iid for iid in current if iid not in desired_iids]
    deleted = set(to_delete)
    sim = [iid for iid in current if iid not in deleted]
    pos_of = {iid: idx for idx, iid in enumerate(sim)}

    ops: list[tuple[int, str, str]] = []  # (позиция, iid, "insert"|"move")
    k = 0
    while (k < len(sim) and k < len(desired)
           and sim[k] == desired[k].iid):
        k += 1
    for pos in range(k, len(desired)):
        iid = desired[pos].iid
        if iid in pos_of:
            if pos_of[iid] != pos:
                # Сдвиг в симуляции: как у tree.move — снять и вставить.
                sim.pop(pos_of[iid])
                sim.insert(pos, iid)
                lo, hi = min(pos_of[iid], pos), max(pos_of[iid], pos)
                for shift in range(lo, hi + 1):
                    pos_of[sim[shift]] = shift
                ops.append((pos, iid, "move"))
        else:
            sim.insert(pos, iid)
            for shift in range(pos, len(sim)):
                pos_of[sim[shift]] = shift
            ops.append((pos, iid, "insert"))
    return DiffPlan(to_delete=to_delete, ops=ops, updates=list(desired))


class RegistryTree(ttk.Frame):
    """Список реестра: Treeview + вертикальный скролл + диф-обновления."""

    def __init__(self, master, on_select: Callable[[str | None, bool, bool], None],
                 on_activate: Callable[[str], None],
                 columns: Iterable[tuple[str, str, int, str, bool]]):
        super().__init__(master)
        self._on_select = on_select
        self._on_activate = on_activate
        self._key_nav = False
        self._toggle_after = None  # отложенное снятие выделения повторным кликом
        self._rows: dict[str, RowSpec] = {}
        self._dots: dict[str, tk.PhotoImage] = {}

        columns = tuple(columns)
        self._columns = columns

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass  # тема недоступна — остаётся системная
        self._style = style
        self._style_name = "Registry.Treeview"

        self.tree = ttk.Treeview(
            self, columns=[col[0] for col in columns], show="tree headings",
            selectmode="browse", style=self._style_name,
        )
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self._vsb = vsb
        self.tree.configure(yscrollcommand=self._on_yview)
        vsb.grid(row=0, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._apply_columns()
        self._apply_style()
        # Перерисовка при смене темы; подписка снимается вместе с виджетом
        # (LEAK-1: owner-подписка).
        ui.register_theme_listener(self._apply_style, owner=self)

        self.tree.bind("<<TreeviewSelect>>", self._handle_select, add="+")
        # Клавиатурная навигация (U13): стрелки двигают выделение без toggle —
        # флаг ставится до класс-обработчика, который меняет выделение.
        for key in ("<Down>", "<Up>", "<Home>", "<End>", "<Prior>", "<Next>"):
            self.tree.bind(key, self._mark_key_nav, add="+")
        # Повторный клик по выделенной строке снимает выделение (2.18.0):
        # в browse-режиме Tk не генерирует событие при неизменном выделении,
        # поэтому toggle обрабатывается в обработчике нажатия.
        self.tree.bind("<Button-1>", self._handle_press, add="+")
        self.tree.bind("<Double-Button-1>", self._handle_activate, add="+")
        self.tree.bind("<Return>", self._handle_activate, add="+")
        self.tree.bind("<KP_Enter>", self._handle_activate, add="+")

    # ------------------------------------------------------------------ #
    #  Стилизация (R-03) и колонки
    # ------------------------------------------------------------------ #
    def _apply_columns(self) -> None:
        headings = [(key, title, width, anchor, stretch)
                    for key, title, width, anchor, stretch in self._columns]
        self.tree["columns"] = [col[0] for col in headings]
        self.tree.heading("#0", text="")
        self.tree.column("#0", width=ui.px(30), minwidth=ui.px(30),
                         stretch=False, anchor="center")
        for key, title, width, anchor, stretch in headings:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=ui.px(width), minwidth=ui.px(width),
                             anchor=anchor, stretch=stretch)

    def set_columns(self, columns) -> None:
        """Смена набора колонок (переход контрагенты <-> документы)."""
        self._columns = tuple(columns)
        self._apply_columns()

    def _apply_style(self) -> None:
        """Цвета/шрифты Treeview из палитры текущей темы (R-03, V2)."""
        style = self._style
        name = self._style_name
        style.configure(
            name, background=ui.resolve(ui.COLORS["card"]),
            fieldbackground=ui.resolve(ui.COLORS["card"]),
            foreground=ui.resolve(ui.COLORS["text_secondary"]),
            rowheight=ui.px(ROW_HEIGHT), borderwidth=0, relief="flat",
            font=ui.font(12),
        )
        style.map(
            name,
            background=[("selected", ui.resolve(ui.COLORS["sidebar_active"]))],
            foreground=[("selected", ui.resolve(ui.COLORS["text"]))],
        )
        style.configure(
            f"{name}.Heading",
            background=ui.resolve(ui.COLORS["chrome"]),
            foreground=ui.resolve(ui.COLORS["text_secondary"]),
            font=ui.font(11, "bold"), borderwidth=0, relief="flat",
        )
        style.map(f"{name}.Heading", background=[("active", ui.resolve(ui.COLORS["chrome"]))])
        style.configure(
            f"Vertical.TScrollbar",
            background=ui.resolve(ui.COLORS["scroll"]),
            troughcolor=ui.resolve(ui.COLORS["card"]),
            borderwidth=0, arrowsize=ui.px(12),
        )
        style.map("Vertical.TScrollbar",
                  background=[("active", ui.resolve(ui.COLORS["scroll_hover"]))])
        # Теги сроков — те же цвета, что были у CTk-строк (2.18.0).
        self.tree.tag_configure("expired", foreground=ui.resolve(ui.COLORS["error_text"]))
        self.tree.tag_configure("soon", foreground=ui.resolve(ui.COLORS["warning_text"]))
        self.tree.tag_configure(
            "muted", foreground=ui.resolve(ui.COLORS["text_secondary"]))
        self.tree.tag_configure(
            "section", foreground=ui.resolve(ui.COLORS["text"]), font=ui.font(12, "bold"))
        self.tree.tag_configure(
            "showmore", foreground=ui.resolve(ui.COLORS["accent"]))

    # ------------------------------------------------------------------ #
    #  Точки типа (V4): изображения в колонке #0
    # ------------------------------------------------------------------ #
    def dot_photo(self, color: str) -> tk.PhotoImage:
        """Кэшированная точка заданного цвета; круглая за счёт прозрачных углов."""
        image = self._dots.get(color)
        if image is not None:
            return image
        size = max(6, ui.px(9))
        image = tk.PhotoImage(master=self, width=size, height=size)
        image.put(color, to=(1, 0, size - 1, size))
        image.put(color, to=(0, 1, size, size - 1))
        self._dots[color] = image
        return image

    def _dot_for(self, spec: RowSpec):
        if not spec.dot:
            return ""
        return self.dot_photo(spec.dot)

    # ------------------------------------------------------------------ #
    #  Операции модели (R-04)
    # ------------------------------------------------------------------ #
    def populate(self, rows: list[RowSpec]) -> None:
        """Полная популяция (смена фильтра, сортировки, drill-переход)."""
        self.tree.delete(*self.tree.get_children(""))
        self._rows = {}
        for row in rows:
            self.tree.insert("", "end", iid=row.iid, values=row.values,
                             tags=row.tags, image=self._dot_for(row))
            self._rows[row.iid] = row
        self._schedule_scrollbar_sync()

    def apply_diff(self, rows: list[RowSpec]) -> None:
        """Точечное обновление до модели: удаление, вставка, перемещение.

        Префикс, совпадающий с моделью, не трогается — iid и позиции
        неизменных строк сохраняются (критерий 5.1). Значения/теги/точка
        обновляются только у изменившихся строк (сверка с кэшем).
        Позиция прокрутки сохраняется (доля yview восстанавливается).
        """
        first, _last = self.tree.yview()
        plan = plan_diff(self.tree.get_children(""), rows)
        if plan.to_delete:
            for iid in plan.to_delete:
                self._rows.pop(iid, None)
            self.tree.delete(*plan.to_delete)
        for pos, iid, op in plan.ops:
            row = rows[pos]
            if op == "move":
                if self.tree.index(iid) != pos:
                    self.tree.move(iid, "", pos)
            else:
                self.tree.insert("", pos, iid=iid, values=row.values,
                                 tags=row.tags, image=self._dot_for(row))
            self._rows[iid] = row
        for row in plan.updates:
            if not self.tree.exists(row.iid):
                continue
            cached = self._rows.get(row.iid)
            if (cached is None or cached.values != row.values
                    or cached.tags != row.tags or cached.dot != row.dot):
                self.tree.item(row.iid, values=row.values, tags=row.tags,
                               image=self._dot_for(row))
                self._rows[row.iid] = row
        self.tree.yview_moveto(first)
        self._schedule_scrollbar_sync()

    def sync_model(self, rows: list[RowSpec]) -> None:
        """Гарантирует соответствие модели после дифа (контрольный инвариант)."""
        self._rows = {row.iid: row for row in rows}

    def row(self, iid: str) -> RowSpec | None:
        return self._rows.get(iid)

    def rows(self) -> dict[str, RowSpec]:
        return dict(self._rows)

    def order(self) -> tuple[str, ...]:
        return tuple(self.tree.get_children(""))

    # ------------------------------------------------------------------ #
    #  Выделение и активация
    # ------------------------------------------------------------------ #
    def set_selection(self, iid: str | None) -> None:
        """Выделить строку (None — снять); отсутствующий iid игнорируется.

        Эхо-события ``<<TreeviewSelect>>`` от программной смены не
        подавляются: обработчик контроллера идемпотентен (SET-семантика)
        и повторное уведомление не меняет состояние.
        """
        try:
            if iid is None:
                self.tree.selection_remove(*self.tree.selection())
            elif self.tree.exists(iid):
                children = self.tree.selection()
                if children != (iid,):
                    self.tree.selection_set(iid)
        except tk.TclError:
            pass  # виджет мог быть уже разрушен

    def selected_iid(self) -> str | None:
        selection = self.tree.selection()
        return selection[0] if selection else None

    def focus_tree(self) -> None:
        try:
            self.tree.focus_set()
        except Exception:
            pass

    def _mark_key_nav(self, _event=None) -> None:
        self._key_nav = True

    def _handle_press(self, event) -> None:
        """Нажатие на выделенную строку — снять выделение после класс-обработчика.

        В browse-режиме Tk не генерирует ``<<TreeviewSelect>>``, когда
        выделение не изменилось, поэтому повторный клик по выделенной
        строке обрабатывается здесь: отложенное снятие генерирует событие
        пустого выделения, которое и доходит до контроллера.
        """
        try:
            iid = self.tree.identify_row(event.y)
        except Exception:
            return
        if iid and iid == self.selected_iid():
            if self._toggle_after is not None:
                try:
                    self.after_cancel(self._toggle_after)
                except Exception:
                    pass
            self._toggle_after = self.after_idle(self._consume_click_toggle)

    def _consume_click_toggle(self) -> None:
        self._toggle_after = None
        if self.selected_iid() is not None:
            self.set_selection(None)
            # Явное действие пользователя — единственный путь снятия выделения:
            # контроллер отличает его от эха популяции/синхронизации (user=True).
            self._on_select(None, False, True)

    def _handle_select(self, _event=None) -> None:
        iid = self.selected_iid()
        key_nav = self._key_nav
        self._key_nav = False
        # Событие дерева — SET-семантика (user=False): клик по новой строке
        # и эхо программной синхронизации обрабатываются одинаково и
        # идемпотентно; снятие выделения идёт только через click-toggle.
        self._on_select(iid, key_nav, False)

    def _handle_activate(self, event) -> None:
        if self._toggle_after is not None:
            # Двойной клик по выделенной строке открывает запись, а не снимает
            # выделение: отложенный toggle отменяется.
            try:
                self.after_cancel(self._toggle_after)
            except Exception:
                pass
            self._toggle_after = None
        iid = self.selected_iid()
        if iid is None and hasattr(event, "y"):
            # Двойной клик: выделение могло не успеть смениться — берём строку
            # под курсором. Для Enter (event.y нет) строка без выделения
            # не активируется.
            iid = self.tree.identify_row(event.y)
        if iid:
            self._on_activate(iid)
        return "break"

    # ------------------------------------------------------------------ #
    #  Прокрутка: полоса видна, только когда есть что скроллить
    # ------------------------------------------------------------------ #
    def _on_yview(self, first, last) -> None:
        try:
            first, last = float(first), float(last)
        except (TypeError, ValueError):
            return
        if first <= 0.001 and last >= 0.999:
            self._vsb.grid_remove()
        else:
            self._vsb.grid()
        try:
            self._vsb.set(first, last)
        except Exception:
            pass

    def _schedule_scrollbar_sync(self) -> None:
        self.after_idle(lambda: self._on_yview(*self.tree.yview()))
