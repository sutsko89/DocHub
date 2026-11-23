"""Страница «Реестр» — четвёртый контроллер ContractApp.

Класс владеет кодом страницы «Реестр»: построение страницы, поиск по
названию и ИНН с фильтрацией по мере ввода, выделение контрагентов и
документов, открытие списка документов и документа в Word, рендер секций
«документы контрагента», «неопознанные» и «корзина», сверка реестра с папкой
документов и диалог разбора файлов вне реестра (п. 2.4).

Модель взаимодействия (этап 4 roadmap — TREE-1 + UX-REG, список на
ttk.Treeview):

* один клик по контрагенту/документу только выделяет строку — панель
  действий под списком показывает кнопки выделенной записи
  («Редактировать»/«Удалить», «Открыть Word»/«Удалить», «Восстановить»/
  «Удалить навсегда» для корзины); повторный клик снимает выделение;
  клавиша Delete удаляет выделенную запись (delete_selected);
* двойной клик по контрагенту проваливается в его документы: страница
  переключается на отдельный вид — заголовок «Документы — <имя>», список
  его документов и кнопка «← Назад» в шапке (и клавиша Escape)
  возвращают к списку контрагентов; двойной клик по документу открывает
  его в Word; Enter активирует выделенную строку, стрелки двигают
  выделение без снятия, Ctrl+F ставит фокус в поиск (U13);
* клик по строке «Показать ещё» (R-05b) показывает следующую порцию
  списка; поиск и сортировка всегда выполняются по всей базе;
* Delete внутри вида документов удаляет только выделенный документ —
  контрагент удаляется только из списка контрагентов.

Список реестра — RegistryTree (registry_tree.py): популяция целиком
только при смене фильтра/сортировки и drill-переходах, изменение состава
и правки записей — дифом (точечные insert/delete/item, строки без
изменений сохраняют iid, п. 5.1).

Состояние страницы остаётся на приложении с прежними именами
(``registry_search_var``, ``registry_list``, ``_registry_selected_id``,
``_registry_rows``, ``_registry_rendered``, ``_orphans`` …): его читают и
меняют другие страницы, сценарии сохранения и smoke-тест — контроллер
владеет кодом, а не данными страницы. Поля 2.19.0:
``_registry_drill_party_id`` (контрагент, в документы которого провалились),
``_registry_title_label`` / ``_registry_hint_label`` / ``_registry_back_btn``
(шапка вида), ``_docs_selected_id`` (выделенный документ),
``_registry_rows[id]`` / ``_docs_rows[id]`` хранят {rec, iid} строки.

Вызовы, которые подменяет smoke-тест, всегда идут через ``self.app.*``:
``_choice_dialog`` подменяется как instance-атрибут, действия панели
вызывают контроллеры напрямую (``app._counterparty_flow.edit_counterparty``,
``app._lifecycle_flow.delete_document``, ``export_document``, …), а показ
папки для неопознанных файлов — через ``app._show_in_folder``.
``ui.confirm`` / ``ui.alert`` — атрибуты общего модуля ui_macos, подмена
``m.ui.confirm`` видна контроллеру напрямую.
"""

from __future__ import annotations

import os
from typing import Any

import customtkinter as ctk
import tkinter as tk

import agreements
import applog
import reconcile
import registry_tree
import storage
import ui_macos as ui
from document_builder import F, digits_only
from registry_tree import (
    DISPLAY_SOFT_LIMIT,
    SHOW_MORE_IID,
    RowSpec,
    plan_display_count,
)
from services import safe_filename


#: Колонки списка контрагентов (U1): ключ, заголовок, ширина, якорь, тянуть.
PARTY_COLUMNS = (
    ("name", "Контрагент", 340, "w", True),
    ("kind", "Тип", 92, "w", False),
    ("inn", "ИНН", 128, "w", False),
    ("id", "ID", 104, "w", False),
    ("address", "Адрес", 360, "w", True),
)

#: Колонки вида «Документы контрагента».
DOC_COLUMNS = (
    ("doc", "Документ", 330, "w", True),
    ("date", "Дата", 116, "w", False),
    ("expiry", "Срок", 350, "w", True),
)

#: Сортировка списка контрагентов (U2): подпись меню -> значение settings.
SORT_VALUE_BY_LABEL = {
    "По добавлению": "added",
    "По алфавиту": "name",
    "По ИНН": "inn",
}
SORT_LABEL_BY_VALUE = {value: label for label, value in SORT_VALUE_BY_LABEL.items()}

#: Фильтр типов в виде документов контрагента (UX-REG).
DOCS_TYPE_VALUE_BY_LABEL = {
    "Все типы": "all",
    "Договоры": storage.KIND_CONTRACT,
    "Доп. соглашения": storage.KIND_ADDITIONAL,
    "Спецификации": storage.KIND_SPEC,
}
DOCS_TYPE_LABEL_BY_VALUE = {value: label for label, value in DOCS_TYPE_VALUE_BY_LABEL.items()}

#: Цвета точек типов (V4-пилот): точка в колонке #0 — PhotoImage.
DOT_IP = "#28C840"
DOT_OOO = "#007AFF"
DOT_CONTRACT = "#007AFF"
DOT_ADDITIONAL = "#FEBC2E"
DOT_SPEC = "#28C840"
DOT_UNKNOWN = "#96969C"

#: Префиксы iid по типу записи (party/doc ids не пересекаются).
IID_PARTY = "p:"
IID_DOC = "d:"
IID_TRASH_PARTY = "tp:"
IID_TRASH_DOC = "td:"
IID_UNKNOWN = "u:"
IID_SEC_UNKNOWN = "sec:unknown"
IID_SEC_TRASH = "sec:trash"

#: Кнопки панели действий по типу выделенной записи (V3).
ACTION_LABELS = {
    "party": ("Редактировать", "Удалить"),
    "doc": ("Открыть Word", "Удалить"),
    "trash_party": ("Восстановить", "Удалить навсегда"),
    "trash_doc": ("Восстановить", "Удалить навсегда"),
    "unknown": ("Показать в папке", "Удалить"),
}


def action_labels_for(kind: str | None) -> tuple[str, str] | None:
    """Подписи кнопок панели действий для типа записи (чистая функция)."""
    return ACTION_LABELS.get(kind or "")


class RegistryController:
    """Страница «Реестр»: построение, поиск, выделение, секции, сверка.

    Код страницы без её данных: все виджеты и переменные публикуются на
    приложении под прежними именами, поэтому сценарии сохранения, другие
    страницы и smoke-тест продолжают работать без правок.
    """

    def __init__(self, app: Any) -> None:
        self.app = app

    @property
    def is_drill_mode(self) -> bool:
        """True, пока открыт вид «Документы контрагента» (проваливание)."""
        return getattr(self.app, "_registry_drill_party_id", None) is not None

    # ------------------------------------------------------------------ #
    #  Построение страницы
    # ------------------------------------------------------------------ #
    def build_page(self):
        app = self.app
        page = ctk.CTkFrame(app.content, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        app.pages["registry"] = page

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * app.scale), pady=(int(18 * app.scale), int(6 * app.scale)))
        app._registry_header = header
        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")
        app._registry_title_label = ctk.CTkLabel(
            title_row, text="Реестр", font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        )
        app._registry_title_label.pack(side="left")
        # 2.19.0: выход из вида документов контрагента; показывается только
        # в режиме проваливания (_update_header_mode).
        app._registry_back_btn = ui.ghost_button(
            title_row, "← Назад", app._registry_flow.back_to_parties,
            width=int(110 * app.scale), height=int(28 * app.scale),
        )
        app._registry_back_btn.pack(side="left", padx=(int(14 * app.scale), 0))
        app._registry_back_btn.pack_forget()
        ui.ghost_button(
            title_row, "Обновить", app._registry_flow.refresh,
            width=int(110 * app.scale), height=int(28 * app.scale),
        ).pack(side="right")
        ui.ghost_button(
            title_row, "Добавить документы…", app._import_documents_dialog,
            width=int(180 * app.scale), height=int(28 * app.scale),
        ).pack(side="right", padx=(0, int(8 * app.scale)))
        app._registry_hint_label = ctk.CTkLabel(
            header,
            text="Клик — выделить запись (действия — на панели под списком, клавиша Delete). "
                 "Двойной клик или Enter — открыть. Стрелки листают список, Ctrl+F — поиск.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        )
        app._registry_hint_label.pack(anchor="w", pady=(int(2 * app.scale), 0))

        search_bar = ctk.CTkFrame(page, fg_color="transparent")
        app._registry_search_bar = search_bar
        search_bar.pack(fill="x", padx=int(24 * app.scale), pady=(0, int(8 * app.scale)))
        app.registry_search_var = tk.StringVar()
        # Поиск идёт и по названию, и по ИНН, с фильтрацией по мере ввода (п. 4.5).
        app.registry_search_var.trace_add("write", self.on_search_changed)
        search_entry = ui.entry(
            search_bar, width=int(280 * app.scale),
            textvariable=app.registry_search_var,
            placeholder="Поиск по названию или ИНН",  # U11 (UX-REG)
        )
        search_entry.pack(side="left")
        app._registry_search_entry = search_entry
        # U11 (UX-REG): Enter в поле поиска применяет фильтр сразу,
        # не дожидаясь паузы debounce (кнопка «Найти» удалена как
        # дублирующая живую фильтрацию).
        search_entry.entry.bind("<Return>", lambda _event: (self.apply_filter(), "break")[1])
        ui.ghost_button(
            search_bar, "Показать всех", self.reset_filter,
            width=int(130 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(8 * app.scale), 0))
        # U2 (UX-REG): сортировка с запоминанием в settings.json.
        app._registry_sort = str(
            (getattr(app, "_settings", None) or {}).get("registry_sort", "added"))
        if app._registry_sort not in SORT_LABEL_BY_VALUE:
            app._registry_sort = "added"
        app._registry_sort_var = tk.StringVar(
            value=SORT_LABEL_BY_VALUE[app._registry_sort])
        sort_menu = ui.option_menu(
            search_bar, list(SORT_VALUE_BY_LABEL),
            width=int(170 * app.scale), variable=app._registry_sort_var,
        )
        sort_menu.pack(side="left", padx=(int(10 * app.scale), 0))
        app._registry_sort_var.trace_add("write", self.on_sort_changed)

        # UX-REG: фильтр типов в виде документов контрагента; в режиме
        # контрагентов строка скрыта (видна строка поиска).
        filter_bar = ctk.CTkFrame(page, fg_color="transparent")
        app._registry_filter_bar = filter_bar
        app._docs_type_filter = "all"
        app._docs_type_var = tk.StringVar(value="Все типы")
        type_menu = ui.option_menu(
            filter_bar, list(DOCS_TYPE_VALUE_BY_LABEL),
            width=int(200 * app.scale), variable=app._docs_type_var,
        )
        type_menu.pack(side="left")
        app._docs_type_var.trace_add("write", self.on_docs_type_changed)

        # Баннеры сроков и файлов вне реестра — над списком.
        banners = ctk.CTkFrame(page, fg_color="transparent")
        app._registry_banners = banners
        banners.pack(fill="x", padx=int(24 * app.scale), pady=(0, int(6 * app.scale)))

        # Подпись списка: «Контрагенты (N)», «показаны K из M» (R-05b) —
        # в drill-виде «Документы — <имя> (N)» с подстрокой контрагента.
        caption_row = ctk.CTkFrame(page, fg_color="transparent")
        caption_row.pack(fill="x", padx=int(24 * app.scale), pady=(0, int(4 * app.scale)))
        app._registry_caption = ctk.CTkLabel(
            caption_row, text="", font=ui.font(14, "bold"),
            text_color=ui.COLORS["text"], anchor="w", height=20,
        )
        app._registry_caption.pack(anchor="w")
        app._registry_subtitle = ctk.CTkLabel(
            caption_row, text="", font=ui.font(11),
            text_color=ui.COLORS["text_tertiary"], anchor="w", height=16,
        )
        # app._registry_subtitle пакуется только в drill-виде (_update_caption).

        # Список на ttk.Treeview (TREE-1): свой скролл, панель действий
        # под списком (V3 — резерв места) вместо кнопок в каждой строке.
        app.registry_list = registry_tree.RegistryTree(
            page, on_select=self.on_tree_select, on_activate=self.on_tree_activate,
            columns=PARTY_COLUMNS,
        )
        app.registry_list.pack(fill="both", expand=True,
                               padx=int(24 * app.scale), pady=(0, int(8 * app.scale)))

        actions = ctk.CTkFrame(page, fg_color="transparent", height=int(36 * app.scale))
        app._registry_actions = actions
        actions.pack(fill="x", padx=int(24 * app.scale), pady=(0, int(10 * app.scale)))
        actions.pack_propagate(False)
        app._registry_actions_hint = ctk.CTkLabel(
            actions, text="Выберите запись в списке — действия появятся здесь.",
            font=ui.font(12), text_color=ui.COLORS["text_tertiary"], anchor="w",
        )
        app._registry_actions_hint.pack(side="left")
        app._registry_action_btn1 = ui.ghost_button(
            actions, "", self._run_action_1, width=int(140 * app.scale),
            height=int(28 * app.scale),
        )
        app._registry_action_btn2 = ui.ghost_button(
            actions, "", self._run_action_2, width=int(140 * app.scale),
            height=int(28 * app.scale),
        )
        app._registry_empty_trash_btn = ui.ghost_button(
            actions, "Очистить корзину",
            # _lifecycle_flow создаётся после построения страниц —
            # разрешаем команду в момент клика, не в build_page.
            lambda: app._lifecycle_flow.empty_trash(),
            width=int(150 * app.scale), height=int(28 * app.scale),
        )
        # Пустая корзина: кнопка очистки скрыта (grid_remove, как у 2.18.0).

        app._registry_selected_id = None
        app._registry_rows = {}
        # 2.19.0: проваливание в контрагента (двойной клик) и выделенный
        # документ внутри вида «Документы контрагента».
        app._registry_drill_party_id = None
        app._docs_selected_id = None
        app._docs_rows = {}
        # Выделенная запись секций (корзина/неопознанные) — действия панели.
        app._registry_selected_extra = None
        # R-05b: сколько строк показано (копится «Показать ещё»).
        app._registry_shown = DISPLAY_SOFT_LIMIT
        app._registry_signature = None
        # Ленивый рендер (п. 9 этапа 5): при старте список не рисуется —
        # реестр отрисуется при первом открытии страницы или при изменениях.
        app._registry_rendered = False

    # ------------------------------------------------------------------ #
    #  Список контрагентов, проваливание, секции реестра ----------------- #
    # ------------------------------------------------------------------ #
    def parties(self) -> list[dict]:
        """Активные контрагенты с учётом поиска по названию и ИНН (подстрока)."""
        query = self.app.registry_search_var.get().strip().lower()
        parties = self.app.store.active_counterparties()
        if not query:
            return parties
        digits = digits_only(query)

        def matches(rec: dict) -> bool:
            fields = rec.get("fields") or {}
            inn = str(fields.get(F.INN, "")).strip()
            if digits and digits in inn:
                return True
            name_parts = (
                storage.display_name(rec),
                str(fields.get(F.FULL_NAME, "")),
                str(fields.get(F.IP_NAME, "")),
                str(fields.get(F.SHORT_NAME, "")),
            )
            haystack = " ".join(name_parts).lower()
            return bool(query) and query in haystack

        return [rec for rec in parties if matches(rec)]

    def sorted_parties(self) -> list[dict]:
        """parties() в выбранном порядке сортировки (U2: добавление/алфавит/ИНН).

        Поиск всегда выполняется по всей базе — сортировка применяется
        к отфильтрованному множеству, лимит отображения (R-05b) — к
        отсортированному.
        """
        parties = self.parties()
        mode = getattr(self.app, "_registry_sort", "added")
        if mode == "name":
            parties.sort(key=lambda rec: storage.display_name(rec).casefold())
        elif mode == "inn":
            # Пустой ИНН — в конец списка, а не в начало.
            parties.sort(key=lambda rec: (
                str((rec.get("fields") or {}).get(F.INN, "")).strip() == "",
                str((rec.get("fields") or {}).get(F.INN, "")).strip(),
            ))
        return parties

    def on_page_open(self):
        """Открытие страницы «Реестр» из меню: список контрагентов с начала.

        Проваливание — режим просмотра, он не переживает уход со страницы:
        вернулись в реестр — снова виден список контрагентов (2.19.0).
        """
        if self.is_drill_mode:
            self.back_to_parties()
            return
        self.refresh()
        tree = getattr(self.app, "registry_list", None)
        if tree is not None:
            tree.focus_tree()

    def on_search_changed(self, *_args):
        """Фильтрует реестр по мере ввода с небольшой задержкой."""
        pending = getattr(self.app, "_registry_search_after", None)
        if pending:
            try:
                self.app.root.after_cancel(pending)
            except Exception:
                pass
        self.app._registry_search_after = self.app.root.after(200, self.refresh)

    def apply_filter(self):
        pending = getattr(self.app, "_registry_search_after", None)
        if pending:
            try:
                self.app.root.after_cancel(pending)
            except Exception:
                pass
            self.app._registry_search_after = None
        self.refresh()

    def reset_filter(self):
        self.app.registry_search_var.set("")
        self.apply_filter()

    def on_sort_changed(self, *_args):
        """Смена сортировки (U2): перестройка списка и settings.json."""
        app = self.app
        label = app._registry_sort_var.get()
        value = SORT_VALUE_BY_LABEL.get(label)
        if value is None or value == app._registry_sort:
            return
        app._registry_sort = value
        save = getattr(app, "save_setting", None)
        if save is not None:
            save("registry_sort", value)
        self.apply_filter()

    def on_docs_type_changed(self, *_args):
        """Смена фильтра типов в виде документов (UX-REG)."""
        app = self.app
        value = DOCS_TYPE_VALUE_BY_LABEL.get(app._docs_type_var.get())
        if value is None or value == app._docs_type_filter:
            return
        app._docs_type_filter = value
        self.apply_filter()

    def show_more(self):
        """Строка «Показать ещё» (R-05b): следующая порция списка.

        Фильтр, сортировка, выделение и позиция прокрутки сохраняются:
        популяция идёт дифом (append новых строк), подпись обновляется.
        """
        app = self.app
        app._registry_shown += DISPLAY_SOFT_LIMIT
        self.refresh()

    # ------------------------------------------------------------------ #
    #  Выделение и активация строк (Treeview)
    # ------------------------------------------------------------------ #
    def on_tree_select(self, iid: str | None, key_nav: bool = False,
                       user: bool = False) -> None:
        """Событие выделения RegistryTree (SET-семантика, идемпотентно).

        События дерева (клик по новой строке, стрелки, эхо программной
        синхронизации) приходят с user=False и только устанавливают
        состояние — повторное уведомление не меняет его. Снятие выделения
        идёт только явным действием пользователя (повторный клик,
        click-toggle виджета): user=True и iid=None.
        """
        app = self.app
        if iid is None:
            if not user:
                return  # эхо популяции/синхронизации — состояние не трогаем
            if self.is_drill_mode:
                app._docs_selected_id = None
            else:
                app._registry_selected_id = None
                app._registry_selected_extra = None
            self.refresh_selection()
            return
        row = app.registry_list.row(iid)
        if row is None:
            return
        data = row.data
        kind = data.get("kind")
        rec = data.get("rec")
        if kind == "party":
            app._registry_selected_extra = None
            app._docs_selected_id = None
            self._select_party_plain(rec)
        elif kind == "doc":
            app._registry_selected_extra = None
            self._select_document_plain(rec)
        else:
            # Секции корзины/неопознанных и служебная строка «Показать ещё»:
            # выделение запоминается для панели действий, клавиша Delete
            # на них не действует (там выбор restore/purge — через панель).
            app._registry_selected_extra = {"kind": kind, "rec": rec}
            app._registry_selected_id = None
            app._docs_selected_id = None
            self.refresh_selection()

    def on_tree_activate(self, iid: str) -> None:
        """Активация строки: двойной клик или Enter (U13)."""
        app = self.app
        if iid == SHOW_MORE_IID:
            self.show_more()
            return
        row = app.registry_list.row(iid)
        if row is None:
            return
        data = row.data
        if data.get("kind") == "party":
            self.open_party_documents(data["rec"])
        elif data.get("kind") == "doc":
            app._lifecycle_flow.export_document(data["rec"])

    def select_party(self, rec: dict):
        """Клик по контрагенту: выделить строку (повторный клик снимает).

        Выбор из реестра переносится на страницы документов (п. 4.2).
        Частичное обновление: строки не пересоздаются, обновляется
        выделение и панель действий. Проваливание остаётся прежним.
        """
        app = self.app
        if app._registry_selected_id == rec.get("id"):
            app._registry_selected_id = None
        else:
            app._registry_selected_id = rec.get("id")
            # Выбор из реестра переносится на страницы документов (п. 4.2).
            app._doc_context_flow.set_current_party(rec)
        app._docs_selected_id = None
        if not app._registry_rows:
            # Страница ещё не отрисована (ленивый рендер) — рисуем целиком.
            self.refresh()
            return
        self.refresh_selection()

    def _select_party_plain(self, rec: dict):
        """Стрелки (U13): выделить строку без toggle."""
        app = self.app
        app._registry_selected_id = rec.get("id")
        app._registry_selected_extra = None
        app._docs_selected_id = None
        app._doc_context_flow.set_current_party(rec)
        self.refresh_selection()

    def open_party_documents(self, rec: dict):
        """Двойной клик по контрагенту: провалиться в его документы (2.19.0).

        Страница переключается на вид «Документы — <имя>»: список документов
        контрагента, кнопка «← Назад» в шапке возвращает к контрагентам.
        """
        app = self.app
        app._registry_selected_id = rec.get("id")
        app._doc_context_flow.set_current_party(rec)
        app._docs_selected_id = None
        app._registry_drill_party_id = rec.get("id")
        self.refresh()

    def back_to_parties(self):
        """Кнопка «← Назад» (или Escape): вернуться к списку контрагентов."""
        app = self.app
        app._registry_drill_party_id = None
        app._docs_selected_id = None
        self.refresh()

    def select_document(self, rec: dict):
        """Клик по документу: выделить строку (повторный клик снимает)."""
        app = self.app
        if app._docs_selected_id == rec.get("id"):
            app._docs_selected_id = None
        else:
            app._docs_selected_id = rec.get("id")
        if not app._docs_rows:
            # Вид документов ещё не отрисован — рисуем целиком.
            self.refresh()
            return
        self.refresh_selection()

    def _select_document_plain(self, rec: dict):
        """Стрелки (U13): выделить документ без toggle."""
        app = self.app
        app._docs_selected_id = rec.get("id")
        app._registry_selected_extra = None
        self.refresh_selection()

    def delete_selected(self) -> bool:
        """Клавиша Delete: удалить выделенный документ или контрагента.

        Документ имеет приоритет (он «мельче»): если выделены и документ,
        и контрагент, удаляется документ. Возвращает True, если удаление
        инициировано (диалог подтверждения — в контроллере жизненного
        цикла).

        Внутри вида «Документы контрагента» (проваливание, 2.19.0) Delete
        удаляет только выделенный документ — контрагент нельзя удалить
        случайно из его списка документов. Выделенной записи корзины
        Delete не касается: у неё выбор «Восстановить»/«Удалить навсегда» —
        через панель действий, случайное удаление навсегда недопустимо.
        """
        app = self.app
        if getattr(app, "_registry_selected_extra", None):
            return False
        doc_id = getattr(app, "_docs_selected_id", None)
        if doc_id is not None:
            rec = app.store.get_document(doc_id)
            if rec is not None:
                app._lifecycle_flow.delete_document(rec)
                return True
        if self.is_drill_mode:
            return False
        party_id = getattr(app, "_registry_selected_id", None)
        if party_id is not None:
            rec = app.store.get_counterparty(party_id)
            if rec is not None:
                app._lifecycle_flow.delete_counterparty(rec)
                return True
        return False

    @staticmethod
    def party_row_text(rec: dict, selected: bool) -> str:
        """Текстовая форма строки контрагента (совместимость тестов).

        Список реестра показывает колонки (party_row_values); текстовая
        форма сохранена для сценариев и отчётов, где нужна одна строка.
        """
        kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
        inn = str(rec.get("fields", {}).get(F.INN, "")).strip() or "—"
        address = str(rec.get("fields", {}).get(F.LEGAL_ADDRESS, "")).strip()
        if len(address) > 48:
            address = address[:47].rstrip() + "…"
        text = f"{storage.display_name(rec)}   ·   {kind}   ·   ИНН {inn}   ·   ID {rec.get('id')}"
        if address:
            text += f"   ·   {address}"
        if selected:
            text = "●  " + text
        return text

    @staticmethod
    def party_row_values(rec: dict) -> tuple[str, str, str, str, str]:
        """Значения колонок строки контрагента (U1: name, тип, ИНН, ID, адрес)."""
        fields = rec.get("fields") or {}
        kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
        inn = str(fields.get(F.INN, "")).strip() or "—"
        address = str(fields.get(F.LEGAL_ADDRESS, "")).strip()
        if len(address) > 48:
            address = address[:47].rstrip() + "…"
        return (
            storage.display_name(rec), kind, inn, str(rec.get("id")), address,
        )

    @staticmethod
    def party_row_dot(rec: dict) -> str:
        """Цвет точки типа контрагента (V4-пилот)."""
        return DOT_IP if rec.get("type") == "ip" else DOT_OOO

    def doc_row_text(self, rec: dict) -> str:
        """Текстовая форма строки документа (совместимость тестов)."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        text = f"{kind_label} № {number}"
        if date_text:
            text += f" от {date_text}"
        text += agreements.row_suffix(rec, until=self._doc_until(rec))
        return text

    def _doc_until(self, rec: dict) -> "str | None":
        """Срок строки документа из лёгкой карты хранилища (LAZY-1).

        Для записей с телом в памяти agreements сам берёт rec["extra"];
        ленивые записи берут срок из карты, без чтения базы на рендер.
        """
        if rec.get("kind") != storage.KIND_ADDITIONAL:
            return None
        return self.app.store.expiry_until(rec.get("id"))

    @staticmethod
    def doc_row_values(rec: dict, until: "str | None" = None) -> tuple[str, str, str]:
        """Значения колонок строки документа: документ, дата, срок (U1).

        ``until`` — срок из лёгкой карты хранилища (LAZY-1); без него срок
        берётся из тела в памяти записи (совместимость со старыми вызовами).
        """
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        suffix = agreements.row_suffix(rec, until=until)
        if suffix.startswith("   ·   "):
            suffix = suffix[len("   ·   "):]
        return (f"{kind_label} № {number}", date_text or "—", suffix)

    @staticmethod
    def doc_row_tags(rec: dict, until: "str | None" = None) -> tuple[str, ...]:
        """Теги строки документа по состоянию срока (2.18.0): красный/амберный."""
        if rec.get("kind") == storage.KIND_ADDITIONAL:
            state = agreements.collect_state(rec, until=until)
            if state == "expired":
                return ("expired",)
            if state == "soon":
                return ("soon",)
        return ()

    def doc_row_color(self, rec: dict) -> tuple:
        """Цвет текста строки документа по состоянию срока (2.18.0).

        Просроченное доп. соглашение — красный, истекающее в течение
        месяца — амберный, остальное — обычный.
        """
        tags = self.doc_row_tags(rec, until=self._doc_until(rec))
        if "expired" in tags:
            return ui.COLORS["error_text"]
        if "soon" in tags:
            return ui.COLORS["warning_text"]
        return ui.COLORS["text"]

    @staticmethod
    def party_subtitle(rec: dict) -> str:
        """Подстрока контрагента в виде документов: тип, ИНН, ID (2.19.0)."""
        kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
        inn = str(rec.get("fields", {}).get(F.INN, "")).strip() or "—"
        return f"{kind}   ·   ИНН {inn}   ·   ID {rec.get('id')}"

    # ------------------------------------------------------------------ #
    #  Модель строк реестра (TREE-1: R-03/R-04, R-05b)
    # ------------------------------------------------------------------ #
    def _view_signature(self) -> tuple:
        """Что заставляет пересобрать список целиком: фильтр, сортировка,
        фильтр типов, drill-переход. Изменение состава и правки записей в
        сигнатуру не входят — они идут дифом (R-04)."""
        app = self.app
        return (
            getattr(app, "_registry_drill_party_id", None),
            app.registry_search_var.get().strip().lower(),
            getattr(app, "_registry_sort", "added"),
            getattr(app, "_docs_type_filter", "all"),
        )

    def _party_rows(self) -> tuple[list[RowSpec], str]:
        """Модель списка контрагентов: порция (R-05b) + служебная строка."""
        app = self.app
        parties = self.sorted_parties()
        total = len(parties)
        shown, has_more = plan_display_count(total, app._registry_shown)
        rows = [
            RowSpec(
                iid=f"{IID_PARTY}{rec.get('id')}",
                values=self.party_row_values(rec),
                tags=(),
                dot=self.party_row_dot(rec),
                data={"kind": "party", "rec": rec},
            )
            for rec in parties[:shown]
        ]
        if has_more:
            hidden = total - shown
            portion = min(DISPLAY_SOFT_LIMIT, hidden)
            rows.append(RowSpec(
                iid=SHOW_MORE_IID,
                values=(f"Показать ещё ({portion} из {hidden})", "", "", "", ""),
                tags=("showmore",),
                data={"kind": "showmore"},
            ))
        query = app.registry_search_var.get().strip()
        caption = (
            f"Найдено по запросу «{query}»: {total}"
            if query else f"Контрагенты ({total})"
        )
        if has_more:
            caption += f" · показаны {shown} из {total}"
        elif shown > DISPLAY_SOFT_LIMIT and total > DISPLAY_SOFT_LIMIT:
            caption += f" · показаны все {total}"
        return rows, caption

    def _doc_rows(self, party: dict) -> tuple[list[RowSpec], str, str]:
        """Модель вида «Документы контрагента» с фильтром типов (UX-REG)."""
        app = self.app
        docs = app.store.documents_for(party["id"])
        type_filter = getattr(app, "_docs_type_filter", "all")
        if type_filter != "all":
            docs = [rec for rec in docs if rec.get("kind") == type_filter]
        rows = [
            RowSpec(
                iid=f"{IID_DOC}{rec.get('id')}",
                values=self.doc_row_values(rec, until=self._doc_until(rec)),
                tags=self.doc_row_tags(rec, until=self._doc_until(rec)),
                dot=self._doc_dot(rec),
                data={"kind": "doc", "rec": rec},
            )
            for rec in docs
        ]
        caption = f"Документы — {storage.display_name(party)} ({len(docs)})"
        if type_filter != "all":
            caption += f" · {DOCS_TYPE_LABEL_BY_VALUE.get(type_filter, type_filter).lower()}"
        return rows, caption, self.party_subtitle(party)

    @staticmethod
    def _doc_dot(rec: dict) -> str | None:
        """Цвет точки типа документа (V4-пилот); None — точки нет."""
        kind = rec.get("kind")
        return {
            storage.KIND_CONTRACT: DOT_CONTRACT,
            storage.KIND_ADDITIONAL: DOT_ADDITIONAL,
            storage.KIND_SPEC: DOT_SPEC,
        }.get(kind)

    def _unknown_rows(self) -> list[RowSpec]:
        """Секция «Неопознанные документы» в дереве (модель вместо CTk-строк)."""
        app = self.app
        docs = [
            rec for rec in app.store.documents
            if rec.get("kind") == storage.KIND_UNKNOWN
            and not storage.is_deleted(rec)
        ]
        if not docs:
            return []
        rows = [RowSpec(
            iid=IID_SEC_UNKNOWN,
            values=(f"Неопознанные документы ({len(docs)}) — файлы скопированы "
                    "в папку «unknown» внутри папки документов", "", "", "", ""),
            tags=("section",),
            data={"kind": "section"},
        )]
        for rec in docs:
            source = str(
                (app.store.document_body(rec.get("id")).get("extra") or {}).get("source_file") or ""
            )
            name = os.path.basename(source) if source else "Документ"
            rows.append(RowSpec(
                iid=f"{IID_UNKNOWN}{rec.get('id')}",
                values=(name, "файл", "—", "—", "папка «unknown»"),
                tags=("muted",),
                dot=DOT_UNKNOWN,
                data={"kind": "unknown", "rec": rec},
            ))
        return rows

    def _trash_rows(self) -> list[RowSpec]:
        """Секция «Корзина» в дереве (модель вместо CTk-строк)."""
        app = self.app
        trashed_parties = app.store.trashed_counterparties()
        trashed_docs = [
            rec for rec in app.store.trashed_documents()
            if not rec.get("deleted_with_party")
        ]
        if not trashed_parties and not trashed_docs:
            return []
        total = len(trashed_parties) + len(trashed_docs)
        rows = [RowSpec(
            iid=IID_SEC_TRASH,
            values=(f"Корзина ({total}) — записи удалены из реестра, файлы на "
                    "диске остались; восстановление вернёт их в работу",
                    "", "", "", ""),
            tags=("section",),
            data={"kind": "section"},
        )]
        for party in trashed_parties:
            docs = app.store.documents_for(party.get("id"), include_deleted=True)
            rows.append(RowSpec(
                iid=f"{IID_TRASH_PARTY}{party.get('id')}",
                values=(f"Контрагент «{storage.display_name(party)}»", "корзина",
                        "—", str(party.get("id")), f"документов: {len(docs)}"),
                tags=("muted",),
                data={"kind": "trash_party", "rec": party},
            ))
        for rec in trashed_docs:
            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            party_name = app.store.party_name(rec.get("party_id"))
            doc, date, expiry = self.doc_row_values(rec, until=self._doc_until(rec))
            rows.append(RowSpec(
                iid=f"{IID_TRASH_DOC}{rec.get('id')}",
                values=(f"{doc} · {party_name}" if party_name else doc,
                        date, expiry or "—", "", "корзина"),
                tags=("muted",),
                data={"kind": "trash_doc", "rec": rec},
            ))
        return rows

    def build_model(self) -> tuple[list[RowSpec], str, str]:
        """Модель строк текущего вида + подпись + подстрока.

        Контракт R-05b: поиск и сортировка — по всей базе; порция режет
        только отображение. Секции «Неопознанные» и «Корзина» добавляются
        в модель после порции контрагентов.
        """
        app = self.app
        drill_id = getattr(app, "_registry_drill_party_id", None)
        if drill_id is not None:
            party = app.store.get_counterparty(drill_id)
            if party is not None and not storage.is_deleted(party):
                rows, caption, subtitle = self._doc_rows(party)
                return rows, caption, subtitle
            # Контрагент удалён (мягкое удаление) — тихий возврат к списку.
            app._registry_drill_party_id = None
            app._docs_selected_id = None
        rows, caption = self._party_rows()
        rows = rows + self._unknown_rows() + self._trash_rows()
        return rows, caption, ""

    def refresh(self):
        """Обновляет реестр (поиск, изменения состава, проваливание).

        Полная популяция — при смене фильтра/сортировки/фильтра типов и
        drill-переходах (сигнатура вида) и при неотрисованном списке;
        остальное — диф (R-04). Длительные перерисовки (более 500 мс)
        попадают в журнал (п. 10 этапа 5).
        """
        with applog.timed("Перерисовка реестра", 500, "ui"):
            self.refresh_smart()
        self.app._registry_rendered = True

    def refresh_smart(self):
        """Выбирает между полной популяцией и дифом (R-04)."""
        app = self.app
        tree = getattr(app, "registry_list", None)
        if tree is None:
            return  # страница ещё не построена (юнит-тесты без Tk)
        if not app._registry_rendered or not app._registry_rows and not self.is_drill_mode:
            self.refresh_full()
            return
        # Контрагент drill-вида мог быть удалён — возврат меняет сигнатуру.
        drill_id = getattr(app, "_registry_drill_party_id", None)
        if drill_id is not None:
            party = app.store.get_counterparty(drill_id)
            if party is None or storage.is_deleted(party):
                self.refresh_full()
                return
        if self._view_signature() != getattr(app, "_registry_signature", None):
            self.refresh_full()
            return
        rows, caption, subtitle = self.build_model()
        tree.apply_diff(rows)
        self._after_model(rows, caption, subtitle)

    def refresh_full(self):
        """Полная популяция списка (смена фильтра, сортировки, drill)."""
        app = self.app
        tree = app.registry_list
        app._registry_signature = self._view_signature()
        app._registry_shown = DISPLAY_SOFT_LIMIT
        columns = DOC_COLUMNS if self.is_drill_mode else PARTY_COLUMNS
        tree.set_columns(columns)
        self._update_header_mode()
        rows, caption, subtitle = self.build_model()
        tree.populate(rows)
        self._after_model(rows, caption, subtitle)

    def _after_model(self, rows: list[RowSpec], caption: str, subtitle: str) -> None:
        """После популяции/дифа: карты состояния, подпись, баннеры, выделение.

        Контрольный assert (R-04, debug-сборка): порядок строк дерева
        совпадает с моделью — рассинхрон карты id->iid невозможен.
        """
        app = self.app
        tree = app.registry_list
        app._registry_rows = {
            row.data["rec"].get("id"): {"rec": row.data["rec"], "iid": row.iid}
            for row in rows if row.data.get("kind") == "party"
        }
        app._docs_rows = {
            row.data["rec"].get("id"): {"rec": row.data["rec"], "iid": row.iid}
            for row in rows if row.data.get("kind") == "doc"
        }
        self._update_caption(caption, subtitle)
        self._update_banners()
        self.refresh_selection()
        if __debug__:
            assert tree.order() == tuple(row.iid for row in rows), \
                "рассинхрон дерева и модели строк реестра"

    def _update_caption(self, caption: str, subtitle: str) -> None:
        """Подпись списка и подстрока drill-вида."""
        app = self.app
        label = getattr(app, "_registry_caption", None)
        if label is None:
            return
        try:
            label.configure(text=caption)
            subtitle_label = app._registry_subtitle
            if subtitle:
                subtitle_label.configure(text=subtitle)
                subtitle_label.pack(anchor="w")
            else:
                subtitle_label.pack_forget()
        except Exception:
            pass  # страница могла быть разрушена

    def _update_banners(self):
        """Баннеры сроков и файлов вне реестра над списком (2.18.0 / 2.4).

        Перестраиваются при каждом обновлении: это 0–2 компактных фрейма,
        дорогое содержимое (список) живёт в Treeview.
        """
        app = self.app
        banners = getattr(app, "_registry_banners", None)
        if banners is None:
            return
        for child in banners.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        self._render_expiry_banner(banners)
        self._render_orphans_banner(banners)

    def _render_expiry_banner(self, parent):
        """Баннер о сроках доп. соглашений над списком (2.18.0)."""
        app = self.app
        try:
            buckets = agreements.collect(app.store)
        except Exception:
            return
        expired, soon = buckets["expired"], buckets["soon"]
        if not expired and not soon:
            return
        lines: list[str] = []
        if expired:
            lines.append(f"Просрочено доп. соглашений: {len(expired)} — "
                         "строки отмечены красным.")
        if soon:
            lines.append(f"Истекает в течение месяца: {len(soon)} — "
                         "строки отмечены жёлтым.")
        banner = ctk.CTkFrame(parent, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8)
        banner.pack(fill="x", pady=(0, int(6 * app.scale)))
        ctk.CTkLabel(
            banner, text="\n".join(lines),
            font=ui.font(12), text_color=ui.COLORS["text"], anchor="w", justify="left",
            wraplength=int(560 * app.scale),
        ).pack(anchor="w", padx=int(10 * app.scale), pady=int(8 * app.scale))

    def _render_orphans_banner(self, parent):
        """Баннер файлов вне реестра над списком (п. 2.4)."""
        app = self.app
        if not app._orphans:
            return
        banner = ctk.CTkFrame(parent, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8)
        banner.pack(fill="x", pady=(0, int(6 * app.scale)))
        ctk.CTkLabel(
            banner,
            text=f"В папке документов найдено файлов вне реестра: {len(app._orphans)}.\n"
                 "Их можно восстановить в реестр или убрать в служебную папку _trash.",
            font=ui.font(12), text_color=ui.COLORS["text"], anchor="w", justify="left",
            wraplength=int(500 * app.scale),
        ).pack(anchor="w", padx=int(10 * app.scale), pady=int(8 * app.scale))
        ui.ghost_button(
            banner, "Разобрать", self.open_reconcile_dialog,
            width=int(110 * app.scale), height=int(26 * app.scale),
        ).pack(side="right", padx=(0, int(10 * app.scale)))

    # ------------------------------------------------------------------ #
    #  Выделение и панель действий (V3)
    # ------------------------------------------------------------------ #
    def refresh_selection(self):
        """Частичное обновление при смене выбора: подсветка строки и панель.

        Строки не пересоздаются (5.1): у Treeview меняется только выделение,
        панель действий перестраивается под тип выделенной записи.
        """
        app = self.app
        tree = getattr(app, "registry_list", None)
        if tree is None:
            return
        iid = self._selected_iid()
        tree.set_selection(iid)
        self.refresh_action_panel()

    def _selected_iid(self) -> str | None:
        """iid видимой строки по состоянию выделения (или None)."""
        app = self.app
        tree = app.registry_list
        if self.is_drill_mode:
            doc_id = getattr(app, "_docs_selected_id", None)
            if doc_id is None:
                return None
            iid = f"{IID_DOC}{doc_id}"
        else:
            extra = getattr(app, "_registry_selected_extra", None)
            if extra is not None:
                prefix = {
                    "trash_party": IID_TRASH_PARTY,
                    "trash_doc": IID_TRASH_DOC,
                    "unknown": IID_UNKNOWN,
                }.get(extra["kind"])
                iid = f"{prefix}{extra['rec'].get('id')}" if prefix else None
            else:
                party_id = getattr(app, "_registry_selected_id", None)
                iid = f"{IID_PARTY}{party_id}" if party_id is not None else None
        if iid is None or tree.row(iid) is None:
            return None  # выделенная запись не видна (фильтр/порция)
        return iid

    def _selected_entry(self) -> tuple[str, dict | None]:
        """Тип и запись выделенной строки для панели действий."""
        app = self.app
        iid = self._selected_iid()
        if iid is None:
            return "none", None
        row = app.registry_list.row(iid)
        if row is None:
            return "none", None
        return row.data.get("kind", "none"), row.data.get("rec")

    def refresh_action_panel(self):
        """Кнопки панели под списком по типу выделенной записи (V3).

        Панель имеет постоянную высоту — кнопки не прыгают при выделении
        и снятии выделения (резерв места, V3 UX-аудита).
        """
        app = self.app
        actions = getattr(app, "_registry_actions", None)
        if actions is None:
            return  # юнит-тесты без Tk
        kind, rec = self._selected_entry()
        labels = action_labels_for(kind)
        hint = app._registry_actions_hint
        btn1, btn2 = app._registry_action_btn1, app._registry_action_btn2
        if labels is None or rec is None:
            hint.pack(side="left")
            btn1.pack_forget()
            btn2.pack_forget()
        else:
            try:
                hint.pack_forget()
                btn1.configure(text=labels[0],
                               command=lambda k=kind, r=rec: self.run_action(k, r))
                btn1.pack(side="left", padx=(0, int(8 * app.scale)))
                btn2.configure(text=labels[1],
                               command=lambda k=kind, r=rec: self.run_action(k, r, second=True))
                btn2.pack(side="left")
            except Exception:
                pass  # страница могла быть разрушена
        trash_btn = app._registry_empty_trash_btn
        try:
            has_trash = bool(app.store.trashed_counterparties()) or \
                bool(app.store.trashed_documents())
            if has_trash:
                trash_btn.pack(side="right")
            else:
                trash_btn.pack_forget()
        except Exception:
            pass

    def _run_action_1(self):
        kind, rec = self._selected_entry()
        if rec is not None:
            self.run_action(kind, rec)

    def _run_action_2(self):
        kind, rec = self._selected_entry()
        if rec is not None:
            self.run_action(kind, rec, second=True)

    def run_action(self, kind: str, rec: dict, second: bool = False) -> None:
        """Действие панели: маршрутизация к контроллерам (как у строк 2.18.0)."""
        app = self.app
        if kind == "party":
            if second:
                app._lifecycle_flow.delete_counterparty(rec)
            else:
                app._counterparty_flow.edit_counterparty(rec)
        elif kind == "doc":
            if second:
                app._lifecycle_flow.delete_document(rec)
            else:
                app._lifecycle_flow.export_document(rec)
        elif kind in ("trash_party", "trash_doc"):
            if second:
                if kind == "trash_party":
                    app._lifecycle_flow.purge_counterparty(rec)
                else:
                    app._lifecycle_flow.purge_document(rec)
            else:
                if kind == "trash_party":
                    app._lifecycle_flow.restore_counterparty(rec)
                else:
                    app._lifecycle_flow.restore_document(rec)
        elif kind == "unknown":
            if second:
                app._lifecycle_flow.delete_document(rec)
            else:
                source = str(
                    (app.store.document_body(rec.get("id")).get("extra") or {}).get("source_file") or ""
                )
                app._show_in_folder(source)

    def _update_header_mode(self):
        """Шапка страницы по режиму: «Реестр» или «Документы — <имя>» (2.19.0)."""
        app = self.app
        title = getattr(app, "_registry_title_label", None)
        if title is None:
            return  # страница ещё не построена (юнит-тесты без Tk)
        back_btn = getattr(app, "_registry_back_btn", None)
        hint = getattr(app, "_registry_hint_label", None)
        search_bar = getattr(app, "_registry_search_bar", None)
        filter_bar = getattr(app, "_registry_filter_bar", None)
        drill_id = getattr(app, "_registry_drill_party_id", None)
        if drill_id is not None:
            party = app.store.get_counterparty(drill_id)
            name = storage.display_name(party) if party is not None else ""
            title.configure(text=f"Документы — {name}" if name else "Документы")
            if hint is not None:
                hint.configure(
                    text="Клик — выделить документ («Открыть Word»/«Удалить» на панели, "
                         "клавиша Delete). Двойной клик или Enter — открыть в Word.",
                )
            if back_btn is not None:
                back_btn.pack(side="left", padx=(int(14 * app.scale), 0))
            if search_bar is not None:
                search_bar.pack_forget()  # поиск фильтрует контрагентов, не документы
            if filter_bar is not None:
                filter_bar.pack(
                    fill="x", padx=int(24 * app.scale), pady=(0, int(8 * app.scale)),
                    after=getattr(app, "_registry_header", None),
                )
        else:
            title.configure(text="Реестр")
            if hint is not None:
                hint.configure(
                    text="Клик — выделить запись (действия — на панели под списком, "
                         "клавиша Delete). Двойной клик или Enter — открыть. "
                         "Стрелки листают список, Ctrl+F — поиск.",
                )
            if back_btn is not None:
                back_btn.pack_forget()
            if filter_bar is not None:
                filter_bar.pack_forget()
            if search_bar is not None:
                # after=header: после pack_forget/pack строка поиска должна
                # вернуться на своё место между шапкой и списком, а не в
                # конец упаковки (registry_list — canvas window, ссылка
                # before/after по нему невозможна) (2.19.0).
                search_bar.pack(
                    fill="x", padx=int(24 * app.scale), pady=(0, int(8 * app.scale)),
                    after=getattr(app, "_registry_header", None),
                )

    # ------------------------------------------------------------------ #
    #  Сверка реестра с папкой документов
    # ------------------------------------------------------------------ #
    def scan_consistency(self, orphans: "list[str] | None" = None) -> None:
        """Сверка «реестр ↔ папка документов»: применяет результат (START-1).

        ``orphans`` — готовый список файлов вне реестра из рабочего потока
        (START-1); без него сверка выполняется на месте синхронно
        (ручной вызов и тесты) через bulk-карту source_file (LAZY-1).
        """
        app = self.app
        if orphans is None:
            try:
                base = app._settings_flow.documents_dir(interactive=False)
                orphans = reconcile.find_orphans(
                    base, app.store, source_files=app.store.bulk_source_files(),
                )
            except Exception as exc:  # noqa: BLE001 — фоновая задача не должна падать
                app._log.warning("Сверка реестра с папкой не удалась: %s", type(exc).__name__)
                return
        else:
            app._log.info("Сверка реестра с папкой выполнена в фоне")
        app._orphans = orphans
        if orphans:
            app._log.info("Файлов вне реестра: %d", len(orphans))
        if "registry" in app.pages:
            self.refresh()

    def open_reconcile_dialog(self) -> None:
        """Диалог разбора файлов вне реестра: восстановить или убрать в _trash."""
        app = self.app
        base = app._settings_flow.documents_dir()
        orphans = reconcile.find_orphans(
            base, app.store, source_files=app.store.bulk_source_files(),
        )
        app._orphans = orphans
        if not orphans:
            ui.alert(
                app.root, "Файлы вне реестра",
                "Все файлы документов учтены в реестре.", "info",
            )
            self.refresh()
            return

        preview = "\n".join(f"• {os.path.basename(path)}" for path in orphans[:8])
        if len(orphans) > 8:
            preview += f"\n• …и ещё {len(orphans) - 8}"
        choice = app._choice_dialog(
            "Файлы вне реестра",
            f"В папке документов найдено файлов вне реестра: {len(orphans)}.\n\n"
            f"{preview}\n\n"
            "«Восстановить» — создать записи реестра (тип документа и контрагент "
            "определяются по имени файла и папки).\n"
            "«В корзину» — перенести файлы в служебную папку _trash внутри папки "
            "документов.",
            [
                ("skip", "Оставить", False),
                ("trash", "В корзину", False),
                ("restore", "Восстановить в реестр", True),
            ],
        )
        if choice == "restore":
            added = reconcile.restore_orphans(app.store, base, orphans, safe_filename)
            failed = len(orphans) - added
            if app.store.last_error:
                # D1: сбой финальной записи — честные счётчики «восстановлено /
                # не восстановлено», а не молчаливый успех.
                ui.alert(
                    app.root, "Восстановление не удалось",
                    f"{app.store.last_error}\n\n"
                    f"Восстановлено файлов: {added}, не восстановлено: {failed}.",
                    "error",
                )
            else:
                message = f"Добавлено записей в реестр: {added}."
                if failed:
                    message += (
                        f"\nНе восстановлено файлов: {failed} "
                        "(подробности в журнале)."
                    )
                message += "\n\nПроверьте типы документов и контрагентов в реестре."
                ui.alert(
                    app.root, "Реестр восстановлен",
                    message, "success",
                )
            app._orphans = reconcile.find_orphans(
                base, app.store, source_files=app.store.bulk_source_files(),
            )
            app._doc_context_flow.refresh_party_combo_values()
            self.refresh()
            app._update_status()
        elif choice == "trash":
            moved = reconcile.move_orphans_to_trash(base, orphans)
            ui.alert(
                app.root, "Файлы перемещены",
                f"Перемещено файлов: {moved}.\n\n"
                f"Папка: {os.path.join(base, reconcile.TRASH_DIR_NAME)}",
                "info",
            )
            app._orphans = reconcile.find_orphans(
                base, app.store, source_files=app.store.bulk_source_files(),
            )
            self.refresh()
