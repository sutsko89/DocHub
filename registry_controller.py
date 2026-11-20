"""Страница «Реестр» — четвёртый контроллер ContractApp.

Класс владеет кодом страницы «Реестр»: построение страницы, поиск по
названию и ИНН с фильтрацией по мере ввода, выделение контрагентов и
документов, открытие списка документов и документа в Word, рендер секций
«документы контрагента», «неопознанные» и «корзина», сверка реестра с папкой
документов и диалог разбора файлов вне реестра (п. 2.4).

Модель взаимодействия (2.19.0, отзыв пользователя — «проваливание»):

* один клик по контрагенту/документу только выделяет строку — у
  выделенной строки появляются кнопки действий («Редактировать»,
  «Удалить» / «Открыть Word», «Удалить»), повторный клик снимает
  выделение; клавиша Delete удаляет выделенную запись (delete_selected);
* двойной клик по контрагенту проваливается в его документы: страница
  переключается на отдельный вид — заголовок «Документы — <имя>»,
  список его документов и кнопка «← Назад» в шапке (и клавиша Escape)
  возвращают к списку контрагентов. CTk-кнопка не отдаёт
  <Double-Button-1> (текст рисуется на внутреннем виджете поверх канвы),
  поэтому двойной клик распознаётся по времени между двумя command-вызовами
  (DOUBLE_CLICK_MS);
* двойной клик по документу открывает его в Word; Delete внутри вида
  документов удаляет только выделенный документ — контрагент удаляется
  только из списка контрагентов.

Состояние страницы остаётся на приложении с прежними именами
(``registry_search_var``, ``registry_list``, ``_registry_selected_id``,
``_registry_rows``, ``_registry_rendered``, ``_orphans`` …): его читают и
меняют другие страницы, сценарии сохранения и smoke-тест — контроллер
владеет кодом, а не данными страницы. Поля 2.19.0:
``_registry_drill_party_id`` (контрагент, в документы которого провалились),
``_registry_title_label`` / ``_registry_hint_label`` / ``_registry_back_btn``
(шапка вида), ``_docs_selected_id`` (выделенный документ),
``_registry_rows[id]`` хранит и кнопки действий строки.

Вызовы, которые подменяет smoke-тест, всегда идут через ``self.app.*``:
``_choice_dialog`` подменяется как instance-атрибут, кнопки действий
вызывают контроллеры напрямую (``app._counterparty_flow.edit_counterparty``,
``app._lifecycle_flow.delete_document``, ``export_document``, …), а показ
папки для неопознанных файлов — через ``app._show_in_folder``.
``ui.confirm`` / ``ui.alert`` — атрибуты общего модуля ui_macos, подмена
``m.ui.confirm`` видна контроллеру напрямую.
"""

from __future__ import annotations

import os
import time
from typing import Any

import customtkinter as ctk
import tkinter as tk

import agreements
import applog
import reconcile
import storage
import ui_macos as ui
from document_builder import F, digits_only
from services import safe_filename


class RegistryController:
    """Страница «Реестр»: построение, поиск, выделение, секции, сверка.

    Код страницы без её данных: все виджеты и переменные публикуются на
    приложении под прежними именами, поэтому сценарии сохранения, другие
    страницы и smoke-тест продолжают работать без правок.
    """

    # Окно двойного клика: два клика по одной строке в пределах этого
    # интервала считаются «открыть» (см. докстринг модуля).
    DOUBLE_CLICK_MS = 400

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
            text="Клик — выделить контрагента (Редактировать/Удалить, клавиша Delete). "
                 "Двойной клик — открыть его документы.",
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
        )
        search_entry.pack(side="left")
        ui.ghost_button(
            search_bar, "Найти", app._registry_flow.apply_filter,
            width=int(92 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(10 * app.scale), 0))
        ui.ghost_button(
            search_bar, "Показать всех", self.reset_filter,
            width=int(130 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(8 * app.scale), 0))

        app.registry_list = ctk.CTkScrollableFrame(
            page, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        app.registry_list.pack(fill="both", expand=True, padx=int(24 * app.scale), pady=(0, int(10 * app.scale)))
        app.registry_list.grid_columnconfigure(0, weight=1)
        app.scrollers.append(app.registry_list)

        app._registry_selected_id = None
        app._registry_rows = {}
        # 2.19.0: проваливание в контрагента (двойной клик) и выделенный
        # документ внутри вида «Документы контрагента».
        app._registry_drill_party_id = None
        app._docs_selected_id = None
        app._docs_rows = {}
        # Времена последних кликов для распознавания двойного клика.
        app._registry_last_click_id = None
        app._registry_last_click_ts = 0.0
        app._docs_last_click_id = None
        app._docs_last_click_ts = 0.0
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

    def on_page_open(self):
        """Открытие страницы «Реестр» из меню: список контрагентов с начала.

        Проваливание — режим просмотра, он не переживает уход со страницы:
        вернулись в реестр — снова виден список контрагентов (2.19.0).
        """
        if self.is_drill_mode:
            self.back_to_parties()
            return
        self.refresh()

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

    # -- Клик и двойной клик --------------------------------------------- #
    def _is_double_click(self, state_key: str, item_id) -> bool:
        """True, если предыдущий клик был по той же записи недавно.

        Состояние хранится на приложении (``app._registry_last_click_*``
        / ``app._docs_last_click_*``), как и прочее состояние страницы.
        """
        app = self.app
        now = time.monotonic()
        last_id = getattr(app, f"_{state_key}_last_click_id", None)
        last_ts = getattr(app, f"_{state_key}_last_click_ts", 0.0)
        setattr(app, f"_{state_key}_last_click_ts", now)
        if last_id == item_id and (now - last_ts) * 1000 <= self.DOUBLE_CLICK_MS:
            setattr(app, f"_{state_key}_last_click_id", None)
            return True
        setattr(app, f"_{state_key}_last_click_id", item_id)
        return False

    def on_party_click(self, rec: dict):
        """Клик по строке контрагента: выделение, двойной клик — документы."""
        if self._is_double_click("registry", rec.get("id")):
            self.open_party_documents(rec)
        else:
            self.select_party(rec)

    def on_doc_click(self, rec: dict):
        """Клик по документу: выделение, двойной клик — открыть в Word."""
        if self._is_double_click("docs", rec.get("id")):
            self.app._lifecycle_flow.export_document(rec)
        else:
            self.select_document(rec)

    def select_party(self, rec: dict):
        """Клик по контрагенту: выделить строку (повторный клик снимает).

        Частичное обновление (п. 5.1): строки контрагентов не пересоздаются,
        у выделенной появляются кнопки действий. Проваливание остаётся
        прежним: выделение не влияет на текущий вид (2.19.0).
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

    def delete_selected(self) -> bool:
        """Клавиша Delete: удалить выделенный документ или контрагента.

        Документ имеет приоритет (он «мельче»): если выделены и документ,
        и контрагент, удаляется документ. Возвращает True, если удаление
        инициировано (диалог подтверждения — в контроллере жизненного
        цикла).

        Внутри вида «Документы контрагента» (проваливание, 2.19.0) Delete
        удаляет только выделенный документ — контрагент нельзя удалить
        случайно из его списка документов.
        """
        app = self.app
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
        """Текст строки контрагента в реестре (название, тип, ИНН, ID, адрес)."""
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

    def doc_row_text(self, rec: dict) -> str:
        """Текст строки документа: тип, номер, дата и срок действия (2.18.0)."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        text = f"{kind_label} № {number}"
        if date_text:
            text += f" от {date_text}"
        text += agreements.row_suffix(rec)
        return text

    def doc_row_color(self, rec: dict) -> tuple:
        """Цвет текста строки документа по состоянию срока (2.18.0).

        Просроченное доп. соглашение — красный, истекающее в течение
        месяца — амберный, остальное — обычный.
        """
        if rec.get("kind") == storage.KIND_ADDITIONAL:
            state = agreements.collect_state(rec)
            if state == "expired":
                return ui.COLORS["error_text"]
            if state == "soon":
                return ui.COLORS["warning_text"]
        return ui.COLORS["text"]

    @staticmethod
    def party_subtitle(rec: dict) -> str:
        """Подстрока контрагента в виде документов: тип, ИНН, ID (2.19.0)."""
        kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
        inn = str(rec.get("fields", {}).get(F.INN, "")).strip() or "—"
        return f"{kind}   ·   ИНН {inn}   ·   ID {rec.get('id')}"

    def refresh_selection(self):
        """Частичное обновление при смене выбора (п. 5.1).

        Строки остаются теми же виджетами — меняется только подсветка и
        видимость кнопок действий: в списке контрагентов обновляются
        строки контрагентов, в виде документов — строки документов.
        Полная перерисовка нужна только при изменении состава списка
        (поиск, добавление, удаление) и при проваливании/возврате.
        """
        app = self.app
        if self.is_drill_mode:
            selected_id = getattr(app, "_docs_selected_id", None)
            for row_id, entry in app._docs_rows.items():
                selected = selected_id == row_id
                btn = entry["btn"]
                try:
                    btn.configure(
                        text=self.doc_row_text(entry["rec"]),
                        font=ui.font(12, "bold" if selected else "normal"),
                        text_color=(ui.COLORS["text"] if selected
                                    else self.doc_row_color(entry["rec"])),
                    )
                except Exception:
                    pass
                self._toggle_row_actions(entry, selected)
        else:
            for party_id, entry in app._registry_rows.items():
                selected = app._registry_selected_id == party_id
                btn = entry["btn"]
                try:
                    btn.configure(
                        text=self.party_row_text(entry["rec"], selected),
                        font=ui.font(12, "bold" if selected else "normal"),
                        text_color=(ui.COLORS["text"] if selected
                                    else ui.COLORS["text_secondary"]),
                    )
                except Exception:
                    pass
                self._toggle_row_actions(entry, selected)
        app.root.after(60, app._refresh_scrollbars)

    @staticmethod
    def _toggle_row_actions(entry: dict, selected: bool) -> None:
        """Показывает кнопки действий только у выделенной строки (2.18.0)."""
        for key in ("edit_btn", "del_btn", "open_btn"):
            button = entry.get(key)
            if button is None:
                continue
            try:
                if selected:
                    button.grid()
                else:
                    button.grid_remove()
            except Exception:
                pass  # виджет уже разрушен

    def refresh(self):
        """Перерисовывает реестр целиком (поиск, изменения состава, проваливание).

        Длительные перерисовки (более 500 мс на больших объёмах) попадают
        в журнал (п. 10 этапа 5); клик по контрагенту/документу обновляет
        только подсветку строк (refresh_selection, п. 5.1).
        """
        with applog.timed("Перерисовка реестра", 500, "ui"):
            self.refresh_full()
        self.app._registry_rendered = True

    def refresh_full(self):
        app = self.app
        for child in app.registry_list.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        app._registry_rows = {}
        app._docs_rows = {}
        self._update_header_mode()

        # 2.19.0: проваливание — вид «Документы контрагента». Если контрагент
        # к этому моменту удалён (мягкое удаление), молча возвращаемся к списку.
        drill_id = getattr(app, "_registry_drill_party_id", None)
        if drill_id is not None:
            party = app.store.get_counterparty(drill_id)
            if party is None or storage.is_deleted(party):
                app._registry_drill_party_id = None
                app._docs_selected_id = None
            else:
                self.render_documents_view(party)
                app.root.after(60, app._refresh_scrollbars)
                return

        parties = self.parties()
        query = app.registry_search_var.get().strip()
        caption = (
            f"Найдено по запросу «{query}»: {len(parties)}"
            if query else f"Контрагенты ({len(parties)})"
        )
        ctk.CTkLabel(
            app.registry_list, text=caption,
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(6 * app.scale)))

        row_index = 1
        row_index = self.render_expiry_banner(row_index)
        if app._orphans:
            banner = ctk.CTkFrame(
                app.registry_list, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8,
            )
            banner.grid(row=row_index, column=0, sticky="ew", pady=(0, int(8 * app.scale)))
            banner.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                banner,
                text=f"В папке документов найдено файлов вне реестра: {len(app._orphans)}.\n"
                     "Их можно восстановить в реестр или убрать в служебную папку _trash.",
                font=ui.font(12), text_color=ui.COLORS["text"], anchor="w", justify="left",
                wraplength=int(500 * app.scale),
            ).grid(row=0, column=0, sticky="ew", padx=int(10 * app.scale),
                   pady=int(8 * app.scale))
            ui.ghost_button(
                banner, "Разобрать", self.open_reconcile_dialog,
                width=int(110 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(0, int(10 * app.scale)))
            row_index += 1

        if not parties:
            empty_text = (
                "Ничего не найдено — измените запрос (поиск идёт по названию и ИНН)."
                if query else
                "Пока нет ни одного контрагента.\n"
                "Добавьте первого через кнопку «Добавить контрагента»."
            )
            ctk.CTkLabel(
                app.registry_list, text=empty_text,
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=row_index, column=0, sticky="ew", pady=(0, int(12 * app.scale)))
            row_index += 1

        for rec in parties:
            selected = app._registry_selected_id == rec.get("id")
            text = self.party_row_text(rec, selected)
            row = ctk.CTkFrame(app.registry_list, fg_color="transparent")
            row.grid(row=row_index, column=0, sticky="ew", pady=1)
            row.grid_columnconfigure(0, weight=1)
            btn = ctk.CTkButton(
                row,
                text=text,
                anchor="w",
                font=ui.font(12, "bold" if selected else "normal"),
                fg_color="transparent",
                hover_color=ui.COLORS["sidebar_hover"],
                text_color=ui.COLORS["text"] if selected else ui.COLORS["text_secondary"],
                corner_radius=8,
                height=int(30 * app.scale),
                command=lambda r=rec: self.on_party_click(r),
            )
            btn.grid(row=0, column=0, sticky="ew")
            # Кнопки действий строки: создаются всегда, видны только у
            # выделенной строки (2.18.0).
            edit_btn = ui.ghost_button(
                row, "Редактировать", lambda r=rec: app._counterparty_flow.edit_counterparty(r),
                width=int(110 * app.scale), height=int(26 * app.scale),
            )
            edit_btn.grid(row=0, column=1, padx=(int(6 * app.scale), 0))
            del_btn = ui.ghost_button(
                row, "Удалить", lambda r=rec: app._lifecycle_flow.delete_counterparty(r),
                width=int(92 * app.scale), height=int(26 * app.scale),
            )
            del_btn.grid(row=0, column=2, padx=(int(6 * app.scale), 0))
            entry = {"rec": rec, "frame": row, "btn": btn,
                     "edit_btn": edit_btn, "del_btn": del_btn}
            app._registry_rows[rec.get("id")] = entry
            self._toggle_row_actions(entry, selected)
            row_index += 1

        row_index = self.render_unknown(row_index)
        self.render_trash(row_index)
        app.root.after(60, app._refresh_scrollbars)

    def _update_header_mode(self):
        """Шапка страницы по режиму: «Реестр» или «Документы — <имя>» (2.19.0)."""
        app = self.app
        title = getattr(app, "_registry_title_label", None)
        if title is None:
            return  # страница ещё не построена (юнит-тесты без Tk)
        back_btn = getattr(app, "_registry_back_btn", None)
        hint = getattr(app, "_registry_hint_label", None)
        search_bar = getattr(app, "_registry_search_bar", None)
        drill_id = getattr(app, "_registry_drill_party_id", None)
        if drill_id is not None:
            party = app.store.get_counterparty(drill_id)
            name = storage.display_name(party) if party is not None else ""
            title.configure(text=f"Документы — {name}" if name else "Документы")
            if hint is not None:
                hint.configure(
                    text="Клик — выделить документ («Открыть Word»/«Удалить», клавиша "
                         "Delete). Двойной клик — открыть документ в Word.",
                )
            if back_btn is not None:
                back_btn.pack(side="left", padx=(int(14 * app.scale), 0))
            if search_bar is not None:
                search_bar.pack_forget()  # поиск фильтрует контрагентов, не документы
        else:
            title.configure(text="Реестр")
            if hint is not None:
                hint.configure(
                    text="Клик — выделить контрагента (Редактировать/Удалить, клавиша "
                         "Delete). Двойной клик — открыть его документы.",
                )
            if back_btn is not None:
                back_btn.pack_forget()
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
    #  Секции реестра
    # ------------------------------------------------------------------ #
    def render_expiry_banner(self, row_index: int) -> int:
        """Баннер о сроках доп. соглашений над списком (2.18.0).

        Показывается, если есть просроченные или истекающие в течение
        месяца соглашения; клик по списку не нужен — баннер информационный,
        детали видны в строках документов (красный/амберный текст).
        """
        app = self.app
        try:
            buckets = agreements.collect(app.store)
        except Exception:
            return row_index
        expired, soon = buckets["expired"], buckets["soon"]
        if not expired and not soon:
            return row_index
        lines: list[str] = []
        if expired:
            lines.append(f"Просрочено доп. соглашений: {len(expired)} — "
                         "строки отмечены красным.")
        if soon:
            lines.append(f"Истекает в течение месяца: {len(soon)} — "
                         "строки отмечены жёлтым.")
        banner = ctk.CTkFrame(
            app.registry_list, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8,
        )
        banner.grid(row=row_index, column=0, sticky="ew", pady=(0, int(8 * app.scale)))
        banner.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            banner, text="\n".join(lines),
            font=ui.font(12), text_color=ui.COLORS["text"], anchor="w", justify="left",
            wraplength=int(560 * app.scale),
        ).grid(row=0, column=0, sticky="ew", padx=int(10 * app.scale), pady=int(8 * app.scale))
        return row_index + 1

    def render_documents_view(self, party: dict) -> None:
        """Вид «Документы — <имя>»: список документов контрагента (2.19.0).

        Проваливание: двойной клик по контрагенту переключает страницу на
        этот вид — список его документов вместо списка контрагентов.
        Возврат — кнопка «← Назад» в шапке или клавиша Escape. Строки
        документов кликом выделяются (кнопки «Открыть Word»/«Удалить» у
        выделенной), двойным кликом открываются в Word.
        """
        app = self.app
        header = ctk.CTkFrame(app.registry_list, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, int(6 * app.scale)))
        header.grid_columnconfigure(0, weight=1)
        docs = app.store.documents_for(party["id"])
        ctk.CTkLabel(
            header, text=f"Документы — {storage.display_name(party)} ({len(docs)})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header, text=self.party_subtitle(party),
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"], anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(int(2 * app.scale), 0))

        if not docs:
            ctk.CTkLabel(
                app.registry_list,
                text="У контрагента пока нет документов.\n"
                     "Создайте их на странице «Документ» или добавьте файлы "
                     "кнопкой «Добавить документы…».",
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=1, column=0, sticky="ew", pady=(int(12 * app.scale), 0))
            return

        for offset, rec in enumerate(docs, start=1):
            selected = getattr(app, "_docs_selected_id", None) == rec.get("id")
            text = self.doc_row_text(rec)
            color = self.doc_row_color(rec) if not selected else ui.COLORS["text"]
            row = ctk.CTkFrame(app.registry_list, fg_color="transparent")
            row.grid(row=offset, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            btn = ctk.CTkButton(
                row, text=text, anchor="w",
                font=ui.font(12, "bold" if selected else "normal"),
                fg_color="transparent",
                hover_color=ui.COLORS["sidebar_hover"],
                text_color=color,
                corner_radius=8,
                height=int(28 * app.scale),
                command=lambda d=rec: self.on_doc_click(d),
            )
            btn.grid(row=0, column=0, sticky="ew")
            open_btn = ui.ghost_button(
                row, "Открыть Word", lambda d=rec: app._lifecycle_flow.export_document(d),
                width=int(120 * app.scale), height=int(26 * app.scale),
            )
            open_btn.grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            del_btn = ui.ghost_button(
                row, "Удалить", lambda d=rec: app._lifecycle_flow.delete_document(d),
                width=int(92 * app.scale), height=int(26 * app.scale),
            )
            del_btn.grid(row=0, column=2, padx=(int(6 * app.scale), 0))
            entry = {"rec": rec, "frame": row, "btn": btn,
                     "open_btn": open_btn, "del_btn": del_btn}
            app._docs_rows[rec.get("id")] = entry
            self._toggle_row_actions(entry, selected)

    def render_unknown(self, row_index: int) -> int:
        """Неопознанные документы (папка «unknown») в конце реестра."""
        app = self.app
        docs = [
            rec for rec in app.store.documents
            if rec.get("kind") == storage.KIND_UNKNOWN
            and not storage.is_deleted(rec)
        ]
        if not docs:
            return row_index

        holder = ctk.CTkFrame(app.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(10 * app.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            holder, text=f"Неопознанные документы ({len(docs)})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(2 * app.scale)))
        ctk.CTkLabel(
            holder,
            text="Файлы, которые не удалось распознать, скопированы "
                 "в папку «unknown» внутри папки документов.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left",
            wraplength=int(560 * app.scale),
        ).grid(row=1, column=0, sticky="ew", pady=(0, int(4 * app.scale)))

        for offset, rec in enumerate(docs, start=2):
            source = str((rec.get("extra") or {}).get("source_file") or "")
            name = os.path.basename(source) if source else "Документ"
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=name, font=ui.font(12),
                text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            # Неопознанный файл нельзя запускать по клику: только показать в папке.
            ui.ghost_button(
                row, "Показать в папке",
                lambda s=source: app._show_in_folder(s),
                width=int(140 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda d=rec: app._lifecycle_flow.delete_document(d),
                width=int(92 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=2, padx=(int(6 * app.scale), 0))
        return row_index + 1

    # -- Корзина в реестре ------------------------------------------------ #
    def render_trash(self, row_index: int) -> int:
        """Раздел «Корзина»: восстановление и окончательная очистка (п. 2.6)."""
        app = self.app
        trashed_parties = app.store.trashed_counterparties()
        trashed_docs = [
            rec for rec in app.store.trashed_documents()
            if not rec.get("deleted_with_party")
        ]
        if not trashed_parties and not trashed_docs:
            return row_index

        holder = ctk.CTkFrame(app.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(12 * app.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        total = len(trashed_parties) + len(trashed_docs)
        ctk.CTkLabel(
            holder, text=f"Корзина ({total})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ui.ghost_button(
            holder, "Очистить корзину", app._lifecycle_flow.empty_trash,
            width=int(150 * app.scale), height=int(26 * app.scale),
        ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
        ctk.CTkLabel(
            holder,
            text="Записи удалены из реестра, файлы на диске остались. "
                 "Восстановление вернёт документы в работу.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left", wraplength=int(560 * app.scale),
        ).grid(row=1, column=0, columnspan=2, sticky="ew",
               pady=(int(2 * app.scale), int(4 * app.scale)))

        offset = 2
        for party in trashed_parties:
            docs = app.store.documents_for(party.get("id"), include_deleted=True)
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, columnspan=2, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row,
                text=f"Контрагент «{storage.display_name(party)}» · документов: {len(docs)}",
                font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Восстановить", lambda r=party: app._lifecycle_flow.restore_counterparty(r),
                width=int(120 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda r=party: app._lifecycle_flow.purge_counterparty(r),
                width=int(140 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=2, padx=(int(6 * app.scale), 0))
            offset += 1

        for rec in trashed_docs:
            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            number = rec.get("number") or "б/н"
            date_text = rec.get("date") or ""
            party_name = app.store.party_name(rec.get("party_id"))
            text = f"{kind_label} № {number}"
            if date_text:
                text += f" от {date_text}"
            text += f" · {party_name}"
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, columnspan=2, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=text, font=ui.font(12),
                text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Восстановить", lambda d=rec: app._lifecycle_flow.restore_document(d),
                width=int(120 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda d=rec: app._lifecycle_flow.purge_document(d),
                width=int(140 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=2, padx=(int(6 * app.scale), 0))
            offset += 1
        return row_index + 1

    # ------------------------------------------------------------------ #
    #  Сверка реестра с папкой документов
    # ------------------------------------------------------------------ #
    def scan_consistency(self) -> None:
        """Фоновая сверка: ищет файлы документов вне реестра (п. 2.4)."""
        app = self.app
        try:
            base = app._settings_flow.documents_dir(interactive=False)
            orphans = reconcile.find_orphans(base, app.store)
        except Exception as exc:  # noqa: BLE001 — фоновая задача не должна падать
            app._log.warning("Сверка реестра с папкой не удалась: %s", type(exc).__name__)
            return
        app._orphans = orphans
        if orphans:
            app._log.info("Файлов вне реестра: %d", len(orphans))
        if "registry" in app.pages:
            self.refresh()

    def open_reconcile_dialog(self) -> None:
        """Диалог разбора файлов вне реестра: восстановить или убрать в _trash."""
        app = self.app
        base = app._settings_flow.documents_dir()
        orphans = reconcile.find_orphans(base, app.store)
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
            if app.store.last_error:
                ui.alert(
                    app.root, "Восстановление не удалось",
                    app.store.last_error, "error",
                )
            else:
                ui.alert(
                    app.root, "Реестр восстановлен",
                    f"Добавлено записей в реестр: {added}.\n\n"
                    "Проверьте типы документов и контрагентов в реестре.",
                    "success",
                )
            app._orphans = reconcile.find_orphans(base, app.store)
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
            app._orphans = reconcile.find_orphans(base, app.store)
            self.refresh()
