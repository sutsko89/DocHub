"""Страница «Реестр» — четвёртый контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет кодом страницы «Реестр»:
построение страницы, поиск по названию и ИНН с фильтрацией по мере ввода,
выбор контрагента с частичным обновлением (п. 5.1), рендер секций
«документы контрагента», «неопознанные» и «корзина», сверка реестра с папкой
документов и диалог разбора файлов вне реестра (п. 2.4).

Контроллер получает ссылку на приложение и обращается к общим службам через
него (``self.app.store``, ``self.app.root`` …). Состояние страницы остаётся
на приложении с прежними именами (``pages``, ``registry_search_var``,
``registry_list``, ``_registry_selected_id``, ``_registry_rows``,
``_registry_docs_frame``, ``_registry_rendered``, ``_orphans`` …): его
читают и меняют другие страницы, сценарии сохранения и smoke-тест —
контроллер владеет кодом, а не данными страницы.

Вызовы, которые подменяет smoke-тест, всегда идут через ``self.app.*``:
``_choice_dialog`` подменяется как instance-атрибут, кнопки действий
вызывают фасады приложения с прежними именами (``_edit_counterparty``,
``_delete_document``, ``export_document``, ``_restore_*`` …), а показ папки
для неопознанных файлов — через фасад ``ContractApp._show_in_folder``,
тело которого резолвит модульный ``main.show_in_folder`` на каждый вызов.
``ui.confirm`` / ``ui.alert`` — атрибуты общего модуля ui_macos, подмена
``m.ui.confirm`` видна контроллеру напрямую.

Поведение один в один повторяет код, вынесенный из main.py версии 2.4.0,
включая частичное обновление реестра при выборе (п. 5.1) и ленивый рендер
списка при старте (п. 9 этапа 5).
"""

from __future__ import annotations

import os
from typing import Any

import customtkinter as ctk
import tkinter as tk

import applog
import reconcile
import storage
import ui_macos as ui
from document_builder import F, digits_only
from services import safe_filename


class RegistryController:
    """Страница «Реестр»: построение, поиск, выбор, секции, сверка (п. 2.4, 5.1).

    Код страницы без её данных: все виджеты и переменные публикуются на
    приложении под прежними именами, поэтому сценарии сохранения, other
    страницы и smoke-тест продолжают работать без правок.
    """

    def __init__(self, app: Any) -> None:
        self.app = app

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
        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")
        ctk.CTkLabel(
            title_row, text="Реестр", font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        ).pack(side="left")
        ui.ghost_button(
            title_row, "Обновить", app.refresh_registry,
            width=int(110 * app.scale), height=int(28 * app.scale),
        ).pack(side="right")
        ui.ghost_button(
            title_row, "Добавить документы…", app._import_documents_dialog,
            width=int(180 * app.scale), height=int(28 * app.scale),
        ).pack(side="right", padx=(0, int(8 * app.scale)))
        ctk.CTkLabel(
            header, text="Список контрагентов. Выберите контрагента — ниже появятся его документы.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * app.scale), 0))

        search_bar = ctk.CTkFrame(page, fg_color="transparent")
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
            search_bar, "Найти", app._registry_apply_filter,
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
        app._registry_docs_row = 1
        app._registry_docs_frame = None
        # Ленивый рендер (п. 9 этапа 5): при старте список не рисуется —
        # реестр отрисуется при первом открытии страницы или при изменениях.
        app._registry_rendered = False

    # ------------------------------------------------------------------ #
    #  Реестр: список контрагентов и документы выбранного --------------- #
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

    def on_search_changed(self, *_args):
        """Фильтрует реестр по мере ввода с небольшой задержкой."""
        pending = getattr(self.app, "_registry_search_after", None)
        if pending:
            try:
                self.app.root.after_cancel(pending)
            except Exception:
                pass
        self.app._registry_search_after = self.app.root.after(200, self.app.refresh_registry)

    def apply_filter(self):
        pending = getattr(self.app, "_registry_search_after", None)
        if pending:
            try:
                self.app.root.after_cancel(pending)
            except Exception:
                pass
            self.app._registry_search_after = None
        self.app.refresh_registry()

    def reset_filter(self):
        self.app.registry_search_var.set("")
        self.apply_filter()

    def select_party(self, rec: dict):
        """Клик по контрагенту: показать/скрыть его документы.

        Частичное обновление (п. 5.1): строки контрагентов не пересоздаются,
        перерисовываются только подсветка выбора и секция документов.
        """
        if self.app._registry_selected_id == rec.get("id"):
            self.app._registry_selected_id = None
        else:
            self.app._registry_selected_id = rec.get("id")
            # Выбор из реестра переносится на страницы документов (п. 4.2).
            self.app._set_current_party(rec)
        if not self.app._registry_rows:
            # Страница ещё не отрисована (ленивый рендер) — рисуем целиком.
            self.app.refresh_registry()
            return
        self.refresh_selection()

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

    def refresh_selection(self):
        """Частичное обновление реестра при смене выбора (п. 5.1).

        Строки контрагентов остаются теми же виджетами (меняется только
        подсветка), секция документов выбранного пересобирается на своём
        месте. Полная перерисовка нужна только при изменении состава списка
        (поиск, добавление, удаление).
        """
        for party_id, entry in self.app._registry_rows.items():
            selected = self.app._registry_selected_id == party_id
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
        if self.app._registry_docs_frame is not None:
            try:
                self.app._registry_docs_frame.destroy()
            except Exception:
                pass
            self.app._registry_docs_frame = None
        self.render_documents(self.app._registry_docs_row)
        self.app.root.after(60, self.app._refresh_scrollbars)

    def refresh(self):
        """Перерисовывает реестр целиком (поиск, изменения состава).

        Длительные перерисовки (более 500 мс на больших объёмах) попадают
        в журнал (п. 10 этапа 5); клик по контрагенту перерисовывает только
        подсветку и секцию документов (refresh_selection, п. 5.1).
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
        app._registry_docs_frame = None

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
        app._registry_rows = {}
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
                command=lambda r=rec: self.select_party(r),
            )
            btn.grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Редактировать", lambda r=rec: app._edit_counterparty(r),
                width=int(110 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(6 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda r=rec: app._delete_counterparty(r),
                width=int(92 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=2, padx=(int(6 * app.scale), 0))
            app._registry_rows[rec.get("id")] = {"rec": rec, "frame": row, "btn": btn}
            row_index += 1

        app._registry_docs_row = row_index
        row_index = self.render_documents(row_index)
        row_index = self.render_unknown(row_index)
        self.render_trash(row_index)
        app.root.after(60, app._refresh_scrollbars)

    # ------------------------------------------------------------------ #
    #  Секции реестра
    # ------------------------------------------------------------------ #
    def render_documents(self, row_index: int) -> int:
        """Документы выбранного контрагента под списком. Возвращает следующий row."""
        app = self.app
        holder = ctk.CTkFrame(app.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(10 * app.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        app._registry_docs_frame = holder

        party = app.store.get_counterparty(app._registry_selected_id) \
            if app._registry_selected_id is not None else None
        if party is None:
            ctk.CTkLabel(
                holder, text="Выберите контрагента из списка — ниже появятся его документы.",
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=0, column=0, sticky="ew")
            return row_index + 1

        docs = app.store.documents_for(party["id"])
        ctk.CTkLabel(
            holder, text=f"Документы — {storage.display_name(party)} ({len(docs)})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(6 * app.scale)))

        if not docs:
            ctk.CTkLabel(
                holder, text="У контрагента пока нет документов.",
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=1, column=0, sticky="ew")
            return row_index + 1

        for offset, rec in enumerate(docs, start=1):
            kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
            number = rec.get("number") or "б/н"
            date_text = rec.get("date") or ""
            text = f"{kind_label} № {number}"
            if date_text:
                text += f" от {date_text}"

            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=text, font=ui.font(12), text_color=ui.COLORS["text"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Открыть Word", lambda d=rec: app.export_document(d),
                width=int(120 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda d=rec: app._delete_document(d),
                width=int(92 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=2, padx=(int(6 * app.scale), 0))
        return row_index + 1

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
                row, "Удалить", lambda d=rec: app._delete_document(d),
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
            holder, "Очистить корзину", app._empty_trash,
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
                row, "Восстановить", lambda r=party: app._restore_counterparty(r),
                width=int(120 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda r=party: app._purge_counterparty(r),
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
                row, "Восстановить", lambda d=rec: app._restore_document(d),
                width=int(120 * app.scale), height=int(26 * app.scale),
            ).grid(row=0, column=1, padx=(int(8 * app.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda d=rec: app._purge_document(d),
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
            base = app._documents_dir(interactive=False)
            orphans = reconcile.find_orphans(base, app.store)
        except Exception as exc:  # noqa: BLE001 — фоновая задача не должна падать
            app._log.warning("Сверка реестра с папкой не удалась: %s", type(exc).__name__)
            return
        app._orphans = orphans
        if orphans:
            app._log.info("Файлов вне реестра: %d", len(orphans))
        if "registry" in app.pages:
            app.refresh_registry()

    def open_reconcile_dialog(self) -> None:
        """Диалог разбора файлов вне реестра: восстановить или убрать в _trash."""
        app = self.app
        base = app._documents_dir()
        orphans = reconcile.find_orphans(base, app.store)
        app._orphans = orphans
        if not orphans:
            ui.alert(
                app.root, "Файлы вне реестра",
                "Все файлы документов учтены в реестре.", "info",
            )
            app.refresh_registry()
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
            app._refresh_party_combo_values()
            app.refresh_registry()
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
            app.refresh_registry()
