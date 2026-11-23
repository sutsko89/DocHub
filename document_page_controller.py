"""Страницы «Договор» и «Доп. соглашение» — седьмой контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет сборкой страниц договора поставки
и доп. соглашения, отложенными текстовыми предпросмотрами (п. 5.6), поиском
контрагента по ИНН со страницы договора, подстановкой договора-основания
при выборе контрагента (п. 5.1, п. 4.2) и полями «сумма/срок прописью»
доп. соглашения.

Контроллер получает ссылку на приложение и обращается к общим службам
через него (``self.app.store``, ``self.app.root`` …). Состояние страниц
остаётся на приложении с прежними именами (``contract_preview``,
``contract_party_var``, ``addl_contract_info_var`` …, таймеры
``_preview_contract_after`` / ``_preview_additional_after`` — как
``_spec_preview_after`` у DocumentSaveController): его читают защита
несохранённых данных (п. 4.2, 4.4), контроллер сохранения документов
(шаг 5) и smoke-тест — контроллер владеет кодом, а не данными страницы.

Связность минимальна и идёт через фасады приложения:

* единый «текущий контрагент» — ``app._doc_context_flow`` (party_from_var,
  set_current_party, on_party_selected, contract_from_var): выбор,
  сделанный на одной странице
  документов, переносится на остальные (п. 4.2), а трассы переменных
  привязаны к фасадам ContractApp, чтобы подмена маршрута сохранялась;
* печать предпросмотра — ``app._doc_save_flow.render_contract_docx`` /
  ``render_additional_docx`` (контроллер сохранения, шаг 5);
  сам предпросмотр — ``docx_to_text`` из document_builder, печатные формы
  живут в services/document_builder (п. 9.1, 9.2);
* ``ui.alert`` — атрибут общего модуля ui_macos: подмена ``m.ui.alert``
  видна контроллеру напрямую;
* записи ``app._page_widgets`` и ``app._error_var_for`` — общий реестр
  inline-валидации страниц, его читают защита несохранённых данных и
  подсветка ошибок.

Ширины разметки (``main.ENTRY_WIDTH``, ``main.SHORT_WIDTH``,
``main.LEFT_PANEL_WIDTH``) и хелпер курсора
``main.restore_caret_after_reformat`` берутся поздним импортом main:
модуль main импортирует этот контроллер, поэтому верхнеуровневый импорт
дал бы цикл, а к моменту любого вызова метода main уже полностью
инициализирован.

Поведение один в один повторяет код, вынесенный из main.py версии 2.7.0.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

import customtkinter as ctk
import tkinter as tk

import ui_macos as ui
from document_builder import (
    F, MONTHS_GENITIVE, digits_only, docx_to_text, format_number_grouped,
    number_to_words_ru, amount_to_words_ru, parse_valid_until,
    valid_until_default,
)


class DocumentPageController:
    """Страницы «Договор поставки» и «Доп. соглашение».

    Состояние страниц (виджеты, переменные, таймеры предпросмотров)
    живёт на приложении — контроллер получает его через ``self.app.*``
    на каждый вызов.
    """

    # Задержка отложенного предпросмотра: не собирать .docx на каждый
    # символ ввода (п. 5.6; прежнее значение из main.py).
    PREVIEW_DELAY_MS = 120

    def __init__(self, app: Any) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Страница «Договор поставки»
    # ------------------------------------------------------------------ #
    def build_contract_page(self):
        """Страница «Договор поставки» в основном окне программы."""
        app = self.app
        import main  # поздний импорт: константы разметки main.py (см. докстринг)

        _page, left, _right, textbox = app._make_page(
            "contract", "Договор поставки",
            "Выберите контрагента, укажите номер и дату. «Сохранить договор» запишет его "
            "в реестр и сразу откроет готовый документ в Word.",
            with_text=True, text_title="Предпросмотр договора (только чтение)",
        )
        app.contract_preview = textbox

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Контрагент")

        app.contract_party_var = tk.StringVar(value="")
        app.contract_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(main.ENTRY_WIDTH * app.scale), variable=app.contract_party_var,
        )
        card.wide(app.contract_party_combo, pady=(int(2 * app.scale), int(10 * app.scale)))

        # U3 (UX-A): общий блок поиска — теперь на всех страницах документов.
        self.build_party_search_row(card, "contract_search_var")

        # -- Договор ----------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Договор")

        app.contract_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(main.SHORT_WIDTH * app.scale),
            textvariable=app.contract_number_var,
        )
        app._page_widgets[("contract", F.NUMBER)] = number
        card.row(F.NUMBER, number, required=True,
                 error_var=app._error_var_for("contract", F.NUMBER))
        app.contract_number_var.trace_add(
            "write", lambda *_: app._on_doc_field_changed("contract", F.NUMBER))

        app.contract_day_var = tk.StringVar()
        app.contract_month_var = tk.StringVar(
            value=MONTHS_GENITIVE[datetime.now().month - 1])
        app._add_date_row(card, "contract", app.contract_day_var, app.contract_month_var)
        app.contract_day_var.trace_add(
            "write", lambda *_: app._on_doc_field_changed("contract", F.DAY))
        app.contract_month_var.trace_add(
            "write", lambda *_: app._on_doc_field_changed("contract", F.MONTH))

        app._doc_context_flow.refresh_party_combo_values()

        # Предпросмотр обновляется при любом изменении полей договора.
        app.contract_party_var.trace_add("write", self.preview_contract)
        app.contract_party_var.trace_add(
            "write", lambda *_: app._doc_context_flow.on_party_selected("contract"))
        app.contract_number_var.trace_add("write", self.preview_contract)
        app.contract_day_var.trace_add("write", self.preview_contract)
        app.contract_month_var.trace_add("write", self.preview_contract)
        self.preview_contract()

    def preview_contract(self, *_args):
        """Откладывает перерисовку предпросмотра договора.

        Таймер живёт на приложении (``app._preview_contract_after``):
        новый ввод отменяет ещё не сработавшую перерисовку (п. 5.6).
        """
        app = self.app
        if getattr(app, "_preview_contract_after", None):
            try:
                app.root.after_cancel(app._preview_contract_after)
            except Exception:
                pass
        app._preview_contract_after = app.root.after(
            self.PREVIEW_DELAY_MS, self.render_contract_preview)

    def render_contract_preview(self):
        """Собирает текст предпросмотра договора из текущих полей."""
        app = self.app
        app._preview_contract_after = None
        rec = app._doc_context_flow.party_from_var(app.contract_party_var)
        if rec is None:
            app._set_text(
                app.contract_preview,
                "Выберите контрагента и заполните поля слева — здесь появится текст договора.",
            )
            return
        data = dict(rec.get("fields", {}))
        data[F.NUMBER] = app.contract_number_var.get().strip()
        data[F.DAY] = app.contract_day_var.get().strip()
        data[F.MONTH] = app.contract_month_var.get().strip()
        is_ip = rec.get("type") == "ip"
        try:
            doc = app._doc_save_flow.render_contract_docx(data, is_ip)
            app._set_text(app.contract_preview, docx_to_text(doc))
        except Exception as exc:
            app._set_text(app.contract_preview, f"Не удалось собрать предпросмотр: {exc}")

    # ------------------------------------------------------------------ #
    #  Поиск контрагента по ИНН — общий блок страниц документов (U3)
    # ------------------------------------------------------------------ #
    def build_party_search_row(self, card, attr: str) -> None:
        """Строка «Поиск по ИНН» в карточке «Контрагент» (U3, UX-A).

        Переиспользует готовый блок страницы «Договор» на страницах
        «Доп. соглашение» и «Спецификация»: найденный контрагент
        подставляется сразу во все страницы (set_current_party), поэтому
        больше не нужно листать комбобокс, когда контрагентов много.

        ``attr`` — имя атрибута приложения для StringVar поиска;
        на «Договоре» сохраняется прежнее имя ``contract_search_var``
        (его читают тесты и smoke-сценарий).
        """
        app = self.app
        var = tk.StringVar()
        setattr(app, attr, var)
        holder = ctk.CTkFrame(card.body, fg_color="transparent")
        search_entry = ui.entry(
            holder, width=int(180 * app.scale), textvariable=var,
            validatecommand=app._digits_vcmd_for(12),
        )
        ui.enable_sanitized_paste(search_entry.entry, digits_only, max_len=12)
        search_entry.pack(side="left")
        try:  # Enter в поиске ищет, а не сохраняет документ
            search_entry.bind_entry(
                "<Return>", lambda _e, v=var: (self.find_party_by_inn(v), "break")[1])
            search_entry.bind_entry(
                "<KP_Enter>", lambda _e, v=var: (self.find_party_by_inn(v), "break")[1])
        except Exception:
            pass
        ui.ghost_button(
            holder, "Найти", lambda v=var: self.find_party_by_inn(v),
            width=int(92 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(10 * app.scale), 0))
        card.row(
            "Поиск по ИНН", holder, stretch=True,
            hint="Найдёт контрагента в реестре и подставит его во все страницы.",
        )

    def find_party_by_inn(self, source_var=None):
        """Ищет контрагента по ИНН в реестре и подставляет во все страницы (U3)."""
        app = self.app
        var = source_var if source_var is not None else getattr(app, "contract_search_var", None)
        if var is None:
            return
        inn = digits_only(var.get())
        rec = app.store.find_by_inn(inn)
        if rec is None:
            ui.alert(
                app.root, "Контрагент не найден",
                f"В реестре программы нет контрагента с ИНН {inn or '—'}.\n"
                "Добавьте его через кнопку «Добавить контрагента».",
                "error",
            )
            return
        app._doc_context_flow.set_current_party(rec)
        var.set("")

    # ------------------------------------------------------------------ #
    #  Страница «Доп. соглашение»
    # ------------------------------------------------------------------ #
    def build_additional_page(self):
        """Страница «Доп. соглашение» в основном окне программы."""
        app = self.app
        import main  # поздний импорт: константы разметки main.py (см. докстринг)

        _page, left, _right, textbox = app._make_page(
            "additional", "Доп. соглашение",
            "Выберите контрагента и договор-основание. «Сохранить» запишет документ "
            "в реестр и сразу откроет его в Word.",
            with_text=True, text_title="Предпросмотр доп. соглашения (только чтение)",
        )
        app.addl_preview = textbox

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Контрагент")

        app.addl_party_var = tk.StringVar(value="")
        app.addl_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(main.ENTRY_WIDTH * app.scale), variable=app.addl_party_var,
        )
        card.wide(app.addl_party_combo, pady=(int(2 * app.scale), int(10 * app.scale)))
        # U3 (UX-A): поиск по ИНН как на «Договоре» — не листать комбобокс.
        self.build_party_search_row(card, "addl_search_var")
        app.addl_party_var.trace_add("write", self.on_addl_party_changed)
        app.addl_party_var.trace_add(
            "write", lambda *_: app._doc_context_flow.on_party_selected("additional"))

        # -- Договор-основание -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Договор-основание")

        app.addl_contract_var = tk.StringVar(value="")
        app.addl_contract_combo = ui.option_menu(
            card.body, ["— нет сохранённых договоров —"],
            width=int(main.ENTRY_WIDTH * app.scale), variable=app.addl_contract_var,
        )
        card.wide(app.addl_contract_combo, pady=(int(2 * app.scale), int(4 * app.scale)))
        app.addl_contract_info_var = tk.StringVar(value="Договор-основание не выбран.")
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=app.addl_contract_info_var,
                font=ui.font(11), text_color=ui.COLORS["text_secondary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 40) * app.scale),
            ),
            pady=(int(2 * app.scale), int(2 * app.scale)),
        )
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Договоры берутся из реестра: сначала создайте договор поставки.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 40) * app.scale),
            ),
            pady=(int(4 * app.scale), int(6 * app.scale)),
        )

        # -- Условия ------------------------------------------------------ #
        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Условия")

        app.addl_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(main.SHORT_WIDTH * app.scale),
            textvariable=app.addl_number_var,
        )
        card.row("Номер доп. соглашения", number,
                 hint="Можно оставить пустым — сохранится как «б/н».")

        app._limit_var = tk.StringVar()
        app._limit_var.trace_add("write", self.format_limit)
        limit_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        app.limit_entry = ui.entry(
            limit_holder, width=int(220 * app.scale), textvariable=app._limit_var,
            validatecommand=(app._limit_vcmd, "%P"),
        )
        app.limit_entry.pack(side="left")
        ctk.CTkLabel(
            limit_holder, text="руб.", font=ui.font(12),
            text_color=ui.COLORS["text_tertiary"],
        ).pack(side="left", padx=(int(8 * app.scale), 0))
        app._page_widgets[("additional", F.LIMIT)] = app.limit_entry
        card.row(
            F.LIMIT, limit_holder, stretch=True, required=True,
            hint="Сумма по доп. соглашению к договору поставки.",
            error_var=app._error_var_for("additional", F.LIMIT),
        )

        app.limit_words_label = ctk.CTkLabel(
            card.body, text="Сумма прописью: —", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left", height=18,
        )
        card.wide(app.limit_words_label, pady=(int(4 * app.scale), int(8 * app.scale)))

        term_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        app.term_var = tk.StringVar()
        app.term_entry = ui.entry(
            term_holder, width=int(120 * app.scale), textvariable=app.term_var,
            validatecommand=(app._term_vcmd, "%P"),
        )
        app.term_entry.pack(side="left")
        ctk.CTkLabel(
            term_holder, text="дней", font=ui.font(12),
            text_color=ui.COLORS["text_tertiary"],
        ).pack(side="left", padx=(int(8 * app.scale), 0))
        app._page_widgets[("additional", F.TERM)] = app.term_entry
        card.row(
            F.TERM, term_holder, required=True,
            hint="Отсрочка платежа по доп. соглашению.",
            error_var=app._error_var_for("additional", F.TERM),
        )

        # -- Срок действия (2.18.0) --------------------------------------- #
        app.addl_valid_var = tk.StringVar(value=valid_until_default())
        valid_entry = ui.entry(
            card.body, width=int(120 * app.scale), textvariable=app.addl_valid_var,
        )
        app._page_widgets[("additional", F.VALID_UNTIL)] = valid_entry
        card.row(
            F.VALID_UNTIL, valid_entry, required=True,
            hint="Срок действия по п. 5 (формат ДД.ММ.ГГГГ). По умолчанию — "
                 "31 декабря текущего года.",
            error_var=app._error_var_for("additional", F.VALID_UNTIL),
        )

        app.words_label = ctk.CTkLabel(
            card.body, text="Срок прописью: —", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left", height=18,
        )
        card.wide(app.words_label, pady=(int(2 * app.scale), int(2 * app.scale)))

        app.term_var.trace_add("write", self.update_words)
        self.update_words()
        self.update_limit_words()

        # Предпросмотр обновляется при любом изменении данных доп. соглашения.
        app.addl_party_var.trace_add("write", self.preview_additional)
        app.addl_contract_var.trace_add("write", self.preview_additional)
        app.addl_contract_var.trace_add(
            "write", lambda *_: app._doc_context_flow.update_contract_info("additional"))
        app.addl_number_var.trace_add("write", self.preview_additional)
        app._limit_var.trace_add("write", self.preview_additional)
        app.term_var.trace_add("write", self.preview_additional)
        app.addl_valid_var.trace_add("write", self.preview_additional)
        self.preview_additional()

    def preview_additional(self, *_args):
        """Откладывает перерисовку предпросмотра доп. соглашения.

        Таймер живёт на приложении (``app._preview_additional_after``):
        новый ввод отменяет ещё не сработавшую перерисовку (п. 5.6).
        """
        app = self.app
        if getattr(app, "_preview_additional_after", None):
            try:
                app.root.after_cancel(app._preview_additional_after)
            except Exception:
                pass
        app._preview_additional_after = app.root.after(
            self.PREVIEW_DELAY_MS, self.render_additional_preview)

    def render_additional_preview(self):
        """Собирает текст предпросмотра доп. соглашения из текущих полей."""
        app = self.app
        app._preview_additional_after = None
        party = app._doc_context_flow.party_from_var(app.addl_party_var)
        contract = (
            app._doc_context_flow.contract_from_var(app.addl_contract_var, party["id"]) if party else None
        )
        if party is None or contract is None:
            app._set_text(
                app.addl_preview,
                "Выберите контрагента и договор-основание слева — здесь появится текст документа.",
            )
            return
        # LAZY-1: тело ленивого договора дочитывается единым getter'ом.
        if "data" not in contract:
            contract = {**contract, **app.store.document_body(contract.get("id"))}
        data = dict(contract.get("data") or {})
        is_ip = party.get("type") == "ip"
        try:
            doc = app._doc_save_flow.render_additional_docx(
                data, is_ip,
                app.addl_number_var.get().strip(),
                app._limit_var.get().strip(),
                app.term_var.get().strip(),
                app.addl_valid_var.get().strip(),
            )
            app._set_text(app.addl_preview, docx_to_text(doc))
        except Exception as exc:
            app._set_text(app.addl_preview, f"Не удалось собрать предпросмотр: {exc}")

    def on_addl_party_changed(self, *_args):
        """Обновляет список договоров выбранного контрагента (п. 5.1)."""
        app = self.app
        rec = app._doc_context_flow.party_from_var(app.addl_party_var)
        contracts = app.store.contracts_for(rec["id"]) if rec else []
        if contracts:
            values = [app._doc_context_flow.contract_display(c) for c in contracts]
        else:
            values = ["— нет сохранённых договоров —"]
        app.addl_contract_combo.set_values(values)
        app.addl_contract_var.set(values[0] if contracts else "")
        # Явно сообщаем, что договор подставлен автоматически (п. 4.2).
        if contracts:
            app.addl_contract_info_var.set(
                f"Договор-основание: {values[0]} — подставлен автоматически, проверьте."
            )
        else:
            app.addl_contract_info_var.set(
                "Договор-основание не выбран: сохраните договор поставки."
            )

    # ------------------------------------------------------------------ #
    #  Доп. соглашение: суммы прописью и формат лимита
    # ------------------------------------------------------------------ #
    def update_words(self, *_args):
        """«Срок прописью» под полем отсрочки платежа."""
        app = self.app
        value = app.term_var.get().strip()
        if not value:
            app.words_label.configure(text="Срок прописью: —")
        elif value.isdigit():
            app.words_label.configure(
                text=f"Срок прописью: {number_to_words_ru(int(value))}")
        else:
            app.words_label.configure(text="Срок прописью: введите целое число дней")

    def format_limit(self, *_args):
        """Форматирует лимит с разделителями разрядов и обновляет сумму прописью."""
        app = self.app
        import main  # поздний импорт: хелпер курсора main.py (см. докстринг)

        raw = app._limit_var.get()
        formatted = format_number_grouped(raw)
        if raw != formatted:
            entry = None
            caret = None
            try:
                entry = app.limit_entry.entry
                caret = entry.index(tk.INSERT)
            except Exception:
                pass
            app._limit_var.set(formatted)
            main.restore_caret_after_reformat(entry, caret, len(raw), len(formatted))
        self.update_limit_words()

    def update_limit_words(self):
        """«Сумма прописью» под полем лимита доп. соглашения."""
        app = self.app
        digits = digits_only(app._limit_var.get())
        if not digits:
            app.limit_words_label.configure(text="Сумма прописью: —")
            return
        try:
            app.limit_words_label.configure(
                text=f"Сумма прописью: {amount_to_words_ru(int(digits))}")
        except Exception:
            app.limit_words_label.configure(text="Сумма прописью: —")
