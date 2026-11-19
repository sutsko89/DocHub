"""Карточка контрагента — шестой контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет страницей «Добавить контрагента»
и окном редактирования карточки: сборка страницы и полей, сбор данных,
валидация с подсветкой ошибок, выбор при дубликате ИНН (п. 4.1), сохранение
через DataStore и редактирование существующего контрагента.

Контроллер получает ссылку на приложение и обращается к общим службам через
него (``self.app.store``, ``self.app.root`` …), а вызовы, которые подменяет
smoke-тест, всегда идут через ``self.app.*``:

* ``ui.alert`` — атрибут общего модуля ui_macos, подмена ``m.ui.alert``
  видна контроллеру напрямую;
* выбор исхода при дубликате ИНН — через ``self.app._choice_dialog``
  (instance-атрибут приложения, его подменяют smoke-тест и тесты);
* сохранение и поиск дубликатов — через ``self.app.store`` (настоящий
  DataStore);
* кнопка «Заполнить по ИНН» привязана к ``app._dadata_flow.fill_from_inn``:
  он передаёт фабрику потока из main на каждый вызов, а smoke-тест подменяет
  именно ``main.threading.Thread`` (шаг 2 декомпозиции).

Состояние страницы остаётся на приложении с прежними именами
(``_cp_entries``, ``_cp_error_vars``, ``_cp_frames``, ``_cp_type_var``,
``_cp_type_segmented``, ``_cp_status_var``): его читают и меняют
DadataController («Заполнить по ИНН»), защита несохранённых данных (п. 4.2,
4.4) и smoke-тест — контроллер владеет кодом, а не данными страницы.
Контроллер создаётся в ``__init__`` ДО построения страниц:
``_build_counterparty_page`` уже идёт через контроллер.

Зависимости от модуля main — только поздним импортом: ширина полей
(main.ENTRY_WIDTH) и форматирование телефона при вводе
(main.make_phone_formatter) берутся через ``import main`` внутри методов.
Обычный верхнеуровневый импорт дал бы цикл, а к моменту вызова метода main
уже полностью инициализирован. Остальные зависимости — модули services,
storage, ui_macos, document_builder — импортируются напрямую, как в
предыдущих контроллерах.

Поведение один в один повторяет код, вынесенный из main.py версии 2.6.0:
честная обработка сбоя чтения полей (аудит 2.0.1 — сбойное поле не
превращается в пустую строку), выбор из трёх исходов при дубликате ИНН
(п. 4.1), удаление «фантома» при сбое записи реестра и окно редактирования,
меняющее только карточку — документы хранят собственную копию реквизитов
на момент создания.
"""

from __future__ import annotations

import customtkinter as ctk
import tkinter as tk

import services
import storage
import ui_macos as ui
from document_builder import F, digits_only
from services import validate_party_fields


class CounterpartyController:
    """Страница «Добавить контрагента» и редактирование карточки.

    Состояние страницы (виджеты и переменные) живёт на приложении —
    контроллер получает его через ``self.app.*`` на каждый вызов.
    """

    def __init__(self, app) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Страница «Добавить контрагента»
    # ------------------------------------------------------------------ #
    def build_counterparty_page(self):
        """Страница «Добавить контрагента» в основном окне программы."""
        app = self.app
        page = ctk.CTkFrame(app.content, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        app.pages["counterparty"] = page

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * app.scale), pady=(int(18 * app.scale), int(6 * app.scale)))
        ctk.CTkLabel(
            header, text="Добавить контрагента", font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text="Заполните карточку — она сохранится в реестре программы.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * app.scale), 0))

        app._cp_type_var = tk.StringVar(value="Юр. лицо")
        # variable обязателен: без привязки виджет не обновляет _cp_type_var,
        # и «Заполнить по ИНН»/сохранение всегда считали тип «Юр. лицо» —
        # поиск по ИНН у ИП не работал (читалось пустое поле юрлица).
        app._cp_type_segmented = ui.segmented(
            header, ["Юр. лицо", "ИП"], command=self.on_type_change,
            width=int(240 * app.scale), variable=app._cp_type_var,
        )
        app._cp_type_segmented.set("Юр. лицо")
        app._cp_type_segmented.pack(anchor="w", pady=(int(8 * app.scale), 0))

        ctk.CTkLabel(
            header, textvariable=app._cp_status_var, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(4 * app.scale), 0))

        scroller = ctk.CTkScrollableFrame(
            page, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        scroller.pack(fill="both", expand=True, padx=int(24 * app.scale), pady=(int(4 * app.scale), int(12 * app.scale)))
        app.scrollers.append(scroller)

        app._cp_frames: dict[str, ctk.CTkFrame] = {}
        for mode in ("ooo", "ip"):
            frame = ctk.CTkFrame(scroller, fg_color="transparent")
            frame.grid(row=0, column=0, sticky="nsew")
            card = ui.Card(frame, compact=True)
            card.pack(fill="x", pady=(0, int(12 * app.scale)))
            self.build_cp_fields(card, mode)
            app._cp_frames[mode] = frame

        self.on_type_change("Юр. лицо")

    def build_cp_fields(self, card, mode: str):
        import main  # поздний импорт: константы разметки main.py (см. докстринг)

        app = self.app
        error_var = self.error_var_for(mode)

        # -- Коды -------------------------------------------------------- #
        card.section("Коды")
        inn_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        inn = ui.entry(
            inn_holder, width=int(180 * app.scale),
            validatecommand=app._digits_vcmd_for(12),
        )
        ui.enable_sanitized_paste(inn.entry, digits_only, max_len=12)
        inn.pack(side="left")
        app._cp_entries[mode][F.INN] = inn
        ui.ghost_button(
            inn_holder, "Заполнить по ИНН", app._dadata_flow.fill_from_inn,
            width=int(152 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(12 * app.scale), 0))
        card.row(
            "ИНН *", inn_holder, stretch=True,
            hint="Кнопка подтягивает название, адрес, директора и банк из ЕГРЮЛ/ЕГРИП.",
            error_var=error_var(F.INN),
        )

        if mode == "ooo":
            self.add_field(card, mode, F.KPP, F.KPP, kind="digits", max_len=9,
                           width=int(140 * app.scale), required=True, stretch=False)
            self.add_field(card, mode, F.OGRN, F.OGRN, kind="digits", max_len=13,
                           width=int(180 * app.scale), required=True, stretch=False)
        else:
            self.add_field(card, mode, F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                           width=int(180 * app.scale), required=True, stretch=False)

        # -- Организация ------------------------------------------------- #
        card.section("Организация" if mode == "ooo" else "Предприниматель")
        if mode == "ooo":
            self.add_field(card, mode, F.FULL_NAME, "Полное название", required=True)
            self.add_field(card, mode, F.SHORT_NAME, "Сокращённое название", required=True)
        else:
            self.add_field(card, mode, F.IP_NAME, F.IP_NAME, required=True)

        # -- Руководитель ------------------------------------------------ #
        if mode == "ooo":
            card.section("Руководитель")
            self.add_field(card, mode, F.DIRECTOR, F.DIRECTOR, required=True)
            widget = ui.option_menu(card.body, ["Устав", "Доверенность"],
                                    width=int(main.ENTRY_WIDTH * app.scale))
            widget.set("Устав")
            app._cp_entries[mode][F.BASIS] = widget
            card.row(F.BASIS, widget)

        # -- Адреса ------------------------------------------------------ #
        card.section("Адреса")
        self.add_field(card, mode, F.LEGAL_ADDRESS, F.LEGAL_ADDRESS, required=True)
        if mode == "ooo":
            self.add_field(card, mode, F.FACT_ADDRESS, F.FACT_ADDRESS)
        else:
            card.row(
                F.FACT_ADDRESS,
                ctk.CTkLabel(
                    card.body, text="как юридический", font=ui.font(13),
                    text_color=ui.COLORS["text_tertiary"], anchor="w",
                ),
            )

        # -- Контакты ---------------------------------------------------- #
        # Email и телефон необязательны — юридически значимые поля выше (п. 4.5).
        card.section("Контакты")
        self.add_field(card, mode, F.EMAIL, F.EMAIL, kind="email",
                       hint="Необязательно.")
        self.add_field(card, mode, F.PHONE, F.PHONE, kind="phone",
                       hint="Необязательно.")

        # -- Банк -------------------------------------------------------- #
        card.section("Банковские реквизиты")
        self.add_field(card, mode, F.ACCOUNT, F.ACCOUNT, kind="digits", max_len=20,
                       width=int(260 * app.scale), required=True)
        self.add_field(card, mode, F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                       width=int(260 * app.scale), required=True)
        self.add_field(card, mode, F.BANK, F.BANK, required=True)
        self.add_field(card, mode, F.BIK, F.BIK, kind="digits", max_len=9,
                       width=int(140 * app.scale), required=True, stretch=False)

    def add_field(self, card, mode: str, key: str, label: str, kind="text",
                  width=None, required=False, stretch=True, max_len=None,
                  hint: str | None = None):
        import main  # поздний импорт: ширина поля и формат телефона из main.py

        app = self.app
        width = width or int(main.ENTRY_WIDTH * app.scale)
        validate = None
        if kind == "digits":
            validate = app._digits_vcmd_for(max_len) if max_len else (app._digits_vcmd, "%P")
        elif kind == "phone":
            validate = (app._phone_vcmd, "%P")

        widget = ui.entry(card.body, width=width, validatecommand=validate)
        if kind == "digits":
            # Валидатор отбрасывает вставку целиком, если в буфере есть
            # недопустимые символы («р/с 40702810…»). Вставляем только цифры.
            ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
        elif kind == "phone":
            phone_var = tk.StringVar()
            widget.entry.configure(textvariable=phone_var)
            phone_var.trace_add("write", main.make_phone_formatter(phone_var, widget.entry))
            ui.enable_sanitized_paste(
                widget.entry,
                lambda s: "".join(ch for ch in str(s) if ch.isdigit() or ch in "+-()/. "),
            )

        app._cp_entries[mode][key] = widget
        card.row(label, widget, stretch=stretch, required=required,
                 error_var=self.error_var_for(mode)(key), hint=hint)
        return widget

    def error_var_for(self, mode: str):
        def get(key: str) -> tk.StringVar:
            return self.app._cp_error_vars[mode].setdefault(key, tk.StringVar(value=""))
        return get

    def on_type_change(self, value: str):
        mode = "ip" if value == "ИП" else "ooo"
        self.app._cp_frames[mode].tkraise()

    def collect(self, mode: str) -> dict | None:
        """Собирает данные карточки в словарь.

        Аудит 2.0.1: раньше ошибка чтения виджета молча превращалась в пустую
        строку — сбойное поле обнулялось, валидация честно сообщала
        «Заполните поле» о поле, которое пользователь заполнил. Теперь при
        любой ошибке чтения возвращается None (явный маркер сбоя), тип
        исключения попадает в журнал, а вызывающий код обязан сообщить
        пользователю о проблеме вместо тихой записи пустых значений.
        """
        data = {}
        for key, widget in self.app._cp_entries[mode].items():
            try:
                data[key] = widget.get().strip()
            except Exception as exc:
                self.app._log.warning(
                    "Не удалось прочитать поле карточки «%s»: %s",
                    key, type(exc).__name__,
                )
                return None
        if mode == "ip":
            data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")
        return data

    def save_card(self):
        app = self.app
        mode = "ip" if app._cp_type_var.get() == "ИП" else "ooo"
        is_ip = mode == "ip"
        data = self.collect(mode)
        if data is None:
            # Сбор данных сломался — тихо сохранять «пустую» карточку нельзя.
            ui.alert(
                app.root, "Не удалось прочитать поля формы",
                "Часть полей карточки недоступна (ошибка интерфейса).\n\n"
                "Переоткройте страницу «Добавить контрагента» и попробуйте "
                "ещё раз. Подробности — в журнале (%APPDATA%\\DogovorGenerator\\logs).",
                "error",
            )
            return

        errors = validate_party_fields(data, is_ip)
        if errors:
            for var in app._cp_error_vars[mode].values():
                var.set("")
            for widget in app._cp_entries[mode].values():
                if hasattr(widget, "set_error"):
                    widget.set_error(False)
            for key, message in errors.items():
                self.error_var_for(mode)(key).set(message)
                widget = app._cp_entries[mode].get(key)
                if widget is not None and hasattr(widget, "set_error"):
                    widget.set_error(True)
            first = app._cp_entries[mode].get(next(iter(errors)))
            if first is not None:
                try:
                    first.focus_set()
                except Exception:
                    pass
            ui.alert(
                app.root, "Ошибки в данных",
                "Проверьте подсвеченные поля:\n\n"
                + "\n".join(f"• {error}" for error in errors.values()),
                "error",
            )
            return

        # Дубликаты по ИНН: молча создавать вторую карточку нельзя (п. 4.1).
        inn = data.get(F.INN, "").strip()
        duplicate = app.store.find_by_inn(inn)
        if duplicate is not None:
            choice = app._choice_dialog(
                "Контрагент с таким ИНН уже есть",
                f"ИНН {inn} уже указан у «{storage.display_name(duplicate)}» "
                f"(ID {duplicate.get('id')}).\n\n"
                "Открыть существующего или всё равно создать дубль?",
                [
                    ("open", "Открыть существующего", True),
                    ("create", "Создать всё равно", False),
                    ("cancel", "Отмена", False),
                ],
            )
            if choice != "create":
                if choice == "open":
                    # Пользователь явно выбрал существующего — введённые
                    # значения карточки больше не нужны.
                    app._doc_context_flow.reset_form_baseline("counterparty")
                    app._doc_context_flow.set_current_party(duplicate)
                    app._select_page("contract")
                return

        rec = app.store.add_counterparty(mode, data)
        if app.store.last_error:
            # Не оставляем «фантом» в памяти, если запись не легла на диск.
            try:
                app.store.counterparties.remove(rec)
            except ValueError:
                pass
            ui.alert(
                app.root, "Не удалось сохранить контрагента",
                app.store.last_error, "error",
            )
            return
        app._doc_context_flow.reset_form_baseline("counterparty")
        app._doc_context_flow.refresh_party_combo_values()
        app._update_status()
        app._registry_flow.refresh()
        app._doc_context_flow.set_current_party(rec)
        app._notify(f"{storage.display_name(rec)} добавлен в реестр программы.")
        # Сразу переходим к договору с уже выбранным контрагентом.
        app._select_page("contract")

    # ------------------------------------------------------------------ #
    #  Редактирование существующего контрагента
    # ------------------------------------------------------------------ #
    def edit_counterparty(self, rec: dict):
        """Окно редактирования контрагента: поля карточки + «Сохранить»/«Отмена».

        Изменяется только карточка в реестре. Ранее созданные документы не
        затрагиваются: каждый документ хранит собственную копию реквизитов
        на момент создания.
        """
        import main  # поздний импорт: ширина поля и формат телефона из main.py

        app = self.app
        mode = "ip" if rec.get("type") == "ip" else "ooo"
        is_ip = mode == "ip"
        saved = rec.get("fields", {})

        win = ctk.CTkToplevel(app.root)
        win.title(f"Редактирование контрагента — {storage.display_name(rec)}")
        win.geometry(f"{int(760 * app.scale)}x{int(680 * app.scale)}")
        win.minsize(int(660 * app.scale), int(520 * app.scale))
        win.transient(app.root)
        win.after(120, win.grab_set)

        header = ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * app.scale), pady=(int(18 * app.scale), int(6 * app.scale)))
        ctk.CTkLabel(
            header,
            text=f"Редактирование контрагента · {'ИП' if is_ip else 'Юр. лицо'}",
            font=ui.font(18, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=24,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Изменится только карточка в реестре. Ранее созданные документы останутся без изменений.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * app.scale), 0))

        scroller = ctk.CTkScrollableFrame(
            win, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        scroller.pack(fill="both", expand=True, padx=int(24 * app.scale), pady=(int(4 * app.scale), int(8 * app.scale)))
        app.scrollers.append(scroller)

        card = ui.Card(scroller, compact=True)
        card.pack(fill="x", pady=(0, int(8 * app.scale)))

        entries: dict[str, tk.Widget] = {}
        error_vars: dict[str, tk.StringVar] = {}

        def error_var(key: str) -> tk.StringVar:
            return error_vars.setdefault(key, tk.StringVar(value=""))

        def prefill(key: str, widget):
            value = str(saved.get(key, "")).strip()
            if not value:
                return
            try:
                widget.delete(0, tk.END)
                widget.insert(0, value)
            except Exception:
                try:
                    widget.set(value)
                except Exception:
                    pass

        def add_field(key, label, kind="text", width=None, required=False,
                      stretch=True, max_len=None, hint=None):
            width = width or int(main.ENTRY_WIDTH * app.scale)
            validate = None
            if kind == "digits":
                validate = app._digits_vcmd_for(max_len) if max_len else (app._digits_vcmd, "%P")
            elif kind == "phone":
                validate = (app._phone_vcmd, "%P")
            widget = ui.entry(card.body, width=width, validatecommand=validate)
            if kind == "digits":
                # Валидатор отбрасывает вставку целиком, если в буфере есть
                # недопустимые символы. Вставляем только цифры.
                ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
            elif kind == "phone":
                phone_var = tk.StringVar()
                widget.entry.configure(textvariable=phone_var)
                phone_var.trace_add("write", main.make_phone_formatter(phone_var, widget.entry))
                ui.enable_sanitized_paste(
                    widget.entry,
                    lambda s: "".join(ch for ch in str(s) if ch.isdigit() or ch in "+-()/. "),
                )
            prefill(key, widget)
            entries[key] = widget
            card.row(label, widget, stretch=stretch, required=required,
                     error_var=error_var(key), hint=hint)
            return widget

        # -- Коды -------------------------------------------------------- #
        card.section("Коды")
        add_field(F.INN, F.INN, kind="digits", max_len=12,
                  width=int(180 * app.scale), stretch=False, required=True)
        if is_ip:
            add_field(F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                      width=int(180 * app.scale), stretch=False, required=True)
        else:
            add_field(F.KPP, F.KPP, kind="digits", max_len=9,
                      width=int(140 * app.scale), stretch=False, required=True)
            add_field(F.OGRN, F.OGRN, kind="digits", max_len=13,
                      width=int(180 * app.scale), stretch=False, required=True)

        # -- Организация / Предприниматель -------------------------------- #
        card.section("Организация" if not is_ip else "Предприниматель")
        if is_ip:
            add_field(F.IP_NAME, F.IP_NAME, required=True)
        else:
            add_field(F.FULL_NAME, "Полное название", required=True)
            add_field(F.SHORT_NAME, "Сокращённое название", required=True)

        # -- Руководитель ------------------------------------------------- #
        if not is_ip:
            card.section("Руководитель")
            add_field(F.DIRECTOR, F.DIRECTOR, required=True)
            widget = ui.option_menu(card.body, ["Устав", "Доверенность"],
                                    width=int(main.ENTRY_WIDTH * app.scale))
            basis = str(saved.get(F.BASIS, "")).strip()
            widget.set(basis if basis in ("Устав", "Доверенность") else "Устав")
            entries[F.BASIS] = widget
            card.row(F.BASIS, widget)

        # -- Адреса ------------------------------------------------------- #
        card.section("Адреса")
        add_field(F.LEGAL_ADDRESS, F.LEGAL_ADDRESS, required=True)
        if is_ip:
            card.row(
                F.FACT_ADDRESS,
                ctk.CTkLabel(
                    card.body, text="как юридический", font=ui.font(13),
                    text_color=ui.COLORS["text_tertiary"], anchor="w",
                ),
            )
        else:
            add_field(F.FACT_ADDRESS, F.FACT_ADDRESS)

        # -- Контакты ----------------------------------------------------- #
        # Email и телефон необязательны (п. 4.5).
        card.section("Контакты")
        add_field(F.EMAIL, F.EMAIL, kind="email", hint="Необязательно.")
        add_field(F.PHONE, F.PHONE, kind="phone", hint="Необязательно.")

        # -- Банк ---------------------------------------------------------- #
        card.section("Банковские реквизиты")
        add_field(F.ACCOUNT, F.ACCOUNT, kind="digits", max_len=20,
                  width=int(260 * app.scale), required=True)
        add_field(F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                  width=int(260 * app.scale), required=True)
        add_field(F.BANK, F.BANK, required=True)
        add_field(F.BIK, F.BIK, kind="digits", max_len=9,
                  width=int(140 * app.scale), stretch=False, required=True)

        def highlight_errors(errors: dict[str, str]):
            for var in error_vars.values():
                var.set("")
            for widget in entries.values():
                if hasattr(widget, "set_error"):
                    widget.set_error(False)
            for key, message in errors.items():
                error_var(key).set(message)
                widget = entries.get(key)
                if widget is not None and hasattr(widget, "set_error"):
                    widget.set_error(True)
            first = entries.get(next(iter(errors)))
            if first is not None:
                try:
                    first.focus_set()
                except Exception:
                    pass

        def save():
            data = {}
            for key, widget in entries.items():
                try:
                    data[key] = widget.get().strip()
                except Exception:
                    data[key] = ""
            if is_ip:
                data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")

            errors = validate_party_fields(data, is_ip)
            # ИНН должен оставаться уникальным в реестре.
            duplicate = app.store.find_by_inn(data.get(F.INN, "").strip())
            if duplicate is not None and duplicate.get("id") != rec.get("id"):
                errors.setdefault(F.INN, "Контрагент с таким ИНН уже есть в реестре")
            if errors:
                highlight_errors(errors)
                ui.alert(
                    win, "Ошибки в данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {message}" for message in errors.values()),
                    "error",
                )
                return

            # Обновляем только карточку контрагента. Документы не трогаем:
            # они хранят собственную копию реквизитов на момент создания.
            updated = app.store.update_counterparty(rec["id"], fields=data)
            if app.store.last_error:
                ui.alert(
                    win, "Не удалось сохранить изменения",
                    app.store.last_error, "error",
                )
                return
            name = storage.display_name(updated or rec)
            # Переименование не должно «терять» выбранного контрагента (п. 4.2).
            if app._current_party_id == rec.get("id"):
                app._doc_context_flow.set_current_party(updated or rec)
            else:
                app._doc_context_flow.update_current_party_labels()
            app._doc_context_flow.refresh_party_combo_values()
            app._registry_flow.refresh()
            app._doc_page_flow.preview_contract()
            app._doc_page_flow.preview_additional()
            app._update_status()
            win.destroy()
            app._notify(
                f"Карточка «{name}» обновлена. Ранее созданные документы не изменялись."
            )

        footer = ctk.CTkFrame(win, fg_color="transparent")
        footer.pack(fill="x", padx=int(24 * app.scale), pady=(0, int(16 * app.scale)))
        ui.secondary_button(
            footer, "Отмена", win.destroy,
            width=int(140 * app.scale), height=int(30 * app.scale),
        ).pack(side="right")
        ui.primary_button(
            footer, "Сохранить", save,
            width=int(170 * app.scale), height=int(30 * app.scale),
        ).pack(side="right", padx=(0, int(10 * app.scale)))
