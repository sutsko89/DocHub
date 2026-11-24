"""Карточка контрагента — шестой контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет страницей «Добавить контрагента»
и окном редактирования карточки: сборка страницы и полей, сбор данных,
валидация с подсветкой ошибок, выбор при дубликате ИНН (п. 4.1), сохранение
через DataStore и редактирование существующего контрагента.

Этап M4 миграции Tk → PySide6: чтение/снимок/сбой полей и подстановка
прежних значений вынесены в GUI-free ``party_form_core`` — один источник
логики формы для обеих сборок; эта сборка делегирует (поведение и имена
без изменений, тесты без правок).

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

Ширина полей (``ENTRY_WIDTH``) и форматирование телефона при вводе
(``make_phone_formatter``) берутся напрямую из GUI-free модуля app_common
(этап M1 миграции Tk → PySide6 — поздние ``import main`` ликвидированы).
Остальные зависимости — модули services, storage, ui_macos,
document_builder — импортируются напрямую, как в предыдущих контроллерах.

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

import app_common
import dadata_controller
import services
import storage
import ui_macos as ui
from document_builder import F, digits_only
from party_form_core import (
    FORMAT_HINTS,
    READ_ERROR_MARKER,
    form_signature,
    merge_saved_values,
    read_entry_values,
    split_read_failures,
)
from services import validate_party_fields

# GUI-free ядро формы — один источник для обеих сборок (M4, аналог
# spec_table_core для таблицы): чтение/снимок/сбой/подстановка прежних
# значений живут в party_form_core; имена импортированы выше —
# cc.read_entry_values и соседние остаются доступны тестам, поведение
# не изменено. Tk EditCloseGuard остаётся здесь: он Tk-связан
# (window.destroy), Qt-окно правки имеет собственный guard.


class EditCloseGuard:
    """Подтверждение закрытия окна правки с несохранёнными изменениями (U8-MVP).

    Закрытие крестиком молча теряло изменения: базовая защита форм (п. 4.4)
    охватывает страницы, но не Toplevel. Guard снимает значения полей при
    открытии окна и перед закрытием сравнивает снимок с текущими: без
    изменений окно закрывается сразу (безопасно и для Escape), с
    изменениями — спрашивает подтверждение. Сбой чтения поля считается
    изменением (сравнение по маркеру, не по пустоте).
    """

    def __init__(self, window, entries, baseline, confirm_discard, log=None):
        self.window = window
        self.entries = entries
        self.baseline = dict(baseline or {})
        self._confirm_discard = confirm_discard
        self._log = log
        self._dialog_open = False

    def has_changes(self) -> bool:
        """Сравнение текущих значений с снимком открытия окна."""
        return form_signature(self.entries) != self.baseline

    def request_close(self) -> None:
        """Единый обработчик крестика и Escape: закрывает или спрашивает."""
        if not self.has_changes():
            self.window.destroy()
            return
        if self._dialog_open:
            return
        self._dialog_open = True
        try:
            discard = bool(self._confirm_discard())
        finally:
            self._dialog_open = False
        if discard:
            self.window.destroy()


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
        app = self.app
        error_var = self.error_var_for(mode)

        # -- Коды -------------------------------------------------------- #
        card.section("Коды")
        inn_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        inn = ui.entry(
            inn_holder, width=int(180 * app.scale),
            validatecommand=app._digits_vcmd_for(12),
            placeholder=FORMAT_HINTS[F.INN],
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
                           width=int(140 * app.scale), required=True, stretch=False,
                           placeholder=FORMAT_HINTS[F.KPP])
            self.add_field(card, mode, F.OGRN, F.OGRN, kind="digits", max_len=13,
                           width=int(180 * app.scale), required=True, stretch=False,
                           placeholder=FORMAT_HINTS[F.OGRN])
        else:
            self.add_field(card, mode, F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                           width=int(180 * app.scale), required=True, stretch=False,
                           placeholder=FORMAT_HINTS[F.OGRNIP])

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
                                    width=int(app_common.ENTRY_WIDTH * app.scale))
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
                       width=int(260 * app.scale), required=True,
                       placeholder=FORMAT_HINTS[F.ACCOUNT])
        self.add_field(card, mode, F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                       width=int(260 * app.scale), required=True,
                       placeholder=FORMAT_HINTS[F.CORR_ACCOUNT])
        self.add_field(card, mode, F.BANK, F.BANK, required=True)
        self.add_field(card, mode, F.BIK, F.BIK, kind="digits", max_len=9,
                       width=int(140 * app.scale), required=True, stretch=False,
                       placeholder=FORMAT_HINTS[F.BIK])

    def add_field(self, card, mode: str, key: str, label: str, kind="text",
                  width=None, required=False, stretch=True, max_len=None,
                  hint: str | None = None, placeholder: str | None = None):
        app = self.app
        width = width or int(app_common.ENTRY_WIDTH * app.scale)
        validate = None
        if kind == "digits":
            validate = app._digits_vcmd_for(max_len) if max_len else (app._digits_vcmd, "%P")
        elif kind == "phone":
            validate = (app._phone_vcmd, "%P")

        widget = ui.entry(card.body, width=width, validatecommand=validate,
                          placeholder=placeholder)
        if kind == "digits":
            # Валидатор отбрасывает вставку целиком, если в буфере есть
            # недопустимые символы («р/с 40702810…»). Вставляем только цифры.
            ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
        elif kind == "phone":
            phone_var = tk.StringVar()
            widget.entry.configure(textvariable=phone_var)
            phone_var.trace_add("write", app_common.make_phone_formatter(phone_var, widget.entry))
            ui.enable_sanitized_paste(
                widget.entry,
                lambda s: "".join(ch for ch in str(s) if ch.isdigit() or ch in "+-()/. "),
            )

        app._cp_entries[mode][key] = widget
        card.row(label, widget, stretch=stretch, required=required,
                 error_var=self.error_var_for(mode)(key), hint=hint)
        return widget

    def _highlight_page_errors(self, mode: str, errors: dict[str, str]) -> None:
        """Подсвечивает ошибочные поля страницы карточки (путь добавления).

        После фокуса прокручивает страницу к первому ошибочному полю (U5):
        оно может быть за нижней границей окна, и алерт сам по себе
        не показывает, какое поле подсвечено.
        """
        app = self.app
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
            # U5: прокрутка к ошибочному полю внутри страницы карточки.
            try:
                ui.scroll_to_widget(ui.containing_scroller(first), first)
            except Exception:
                pass

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
        # CHKSUM (этап 2): контрольные цифры для НОВОЙ карточки — все
        # реквизиты считаются новыми (previous=None); опечатка в
        # контрольной цифре ловится при вводе, а не у контрагента.
        for key, message in services.checksum_errors(data, is_ip, previous=None).items():
            errors.setdefault(key, message)
        if errors:
            self._highlight_page_errors(mode, errors)
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

        # U8-полностью (UX-B, этап 5): «Заполнить по ИНН» в окне правки.
        # Цель DaData — снимок адресата: поля окна (dict мутируется ниже
        # add_field — цель видит все поля), фиксированный режим карточки,
        # собственная строка статуса, диалоги модальны к окну правки.
        entries: dict[str, tk.Widget] = {}
        error_vars: dict[str, tk.StringVar] = {}
        edit_status_var = tk.StringVar(value="")
        edit_target = dadata_controller.DadataTarget(
            win, entries, mode, edit_status_var,
        )

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
        # U8: строка статуса DaData окна правки (анимация/результат).
        ctk.CTkLabel(
            header, textvariable=edit_status_var,
            font=ui.font(12), text_color=ui.COLORS["accent"], anchor="w", height=16,
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
                      stretch=True, max_len=None, hint=None, placeholder=None):
            width = width or int(app_common.ENTRY_WIDTH * app.scale)
            validate = None
            if kind == "digits":
                validate = app._digits_vcmd_for(max_len) if max_len else (app._digits_vcmd, "%P")
            elif kind == "phone":
                validate = (app._phone_vcmd, "%P")
            widget = ui.entry(card.body, width=width, validatecommand=validate,
                              placeholder=placeholder)
            if kind == "digits":
                # Валидатор отбрасывает вставку целиком, если в буфере есть
                # недопустимые символы. Вставляем только цифры.
                ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
            elif kind == "phone":
                phone_var = tk.StringVar()
                widget.entry.configure(textvariable=phone_var)
                phone_var.trace_add("write", app_common.make_phone_formatter(phone_var, widget.entry))
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
                  width=int(180 * app.scale), stretch=False, required=True,
                  placeholder=FORMAT_HINTS[F.INN])
        # U8-полностью: кнопка рядом с ИНН — как на странице добавления.
        # Диалоги и статус привязаны к окну правки (edit_target); фабрика
        # потока — по умолчанию threading.Thread, как на главной странице.
        fill_btn = ui.ghost_button(
            card.body, "Заполнить по ИНН",
            lambda: app._dadata_flow.fill_from_inn(target=edit_target),
            width=int(152 * app.scale), height=int(28 * app.scale),
        )
        card.row(
            "", fill_btn,
            hint="Кнопка подтягивает название, адрес, директора и банк "
                 "из ЕГРЮЛ/ЕГРИП (U8-полностью, этап 5).",
        )
        if is_ip:
            add_field(F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                      width=int(180 * app.scale), stretch=False, required=True,
                      placeholder=FORMAT_HINTS[F.OGRNIP])
        else:
            add_field(F.KPP, F.KPP, kind="digits", max_len=9,
                      width=int(140 * app.scale), stretch=False, required=True,
                      placeholder=FORMAT_HINTS[F.KPP])
            add_field(F.OGRN, F.OGRN, kind="digits", max_len=13,
                      width=int(180 * app.scale), stretch=False, required=True,
                      placeholder=FORMAT_HINTS[F.OGRN])

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
                                    width=int(app_common.ENTRY_WIDTH * app.scale))
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
                  width=int(260 * app.scale), required=True,
                  placeholder=FORMAT_HINTS[F.ACCOUNT])
        add_field(F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                  width=int(260 * app.scale), required=True,
                  placeholder=FORMAT_HINTS[F.CORR_ACCOUNT])
        add_field(F.BANK, F.BANK, required=True)
        add_field(F.BIK, F.BIK, kind="digits", max_len=9,
                  width=int(140 * app.scale), stretch=False, required=True,
                  placeholder=FORMAT_HINTS[F.BIK])

        # -- Защита карточки при закрытии (DEF-1: D4 + U8-MVP) -------------- #
        # Снимок полей при открытии: сравнение «до/текущие» решает, нужно ли
        # подтверждение при закрытии крестиком или Escape.
        baseline = form_signature(entries)

        def confirm_discard() -> bool:
            choice = app._choice_dialog(
                "Несохранённые изменения",
                "Изменения карточки ещё не сохранены.\n\n"
                "Закрыть окно без сохранения?",
                [
                    ("stay", "Вернуться к правке", True),
                    ("close", "Закрыть без сохранения", False),
                ],
            )
            return choice == "close"

        close_guard = EditCloseGuard(win, entries, baseline, confirm_discard, log=app._log)
        win.protocol("WM_DELETE_WINDOW", close_guard.request_close)
        win.bind("<Escape>", lambda _event: close_guard.request_close())

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
                # U5 (UX-A): прокрутка окна редактирования к ошибочному полю.
                try:
                    ui.scroll_to_widget(ui.containing_scroller(first), first)
                except Exception:
                    pass

        def save():
            # D4: сбой чтения виджета не превращается в пустую строку —
            # иначе «Сохранить» мог записать карточку с обнулёнными реквизитами.
            data, failed = read_entry_values(entries, log=app._log)
            if failed:
                required_keys = (
                    services.REQUIRED_PARTY_COMMON
                    + (services.REQUIRED_IP if is_ip else services.REQUIRED_OOO)
                )
                blocking, optional = split_read_failures(failed, required_keys)
                if blocking:
                    app._log.error(
                        "Сохранение карточки прервано: не удалось прочитать поля %s",
                        ", ".join(blocking),
                    )
                    ui.alert(
                        win, "Не удалось прочитать поля карточки",
                        "Сохранение прервано: часть полей карточки недоступна "
                        "(ошибка интерфейса).\n\nПоля: "
                        + ", ".join(blocking)
                        + "\n\nПереоткройте окно редактирования и попробуйте ещё раз. "
                        "Подробности — в журнале (%APPDATA%\\DogovorGenerator\\logs).",
                        "error",
                    )
                    return
                # Сбой только в необязательных полях: явный выбор пользователя —
                # молчаливой замены нет (частичный результат из карточки DEF-1).
                choice = app._choice_dialog(
                    "Поля недоступны",
                    "Не удалось прочитать необязательные поля: "
                    + ", ".join(optional) + ".\n\n"
                    "Сохранить карточку с прежними значениями этих полей?",
                    [
                        ("keep", "Сохранить с прежними значениями", True),
                        ("cancel", "Вернуться к правке", False),
                    ],
                )
                if choice != "keep":
                    return
                data = merge_saved_values(data, optional, saved)
            if is_ip:
                data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")

            errors = validate_party_fields(data, is_ip)
            # CHKSUM (этап 2): контрольные цифры — только для изменённых
            # реквизитов; старые карточки открываются и пересохраняются
            # без перепроверки (сверка с прежним значением поля).
            for key, message in services.checksum_errors(data, is_ip, previous=saved).items():
                errors.setdefault(key, message)
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
            footer, "Отмена", close_guard.request_close,
            width=int(140 * app.scale), height=int(30 * app.scale),
        ).pack(side="right")
        ui.primary_button(
            footer, "Сохранить", save,
            width=int(170 * app.scale), height=int(30 * app.scale),
        ).pack(side="right", padx=(0, int(10 * app.scale)))
