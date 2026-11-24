"""Страница «Добавить контрагента» и окно редактирования на Qt (этап M4).

Qt-перенос ``counterparty_controller.CounterpartyController`` (страница
добавления) и ``edit_counterparty`` (окно правки). Логика сервиса не
копируется: валидация — существующие ``services.validate_party_fields`` /
``services.checksum_errors`` (CHKSUM ИНН/ОГРН/счетов — общее ядро обеих
сборок), сохранение — ``store.add_counterparty`` / ``update_counterparty``;
чтение/снимок/сбой полей — GUI-free ``party_form_core`` (один источник
с Tk-сборкой, этап M4).

Контракты, переносимые дословно из Tk:

* состав полей и обязательность — без изменений (план M4 «не трогаем»);
  у ИП фактический адрес повторяет юридический;
* сбой чтения поля не превращается в пустую строку (D4): обязательное —
  блокирует сохранение с явным сообщением; необязательное — явный выбор
  «Сохранить с прежними значениями» (DEF-1, молчаливой замены нет);
* контрольные цифры: НОВАЯ карточка — все реквизиты новые (previous=None);
  окно правки — только изменённые (previous=saved): старые карточки
  открываются и пересохраняются без перепроверки;
* дубликаты ИНН: молча создавать вторую карточку нельзя (п. 4.1) — выбор
  «Открыть существующего / Создать всё равно / Отмена»; при правке ИНН
  другого контрагента — ошибка поля;
* сбой записи реестра не оставляет «фантом» в памяти (п. 9.2);
* «Заполнить по ИНН» в обоих окнах (U8-полностью) — снимок адресата,
  guard busy, тихий отказ при закрытом окне (dadata_qt);
* окно правки: изменения только карточки — ранее созданные документы
  хранят собственную копию реквизитов; закрытие с изменениями —
  подтверждение (DEF-1), без изменений — сразу; Escape — тот же guard;
* живой телефон-маска «+7 (999) 123-45-67» и вставка с очисткой
  (аудит 2.0.1): из буфера — только допустимые символы.

Отличия от Tk-сборки (зафиксированы в docs/qt-migration-stage-notes.md):
после добавления каркас открывает «Спецификацию» (ближайшая живая
страница; в Tk уход на «Договор» — плейсхолдер до M6); кнопка
«Редактировать карточку» на панели «Документ для» страницы «Спецификация»
— временная точка входа в окно правки до M5 (реестр с панелью действий);
футерная кнопка Tk «Сохранить контрагента (Ctrl+S)» заменена кнопкой на
самой странице (U14-суффикс сохранён).

Состояние страницы — у контроллера (раздел Д плана); единый текущий
контрагент — на каркасе (``current_party_id``). Каркас получает
``build_page`` (вызвана в конструкторе), ``is_form_dirty`` (п. 4.4) и
``save`` для Ctrl+S.
"""

from __future__ import annotations

from PySide6.QtCore import QRegularExpression
from PySide6.QtGui import QRegularExpressionValidator
from PySide6.QtWidgets import (
    QAbstractButton, QButtonGroup, QComboBox, QDialog, QFrame, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QScrollArea, QStackedWidget, QVBoxLayout,
    QWidget,
)

import dadata_qt
import services
import storage
from document_builder import F, digits_only, format_phone
from party_form_core import (
    FORMAT_HINTS, PHONE_ALLOWED_PATTERN, form_signature, merge_saved_values,
    read_entry_values, split_read_failures,
)

PAGE_TITLE = "Добавить контрагента"
PAGE_SUBTITLE = "Заполните карточку — она сохранится в реестре программы."
TYPE_LABELS = ("Юр. лицо", "ИП")

# Кнопка «Заполнить по ИНН»: подписи подсказок (Tk-паритет).
DADATA_PAGE_HINT = (
    "Кнопка подтягивает название, адрес, директора и банк из ЕГРЮЛ/ЕГРИП."
)
DADATA_EDIT_HINT = (
    "Кнопка подтягивает название, адрес, директора и банк "
    "из ЕГРЮЛ/ЕГРИП (U8-полностью, этап 5)."
)

# Хвосты сообщений (Tk-паритет counterparty_controller).
LOGS_HINT = "Подробности — в журнале (%APPDATA%\\DogovorGenerator\\logs)."


def sanitize_phone_paste(text: str) -> str:
    """Вставка телефона: только допустимые символы (Tk-паритет, аудит 2.0.1)."""
    return "".join(ch for ch in str(text) if ch.isdigit() or ch in "+-()/. ")


def apply_phone_mask(edit) -> None:
    """Переформатирует ввод телефона в маску +7 (999) 123-45-67 (Tk-паритет).

    Алгоритм — app_common.make_phone_formatter; курсор сдвигается с
    поправкой на изменение длины (аналог restore_caret_after_reformat).
    Повторный вызов на переформатированной строке — no-op (рекурсии нет).
    """
    raw = edit.text()
    digits = digits_only(raw)
    if digits.startswith("8"):
        digits = "7" + digits[1:]
    if digits and not digits.startswith("7"):
        digits = "7" + digits
    formatted = format_phone(digits)
    if raw == formatted:
        return
    caret = edit.cursorPosition()
    edit.setText(formatted)
    edit.setCursorPosition(max(0, min(caret + len(formatted) - len(raw), len(formatted))))


class QtEntry:
    """Адаптер QLineEdit под протокол GUI-free ядра формы (party_form_core).

    Ядро читает виджеты через ``.get()`` (Tk-протокол) — один источник
    чтения/снимка для обеих сборок. Уничтоженный C++-объект Qt поднимает
    RuntimeError — ядро классифицирует это как сбой чтения (не пустоту).
    """

    def __init__(self, edit: QLineEdit) -> None:
        self.edit = edit

    def get(self) -> str:
        return self.edit.text()

    def set_text(self, value: str) -> None:
        self.edit.setText(value)

    def set_error(self, on: bool) -> None:
        self.edit.setProperty("errorState", bool(on))
        style = self.edit.style()
        style.unpolish(self.edit)
        style.polish(self.edit)


class QtOptionEntry:
    """Адаптер QComboBox под протокол ядра (Tk ui.option_menu «Основание»)."""

    def __init__(self, combo: QComboBox) -> None:
        self.combo = combo

    def get(self) -> str:
        return self.combo.currentText()

    def set_text(self, value: str) -> None:
        index = self.combo.findText(value)
        if index >= 0:
            self.combo.setCurrentIndex(index)

    def set_error(self, on: bool) -> None:
        self.combo.setProperty("errorState", bool(on))
        style = self.combo.style()
        style.unpolish(self.combo)
        style.polish(self.combo)


class PartyLineEdit(QLineEdit):
    """Поле карточки: вставка с очисткой — из буфера только допустимые символы.

    Контракт Tk enable_sanitized_paste (аудит 2.0.1): стандартная вставка
    с валидатором молча отбрасывается целиком, если в буфере есть
    недопустимые символы («р/с 40702810…») — вместо отклонения вставляем
    очищенное значение. Максимальная длина — через setMaxLength (контракт
    Tk validatecommand «не длиннее max_len»).
    """

    def __init__(self, sanitizer=None, parent=None) -> None:
        super().__init__(parent)
        self._sanitizer = sanitizer

    def insertFromMimeData(self, source) -> None:  # noqa: N802 — имя Qt
        text = (source.text() or "") if source is not None else ""
        if self._sanitizer is not None and text:
            cleaned = self._sanitizer(text)
            if cleaned:
                self.insert(cleaned)
            return
        super().insertFromMimeData(source)


class PartyFormBuilder:
    """Собирает секции/поля карточки — общий код страницы и окна правки.

    Один состав полей для добавления и редактирования (Tk-паритет
    build_cp_fields / edit_counterparty); ошибки — подписи
    services.validate_party_fields, подсветка через errorState/QSS.
    """

    def __init__(self, app, entries: dict, error_labels: dict):
        self.app = app
        self.entries = entries
        self.error_labels = error_labels

    # ------------------------------------------------------------------ #
    #  Примитивы
    # ------------------------------------------------------------------ #
    def _section(self, layout: QVBoxLayout, caption: str) -> None:
        label = QLabel(caption, layout.parentWidget())
        label.setObjectName("cardCaption")
        layout.addWidget(label)

    def _error_label(self, key: str, parent: QWidget) -> QLabel:
        label = QLabel("", parent)
        label.setObjectName("fieldError")
        label.setWordWrap(True)
        self.error_labels[key] = label
        return label

    def _add_field(self, layout: QVBoxLayout, key: str, label: str, kind="text",
                   max_len: int | None = None, required: bool = False,
                   hint: str | None = None, width: int | None = None) -> QLineEdit:
        """Поле ввода + строка подписи + ошибка (Tk card.row-паритет)."""
        sanitizer = None
        if kind == "digits":
            sanitizer = digits_only
        elif kind == "phone":
            sanitizer = sanitize_phone_paste

        edit = PartyLineEdit(sanitizer, layout.parentWidget())
        if kind == "digits":
            pattern = f"^\\d{{0,{max_len}}}$" if max_len else "^\\d*$"
            edit.setValidator(QRegularExpressionValidator(
                QRegularExpression(pattern), edit))
            if max_len:
                edit.setMaxLength(max_len)
        elif kind == "phone":
            edit.setValidator(QRegularExpressionValidator(
                QRegularExpression(PHONE_ALLOWED_PATTERN), edit))
            edit.textChanged.connect(lambda _text, e=edit: apply_phone_mask(e))
        if key in FORMAT_HINTS:
            edit.setPlaceholderText(FORMAT_HINTS[key])

        if required:
            label = label + " *"
        row = QHBoxLayout()
        caption = QLabel(label, layout.parentWidget())
        row.addWidget(caption)
        row.addStretch(1)
        layout.addLayout(row)
        if width:
            edit.setFixedWidth(width)
        layout.addWidget(edit)
        layout.addWidget(self._error_label(key, layout.parentWidget()))
        if hint:
            hint_label = QLabel(hint, layout.parentWidget())
            hint_label.setObjectName("fieldHint")
            hint_label.setWordWrap(True)
            layout.addWidget(hint_label)
        self.entries[key] = QtEntry(edit)
        return edit

    def _add_option(self, layout: QVBoxLayout, key: str, label: str,
                    values: list[str], prefill: str | None = None) -> None:
        combo = QComboBox(layout.parentWidget())
        combo.addItems(values)
        if prefill in values:
            combo.setCurrentText(prefill)
        row = QHBoxLayout()
        caption = QLabel(label, layout.parentWidget())
        row.addWidget(caption)
        row.addStretch(1)
        layout.addLayout(row)
        layout.addWidget(combo)
        self.entries[key] = QtOptionEntry(combo)

    def _add_fact_address_note(self, layout: QVBoxLayout) -> None:
        """ИП: фактический адрес = юридический (подпись без поля, Tk-паритет)."""
        row = QHBoxLayout()
        caption = QLabel(F.FACT_ADDRESS, layout.parentWidget())
        note = QLabel("как юридический", layout.parentWidget())
        note.setObjectName("fieldHint")
        row.addWidget(caption)
        row.addWidget(note, 1)
        layout.addLayout(row)

    def _add_inn_row(self, layout: QVBoxLayout, mode: str, prefill: str | None,
                     on_dadata, hint: str) -> None:
        edit = self._add_field(
            layout, F.INN, F.INN, kind="digits", max_len=12, required=True,
            width=180, hint=hint,
        )
        if prefill:
            edit.setText(str(prefill))
        button = QPushButton("Заполнить по ИНН", layout.parentWidget())
        button.clicked.connect(lambda _checked=False: on_dadata())
        layout.addWidget(button)

    # ------------------------------------------------------------------ #
    #  Сборка карточки
    # ------------------------------------------------------------------ #
    def build(self, layout: QVBoxLayout, mode: str,
              prefill: dict | None = None, on_dadata=None,
              dadata_hint: str = DADATA_PAGE_HINT) -> None:
        """Секции карточки для режима ``mode`` (Tk-паритет состава полей)."""
        is_ip = mode == "ip"
        saved = prefill or {}

        # -- Коды -------------------------------------------------------- #
        self._section(layout, "Коды")
        self._add_inn_row(
            layout, mode, saved.get(F.INN), on_dadata, dadata_hint,
        )
        if is_ip:
            self._add_field(layout, F.OGRNIP, F.OGRNIP, kind="digits",
                            max_len=15, required=True, width=180)
        else:
            self._add_field(layout, F.KPP, F.KPP, kind="digits",
                            max_len=9, required=True, width=140)
            self._add_field(layout, F.OGRN, F.OGRN, kind="digits",
                            max_len=13, required=True, width=180)

        # -- Организация / Предприниматель -------------------------------- #
        self._section(layout, "Предприниматель" if is_ip else "Организация")
        if is_ip:
            self._add_field(layout, F.IP_NAME, F.IP_NAME, required=True)
        else:
            self._add_field(layout, F.FULL_NAME, "Полное название", required=True)
            self._add_field(layout, F.SHORT_NAME, "Сокращённое название", required=True)

        # -- Руководитель ------------------------------------------------- #
        if not is_ip:
            self._section(layout, "Руководитель")
            self._add_field(layout, F.DIRECTOR, F.DIRECTOR, required=True)
            self._add_option(
                layout, F.BASIS, F.BASIS, ["Устав", "Доверенность"],
                prefill=str(saved.get(F.BASIS, "")).strip() or None,
            )

        # -- Адреса ------------------------------------------------------- #
        self._section(layout, "Адреса")
        self._add_field(layout, F.LEGAL_ADDRESS, F.LEGAL_ADDRESS, required=True)
        if is_ip:
            self._add_fact_address_note(layout)
        else:
            self._add_field(layout, F.FACT_ADDRESS, F.FACT_ADDRESS)

        # -- Контакты ----------------------------------------------------- #
        # Email и телефон необязательны — юридически значимые поля выше (п. 4.5).
        self._section(layout, "Контакты")
        self._add_field(layout, F.EMAIL, F.EMAIL, hint="Необязательно.")
        self._add_field(layout, F.PHONE, F.PHONE, kind="phone", hint="Необязательно.")

        # -- Банк ---------------------------------------------------------- #
        self._section(layout, "Банковские реквизиты")
        self._add_field(layout, F.ACCOUNT, F.ACCOUNT, kind="digits", max_len=20,
                        required=True, width=260)
        self._add_field(layout, F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits",
                        max_len=20, required=True, width=260)
        self._add_field(layout, F.BANK, F.BANK, required=True)
        self._add_field(layout, F.BIK, F.BIK, kind="digits", max_len=9,
                        required=True, width=140)

        # Prefill оставшихся текстовых полей (правка: значения сохранённой
        # карточки; добавление — пусто).
        for key, adapter in self.entries.items():
            value = str(saved.get(key, "")).strip()
            if not value:
                continue
            if key not in (F.INN,) and hasattr(adapter, "set_text"):
                adapter.set_text(value)


class CounterpartyPageController:
    """Страница «Добавить контрагента»: поля, валидация, дубликат ИНН, сохранение."""

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        self.page: QWidget | None = None
        self.entries: dict[str, dict] = {"ooo": {}, "ip": {}}
        self.error_labels: dict[str, dict] = {"ooo": {}, "ip": {}}
        self.baseline: tuple | None = None
        self._build_page()
        self.reset_form_baseline()

    # ------------------------------------------------------------------ #
    #  Построение страницы
    # ------------------------------------------------------------------ #
    def _build_page(self) -> None:
        app = self.app
        page = QWidget()
        self.page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(24, 18, 24, 12)
        outer.setSpacing(6)

        title = QLabel(PAGE_TITLE, page)
        title.setObjectName("pageTitle")
        subtitle = QLabel(PAGE_SUBTITLE, page)
        subtitle.setObjectName("pageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        # Переключатель типа карточки (Tk ui.segmented: «Юр. лицо»/«ИП»).
        type_row = QHBoxLayout()
        self.type_buttons: dict[str, QAbstractButton] = {}
        group = QButtonGroup(page)
        group.setExclusive(True)
        for label in TYPE_LABELS:
            button = QPushButton(label, page)
            button.setCheckable(True)
            button.setProperty("variant", "segment")
            button.clicked.connect(
                lambda _checked=False, value=label: self.on_type_change(value)
            )
            group.addButton(button)
            type_row.addWidget(button)
            self.type_buttons[label] = button
        self.type_buttons["Юр. лицо"].setChecked(True)
        type_row.addStretch(1)
        outer.addLayout(type_row)

        # Строка статуса «Заполнить по ИНН» (Tk _cp_status_var).
        self.status_label = QLabel("", page)
        self.status_label.setObjectName("dadataStatus")
        outer.addWidget(self.status_label)

        # Прокручиваемая зона с двумя карточками (ООО/ИП) — Tk CTkScrollableFrame.
        scroll = QScrollArea(page)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        holder = QWidget(scroll)
        holder_layout = QHBoxLayout(holder)
        holder_layout.setContentsMargins(0, 4, 8, 0)
        self.stack = QStackedWidget(holder)
        holder_layout.addWidget(self.stack, 1)
        for mode in ("ooo", "ip"):
            frame = QWidget(self.stack)
            frame_layout = QVBoxLayout(frame)
            frame_layout.setContentsMargins(0, 0, 0, 0)
            frame_layout.setSpacing(4)
            card = QFrame(frame)
            card.setObjectName("card")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(14, 12, 14, 14)
            card_layout.setSpacing(4)
            builder = PartyFormBuilder(app, self.entries[mode], self.error_labels[mode])
            builder.build(
                card_layout, mode,
                on_dadata=lambda m=mode: self._fill_from_inn_for_mode(m),
            )
            frame_layout.addWidget(card)
            frame_layout.addStretch(1)
            self.stack.addWidget(frame)
        scroll.setWidget(holder)
        outer.addWidget(scroll, 1)

        # Кнопка сохранения на странице (Tk-футер «Сохранить контрагента
        # (Ctrl+S)», U14-суффикс); Ctrl+S ведёт каркас — save_current_page.
        footer = QHBoxLayout()
        footer.addStretch(1)
        save_button = QPushButton("Сохранить контрагента (Ctrl+S)", page)
        save_button.setProperty("variant", "primary")
        save_button.clicked.connect(self.save_card)
        footer.addWidget(save_button)
        outer.addLayout(footer)

        self._install_page()

    def _install_page(self) -> None:
        """Заменяет плейсхолдер M2 живой страницей «Добавить контрагента»."""
        app = self.app
        placeholder = app.pages.get("counterparty")
        index = app.stack.indexOf(placeholder) if placeholder is not None else -1
        if placeholder is not None:
            app.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        app.stack.insertWidget(max(0, index), self.page)
        app.pages["counterparty"] = self.page
        if app.current_page == "counterparty":
            app.stack.setCurrentWidget(self.page)

    # ------------------------------------------------------------------ #
    #  Тип карточки и цель DaData
    # ------------------------------------------------------------------ #
    def current_mode(self) -> str:
        """Активный режим карточки по кнопке-переключателю (Tk _cp_type_var)."""
        for label, button in self.type_buttons.items():
            if button.isChecked():
                return "ip" if label == "ИП" else "ooo"
        return "ooo"

    def on_type_change(self, value: str) -> None:
        """Переключение ООО/ИП: показывается своя карточка (Tk tkraise)."""
        mode = "ip" if value == "ИП" else "ooo"
        if label_checked := ("ИП" if mode == "ip" else "Юр. лицо"):
            button = self.type_buttons.get(label_checked)
            if button is not None and not button.isChecked():
                button.setChecked(True)
        self.stack.setCurrentWidget(self.stack.widget(0 if mode == "ooo" else 1))

    def dadata_target(self, mode: str | None = None) -> dadata_qt.DadataTargetQt:
        """Снимок адресата «Заполнить по ИНН» страницы (аналог from_app)."""
        mode = mode or self.current_mode()
        return dadata_qt.DadataTargetQt(
            self.app.window, self.entries[mode], mode,
            self.status_label.setText,
        )

    def _fill_from_inn_for_mode(self, mode: str) -> None:
        """Кнопка «Заполнить по ИНН»: цель строго по нажатой карточке.

        Tk-контракт «переменная обязательна» (баг 2.0.0: без привязки
        тип всегда «Юр. лицо» и поиск по ИНН у ИП читал пустое поле):
        цель строится для конкретного видимого режима, а не по
        переключателю.
        """
        self.app.dadata_flow.fill_from_inn(target=self.dadata_target(mode))

    # ------------------------------------------------------------------ #
    #  Сбор данных и сохранение (Tk-паритет collect/save_card)
    # ------------------------------------------------------------------ #
    def collect(self, mode: str) -> dict | None:
        """Собирает данные карточки; сбой чтения — None (D4, Tk-паритет).

        Аудит 2.0.1: ошибка чтения виджета не превращается в пустую
        строку — вызывающий код обязан сообщить пользователю вместо
        тихой записи пустых значений.
        """
        data = {}
        for key, adapter in self.entries[mode].items():
            try:
                data[key] = adapter.get().strip()
            except Exception as exc:
                self.log.warning(
                    "Не удалось прочитать поле карточки «%s»: %s",
                    key, type(exc).__name__,
                )
                return None
        if mode == "ip":
            data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")
        return data

    def _highlight_errors(self, mode: str, errors: dict[str, str]) -> None:
        """Подсвечивает ошибочные поля и переводит фокус на первое (U5)."""
        for label in self.error_labels[mode].values():
            self._set_error_label(label, "")
        for adapter in self.entries[mode].values():
            adapter.set_error(False)
        for key, message in errors.items():
            label = self.error_labels[mode].get(key)
            if label is not None:
                self._set_error_label(label, message)
            adapter = self.entries[mode].get(key)
            if adapter is not None:
                adapter.set_error(True)
        first = self.entries[mode].get(next(iter(errors), None))
        if first is not None and hasattr(first, "edit"):
            try:
                first.edit.setFocus()
            except RuntimeError:
                pass

    @staticmethod
    def _set_error_label(label: QLabel, message: str) -> None:
        label.setText(message)
        label.setObjectName("fieldErrorActive" if message else "fieldError")
        style = label.style()
        style.unpolish(label)
        style.polish(label)

    def clear_error_highlights(self, mode: str) -> None:
        """Сбрасывает подсветку (повторная попытка сохранения)."""
        self._highlight_errors(mode, {})

    def save_card(self) -> None:
        """Сохранение новой карточки — Tk-паритет save_card."""
        app = self.app
        mode = self.current_mode()
        is_ip = mode == "ip"
        data = self.collect(mode)
        if data is None:
            # Сбор данных сломался — тихо сохранять «пустую» карточку нельзя.
            app.alert(
                app.window, "Не удалось прочитать поля формы",
                "Часть полей карточки недоступна (ошибка интерфейса).\n\n"
                "Переоткройте страницу «Добавить контрагента» и попробуйте "
                "ещё раз. " + LOGS_HINT, "error",
            )
            return

        errors = services.validate_party_fields(data, is_ip)
        # CHKSUM (этап 2): контрольные цифры для НОВОЙ карточки — все
        # реквизиты считаются новыми (previous=None).
        for key, message in services.checksum_errors(data, is_ip, previous=None).items():
            errors.setdefault(key, message)
        if errors:
            self._highlight_errors(mode, errors)
            app.alert(
                app.window, "Ошибки в данных",
                "Проверьте подсвеченные поля:\n\n"
                + "\n".join(f"• {error}" for error in errors.values()),
                "error",
            )
            return

        # Дубликаты по ИНН: молча создавать вторую карточку нельзя (п. 4.1).
        inn = data.get(F.INN, "").strip()
        duplicate = app.store.find_by_inn(inn)
        if duplicate is not None:
            choice = app.choice(
                app.window, "Контрагент с таким ИНН уже есть",
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
                    self.reset_form_baseline()
                    app.set_current_party(duplicate)
                    app.select_page("spec")
                return

        rec = app.store.add_counterparty(mode, data)
        if app.store.last_error:
            # Не оставляем «фантом» в памяти, если запись не легла на диск.
            try:
                app.store.counterparties.remove(rec)
            except ValueError:
                pass
            app.alert(
                app.window, "Не удалось сохранить контрагента",
                app.store.last_error, "error",
            )
            return
        self.reset_form_baseline()
        app.refresh_after_party_change()
        app.set_current_party(rec)
        app.notify(f"{storage.display_name(rec)} добавлен в реестр программы.")
        # Tk уходит на «Договор» с выбранным контрагентом; в Qt-сборке
        # «Договор» — плейсхолдер до M6, ближайшая живая страница —
        # «Спецификация» (решение зафиксировано в заметках этапа, M6
        # вернёт уход на договор).
        app.select_page("spec")

    # ------------------------------------------------------------------ #
    #  Защита несохранённых данных (п. 4.4; Tk document_context_controller)
    # ------------------------------------------------------------------ #
    def form_signature(self) -> tuple | None:
        """Данные формы, потеря которых требует подтверждения.

        Tk-контракт: подпись = все поля активного режима (sorted items);
        сбой чтения — None, страница считается «грязной» (аудит 2.0.1).
        Переключение типа ООО↔ИП меняет подпись — как в Tk.
        """
        try:
            data = self.collect(self.current_mode())
            if data is None:
                return None
            return tuple(sorted(data.items()))
        except Exception as exc:
            self.log.warning(
                "Не удалось снять подпись формы «counterparty»: %s",
                type(exc).__name__,
            )
            return None

    def is_form_dirty(self) -> bool:
        """Сравнение подписи с baseline; сбой чтения (None) — «грязная»."""
        return self.form_signature() != self.baseline

    def reset_form_baseline(self) -> None:
        """Baseline страницы пересчитан (построение, после сохранения)."""
        self.baseline = self.form_signature()

    def save(self) -> None:
        """Точка входа Ctrl+S каркаса (Tk FORM_PAGES: save_card)."""
        self.save_card()


class EditCounterpartyDialog(QDialog):
    """Окно редактирования контрагента (DEF-1: защита несохранённых, Escape).

    Изменяется только карточка в реестре; ранее созданные документы не
    затрагиваются (каждый документ хранит собственную копию реквизитов
    на момент создания). Закрытие с изменениями — подтверждение
    (деструктивная семантика: Enter — «Вернуться к правке», U7); без
    изменений окно закрывается сразу; Escape — тот же guard.
    """

    def __init__(self, app, rec: dict) -> None:
        super().__init__(app.window)
        self.app = app
        self.log = app.log
        self.rec = rec
        self.mode = "ip" if rec.get("type") == "ip" else "ooo"
        self.is_ip = self.mode == "ip"
        self.saved = rec.get("fields", {})
        self.entries: dict = {}
        self.error_labels: dict = {}
        self.baseline: dict = {}
        self._dialog_open = False
        self._closing = False

        self.setWindowTitle(
            f"Редактирование контрагента — {storage.display_name(rec)}"
        )
        self.resize(760, 680)
        self.setMinimumSize(660, 520)
        self.setModal(True)

        self._build()
        # Снимок полей при открытии: сравнение «до/текущие» решает, нужно
        # ли подтверждение при закрытии крестиком или Escape (DEF-1).
        self.baseline = form_signature(self.entries)

    # ------------------------------------------------------------------ #
    #  Построение
    # ------------------------------------------------------------------ #
    def _build(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 18, 24, 12)
        outer.setSpacing(6)

        title = QLabel(
            f"Редактирование контрагента · "
            f"{'ИП' if self.is_ip else 'Юр. лицо'}", self,
        )
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            "Изменится только карточка в реестре. "
            "Ранее созданные документы останутся без изменений.", self,
        )
        subtitle.setObjectName("pageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        # U8: строка статуса DaData окна правки (анимация/результат).
        self.status_label = QLabel("", self)
        self.status_label.setObjectName("dadataStatus")
        outer.addWidget(self.status_label)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        card = QFrame(scroll)
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(14, 12, 14, 14)
        card_layout.setSpacing(4)
        builder = PartyFormBuilder(self.app, self.entries, self.error_labels)
        builder.build(
            card_layout, self.mode,
            prefill=self.saved,
            on_dadata=lambda: self.app.dadata_flow.fill_from_inn(
                target=self.dadata_target()
            ),
            dadata_hint=DADATA_EDIT_HINT,
        )
        scroll.setWidget(card)
        outer.addWidget(scroll, 1)

        footer = QHBoxLayout()
        footer.addStretch(1)
        cancel_button = QPushButton("Отмена", self)
        cancel_button.clicked.connect(self.request_close)
        footer.addWidget(cancel_button)
        save_button = QPushButton("Сохранить", self)
        save_button.setProperty("variant", "primary")
        save_button.clicked.connect(self.save)
        footer.addWidget(save_button)
        outer.addLayout(footer)

    def dadata_target(self) -> dadata_qt.DadataTargetQt:
        """Снимок адресата окна правки: свои поля, статус и родитель диалогов."""
        return dadata_qt.DadataTargetQt(
            self, self.entries, self.mode, self.status_label.setText,
        )

    # ------------------------------------------------------------------ #
    #  Закрытие: guard DEF-1 (крестик, Escape, «Отмена»)
    # ------------------------------------------------------------------ #
    def has_changes(self) -> bool:
        """Сравнение текущих значений со снимком открытия окна."""
        return form_signature(self.entries) != self.baseline

    def reject(self) -> None:  # noqa: N802 — имя Qt (Escape и крестик)
        self.request_close()

    def request_close(self) -> None:
        """Единый обработчик крестика, Escape и «Отмены» (Tk EditCloseGuard)."""
        if self._dialog_open or self._closing:
            return
        if not self.has_changes():
            self._really_close()
            return
        self._dialog_open = True
        try:
            discard = self.app.confirm(
                self, "Несохранённые изменения",
                "Изменения карточки ещё не сохранены.\n\n"
                "Закрыть окно без сохранения?",
                ok_text="Закрыть без сохранения",
                cancel_text="Вернуться к правке",
                destructive=True,
            )
        finally:
            self._dialog_open = False
        if discard:
            self._really_close()

    def _really_close(self) -> None:
        self._closing = True
        super().reject()

    # ------------------------------------------------------------------ #
    #  Подсветка ошибок
    # ------------------------------------------------------------------ #
    @staticmethod
    def _set_error_label(label: QLabel, message: str) -> None:
        label.setText(message)
        label.setObjectName("fieldErrorActive" if message else "fieldError")
        style = label.style()
        style.unpolish(label)
        style.polish(label)

    def _highlight_errors(self, errors: dict[str, str]) -> None:
        for label in self.error_labels.values():
            self._set_error_label(label, "")
        for adapter in self.entries.values():
            adapter.set_error(False)
        for key, message in errors.items():
            label = self.error_labels.get(key)
            if label is not None:
                self._set_error_label(label, message)
            adapter = self.entries.get(key)
            if adapter is not None:
                adapter.set_error(True)
        first = self.entries.get(next(iter(errors), None))
        if first is not None and hasattr(first, "edit"):
            try:
                first.edit.setFocus()
            except RuntimeError:
                pass

    # ------------------------------------------------------------------ #
    #  Сохранение (Tk-паритет save в edit_counterparty)
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        app = self.app
        # D4: сбой чтения виджета не превращается в пустую строку —
        # иначе «Сохранить» мог записать карточку с обнулёнными реквизитами.
        data, failed = read_entry_values(self.entries, log=app.log)
        if failed:
            required_keys = (
                services.REQUIRED_PARTY_COMMON
                + (services.REQUIRED_IP if self.is_ip else services.REQUIRED_OOO)
            )
            blocking, optional = split_read_failures(failed, required_keys)
            if blocking:
                app.log.error(
                    "Сохранение карточки прервано: не удалось прочитать поля %s",
                    ", ".join(blocking),
                )
                app.alert(
                    self, "Не удалось прочитать поля карточки",
                    "Сохранение прервано: часть полей карточки недоступна "
                    "(ошибка интерфейса).\n\nПоля: "
                    + ", ".join(blocking)
                    + f"\n\nПереоткройте окно редактирования и попробуйте "
                      f"ещё раз. {LOGS_HINT}",
                    "error",
                )
                return
            # Сбой только в необязательных полях: явный выбор пользователя —
            # молчаливой замены нет (DEF-1).
            keep = app.confirm(
                self, "Поля недоступны",
                "Не удалось прочитать необязательные поля: "
                + ", ".join(optional) + ".\n\n"
                "Сохранить карточку с прежними значениями этих полей?",
                ok_text="Сохранить с прежними значениями",
                cancel_text="Вернуться к правке",
            )
            if not keep:
                return
            data = merge_saved_values(data, optional, self.saved)
        if self.is_ip:
            data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")

        errors = services.validate_party_fields(data, self.is_ip)
        # CHKSUM: контрольные цифры — только для изменённых реквизитов;
        # старые карточки открываются и пересохраняются без перепроверки.
        for key, message in services.checksum_errors(data, self.is_ip, previous=self.saved).items():
            errors.setdefault(key, message)
        # ИНН должен оставаться уникальным в реестре.
        duplicate = app.store.find_by_inn(data.get(F.INN, "").strip())
        if duplicate is not None and duplicate.get("id") != self.rec.get("id"):
            errors.setdefault(F.INN, "Контрагент с таким ИНН уже есть в реестре")
        if errors:
            self._highlight_errors(errors)
            app.alert(
                self, "Ошибки в данных",
                "Проверьте подсвеченные поля:\n\n"
                + "\n".join(f"• {message}" for message in errors.values()),
                "error",
            )
            return

        # Обновляем только карточку контрагента. Документы не трогаем:
        # они хранят собственную копию реквизитов на момент создания.
        updated = app.store.update_counterparty(self.rec["id"], fields=data)
        if app.store.last_error:
            app.alert(
                self, "Не удалось сохранить изменения",
                app.store.last_error, "error",
            )
            return
        name = storage.display_name(updated or self.rec)
        # Переименование не должно «терять» выбранного контрагента (п. 4.2).
        if app.current_party_id == self.rec.get("id"):
            app.set_current_party(updated or self.rec)
        app.refresh_after_party_change()
        self._closing = True
        self.accept()
        app.notify(
            f"Карточка «{name}» обновлена. "
            "Ранее созданные документы не изменялись."
        )


def open_edit_dialog(app, rec: dict) -> EditCounterpartyDialog | None:
    """Точка входа «Редактировать карточку» (spec-панель до M5, реестр — M5)."""
    if rec is None:
        return None
    dialog = EditCounterpartyDialog(app, rec)
    dialog.show()
    return dialog
