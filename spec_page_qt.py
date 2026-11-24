"""Страница «Спецификация» на Qt (этап M3 миграции, контрольная точка CP1).

Qt-перенос Tk-страницы document_save_controller.DocumentSaveController
(только спецификация; сценарии договора/доп. соглашения — M6) и нужной
части DocumentContextController (выбор контрагента/договора, подпись
формы). Логика сервисного слоя не копируется: сохранение идёт через
существующий ``services.save_document`` (транзакция DEF-4, предохранитель
DEF-2), суммы — через ``document_builder.spec_total`` /
``spec_vat_amount_from_items`` (U12: строка итогов, предпросмотр и
печатная форма не могут разойтись по построению).

Контракты, переносимые дословно из Tk:

* **Контекст** (п. 4.2): список контрагентов с «· ID N» при совпадении
  названий; устаревший выбор честно очищается; договор-основание
  подставляется автоматически и показывается явно.
* **Подпись формы** (п. 4.4): выбор контрагента/договора не входит
  (переносится между страницами); при сбое чтения подпись деградирует
  в None — страница считается «грязной» и просит подтверждение.
* **Отложенный предпросмотр** (п. 5.6, UPD-UI P-07): пересборка через
  120 мс после последнего изменения и только при видимой странице;
  перед сохранением дособирается синхронно.
* **Предпросмотр = печатная форма** (раздел Е CP1): текст собирается
  из шаблона ``spec_template*.docx`` через ``services.build_record_docx``
  (путь сохранения) БЕЗ записи на диск; остаточные ``{{…}}`` видны в
  предпросмотре до попытки сохранения (DEF-2).
* **Сохранение** (п. 4.1, 9.2): валидации, сводка позиций (U6),
  подтверждение повтора номера, честные сообщения отката, тост с
  «Открыть»/«Показать в папке», обновление статуса.

Состояние страницы — у контроллера (план, раздел Д), а не на каркасе;
текущий контрагент — на каркасе (общий контекст будущих страниц, как
``_current_party_id`` в Tk). Каркас получает точку входа ``build_page``
(вызвана в конструкторе) и метод ``save`` для Ctrl+S.
"""

from __future__ import annotations

import os
from datetime import datetime

from PySide6.QtCore import Qt, QRegularExpression, QTimer
from PySide6.QtGui import QRegularExpressionValidator
from PySide6.QtWidgets import (
    QComboBox, QFrame, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QPushButton, QScrollArea, QSplitter, QVBoxLayout, QWidget,
)

import app_common
import counterparty_page_qt
import services
import shell_open
import storage
from document_builder import (
    F, MONTHS_GENITIVE, digits_only, docx_to_text,
    prices_valid_until_default, spec_total,
    unfilled_placeholders,
)
from services import safe_filename

from spec_table_qt import SpecTableQt

# Отложенный предпросмотр (п. 5.6): как у договора и доп. соглашения.
PREVIEW_DELAY_MS = 120

# Поиск по ИНН — только цифры, не длиннее 12 (Tk: app._digits_vcmd_for(12)).
_INN_DIGITS_RE = QRegularExpression("^\\d{0,12}$")

# U6 (UX-A): сводка ошибок позиций — тот же лимит строк, что в Tk.
POSITION_ISSUE_LIMIT = 8


def position_issues_message(items: list[dict]) -> str:
    """Сводка ошибок позиций одним сообщением (U6, UX-аудит T2).

    Чистая функция для тестов; поведение — Tk document_save_controller:
    не более POSITION_ISSUE_LIMIT строк, остальное — «…и ещё N».
    """
    lines: list[str] = []
    for item in items:
        name = str(item.get("name") or "").strip() or "позиция без названия"
        if item.get("qty", 0) <= 0:
            lines.append(f"• «{name}» — количество должно быть больше нуля")
        if item.get("price", 0) <= 0:
            lines.append(f"• «{name}» — цена должна быть больше нуля")
    if not lines:
        return ""
    if len(lines) > POSITION_ISSUE_LIMIT:
        shown = "\n".join(lines[:POSITION_ISSUE_LIMIT])
        return f"{shown}\n• …и ещё {len(lines) - POSITION_ISSUE_LIMIT}"
    return "\n".join(lines)


class SpecPageController:
    """Страница «Спецификация»: контекст, итоги, предпросмотр, сохранение."""

    page_key = "spec"

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        self.page: QWidget | None = None
        self.table: SpecTableQt | None = None
        self.preview_timer: QTimer | None = None
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

        title = QLabel("Спецификация", page)
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            "Позиции поставки к договору; справа — таблица позиций и печатная форма.",
            page,
        )
        subtitle.setObjectName("pageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        # Панель «Документ для: …» — единый текущий контрагент (п. 4.2).
        party_bar = QFrame(page)
        party_bar.setObjectName("partyBar")
        party_layout = QHBoxLayout(party_bar)
        party_layout.setContentsMargins(12, 6, 12, 6)
        self.current_party_label = QLabel(
            "Документ для: контрагент не выбран", party_bar,
        )
        self.current_party_label.setObjectName("currentPartyLabel")
        party_layout.addWidget(self.current_party_label, 1)
        # M4: временная точка входа в окно правки карточки — Tk открывает
        # окно из реестра (панель действий), реестр станет живым в M5.
        self.edit_party_button = QPushButton("Редактировать карточку", party_bar)
        self.edit_party_button.clicked.connect(self.edit_current_party)
        party_layout.addWidget(self.edit_party_button)
        outer.addWidget(party_bar)

        body = QHBoxLayout()
        body.setSpacing(12)
        outer.addLayout(body, 1)

        # -- Левая колонка: карточки параметров -------------------------- #
        left_scroll = QScrollArea(page)
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setFixedWidth(app_common.LEFT_PANEL_WIDTH + 32)
        left_panel = QWidget(left_scroll)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(12)
        self._left_layout = left_layout
        left_scroll.setWidget(left_panel)
        body.addWidget(left_scroll)

        # Карточка «Контрагент» (п. 4.2 + поиск по ИНН U3).
        card_party = self._card("Контрагент")
        self.party_combo = QComboBox(card_party)
        self.party_combo.addItem("— выберите контрагента —")
        card_party.grid.addWidget(self.party_combo, 0, 0)
        search_row = QHBoxLayout()
        self.search_entry = QLineEdit(card_party)
        self.search_entry.setPlaceholderText("ИНН целиком")
        self.search_entry.setMaxLength(12)
        # Только цифры — как в Tk (validatecommand, U3).
        self.search_entry.setValidator(
            QRegularExpressionValidator(_INN_DIGITS_RE, self.search_entry)
        )
        self.search_entry.returnPressed.connect(self.find_party_by_inn)
        search_row.addWidget(self.search_entry, 1)
        find_button = QPushButton("Найти", card_party)
        find_button.clicked.connect(self.find_party_by_inn)
        search_row.addWidget(find_button)
        card_party.grid.addLayout(search_row, 1, 0)
        hint = QLabel(
            "Найдёт контрагента в реестре и подставит его во все страницы.",
            card_party,
        )
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        card_party.grid.addWidget(hint, 2, 0)

        # Карточка «Договор-основание» (п. 4.2: явная строка выбранного).
        card_contract = self._card("Договор-основание")
        self.contract_combo = QComboBox(card_contract)
        self.contract_combo.addItem("— нет сохранённых договоров —")
        card_contract.grid.addWidget(self.contract_combo, 0, 0)
        self.contract_info = QLabel("Договор-основание не выбран.", card_contract)
        self.contract_info.setObjectName("fieldHint")
        self.contract_info.setWordWrap(True)
        card_contract.grid.addWidget(self.contract_info, 1, 0)

        # Карточка «Спецификация»: номер и дата.
        card_spec = self._card("Спецификация")
        number_label = QLabel("Номер спецификации *", card_spec)
        self.number_entry = QLineEdit(card_spec)
        self.number_entry.textChanged.connect(
            lambda _text: self.set_number_error("")
        )
        card_spec.grid.addWidget(number_label, 0, 0)
        card_spec.grid.addWidget(self.number_entry, 1, 0)
        self.number_error = QLabel("", card_spec)
        self.number_error.setObjectName("fieldError")
        card_spec.grid.addWidget(self.number_error, 2, 0)

        date_label = QLabel("Дата", card_spec)
        date_row = QHBoxLayout()
        self.day_entry = QLineEdit(card_spec)
        self.day_entry.setPlaceholderText("число")
        self.day_entry.setMaxLength(2)
        self.day_entry.setFixedWidth(56)
        self.month_combo = QComboBox(card_spec)
        self.month_combo.addItems(list(MONTHS_GENITIVE))
        self.month_combo.setCurrentText(MONTHS_GENITIVE[datetime.now().month - 1])
        today_button = QPushButton("Сегодня", card_spec)
        today_button.clicked.connect(self._fill_today)
        date_row.addWidget(self.day_entry)
        date_row.addWidget(QLabel("число", card_spec))
        date_row.addWidget(self.month_combo, 1)
        date_row.addWidget(today_button)
        card_spec.grid.addWidget(date_label, 3, 0)
        card_spec.grid.addLayout(date_row, 4, 0)

        # Карточка «Условия поставки».
        card_terms = self._card("Условия поставки")
        for row, (caption, attr, placeholder) in enumerate((
            (F.DELIVERY_TERM, "delivery_term",
             "Например: 30 дней с момента подписания договора."),
            (F.PAYMENT_TERMS, "payment_terms",
             "Например: оплата в течение 10 банковских дней после отгрузки."),
            (F.DELIVERY_BASIS, "delivery_basis",
             "Например: самовывоз со склада поставщика."),
        )):
            label = QLabel(caption, card_terms)
            entry = QLineEdit(card_terms)
            entry.setPlaceholderText(placeholder)
            setattr(self, f"{attr}_entry", entry)
            card_terms.grid.addWidget(label, row * 2, 0)
            card_terms.grid.addWidget(entry, row * 2 + 1, 0)
        prices_label = QLabel("Цены действительны до", card_terms)
        self.prices_entry = QLineEdit(card_terms)
        self.prices_entry.setText(prices_valid_until_default())
        prices_hint = QLabel(
            "Автоматически — последний день текущего месяца", card_terms,
        )
        prices_hint.setObjectName("fieldHint")
        card_terms.grid.addWidget(prices_label, 6, 0)
        card_terms.grid.addWidget(self.prices_entry, 7, 0)
        card_terms.grid.addWidget(prices_hint, 8, 0)
        left_layout.addStretch(1)

        # -- Правая колонка: таблица позиций + предпросмотр ---------------- #
        splitter = QSplitter(Qt.Orientation.Vertical, page)
        splitter.setChildrenCollapsible(False)
        body.addWidget(splitter, 1)

        table_holder = QWidget(splitter)
        table_layout = QVBoxLayout(table_holder)
        table_layout.setContentsMargins(0, 0, 0, 0)
        table_layout.setSpacing(4)
        # PERF-1: сообщение об обрезке вставки — тостом, немодально.
        self.table = SpecTableQt(
            table_holder,
            confirm_destructive=self._confirm_destructive,
        )
        self.table.changed.connect(self.on_spec_changed)
        self.table.notice.connect(lambda message: app.notify(message))
        table_layout.addWidget(self.table, 1)
        self.totals_label = QLabel("", table_holder)
        self.totals_label.setObjectName("specTotals")
        table_layout.addWidget(self.totals_label)
        splitter.addWidget(table_holder)

        preview_holder = QWidget(splitter)
        preview_layout = QVBoxLayout(preview_holder)
        preview_layout.setContentsMargins(0, 4, 0, 0)
        preview_layout.setSpacing(4)
        preview_title = QLabel(
            "Предпросмотр печатной формы (только чтение)", preview_holder,
        )
        preview_title.setObjectName("previewTitle")
        preview_layout.addWidget(preview_title)
        self.preview_text = QPlainTextEdit(preview_holder)
        self.preview_text.setReadOnly(True)  # правки не влияют на .docx (п. 4.3)
        self.preview_text.setObjectName("previewText")
        preview_layout.addWidget(self.preview_text, 1)
        splitter.addWidget(preview_holder)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 4)

        # Отложенный предпросмотр (п. 5.6): один QTimer, перезапуск заменой.
        self.preview_timer = QTimer(page)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(PREVIEW_DELAY_MS)
        self.preview_timer.timeout.connect(self.render_spec_preview)

        self._install_page()
        self.refresh_parties(keep_selection=False)
        self.update_spec_totals()
        self.render_spec_preview()

    def _card(self, caption: str) -> QFrame:
        """Карточка левой колонки: рамка + заголовок + сетка полей.

        Сетка доступна как ``card.grid`` (компактный доступ из _build_page).
        """
        card = QFrame(self.page)
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(6)
        caption_label = QLabel(caption, card)
        caption_label.setObjectName("cardCaption")
        layout.addWidget(caption_label)
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(4)
        layout.addLayout(grid)
        card.grid = grid  # noqa: B010 — компактный доступ из _build_page
        self._left_layout.addWidget(card)
        return card

    def _install_page(self) -> None:
        """Заменяет плейсхолдер M2 живой страницей «Спецификация»."""
        app = self.app
        placeholder = app.pages.get("spec")
        index = app.stack.indexOf(placeholder) if placeholder is not None else -1
        if placeholder is not None:
            app.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        app.stack.insertWidget(max(0, index), self.page)
        app.pages["spec"] = self.page
        if app.current_page == "spec":
            app.stack.setCurrentWidget(self.page)

    # ------------------------------------------------------------------ #
    #  Контрагенты и договор-основание (п. 4.2)
    # ------------------------------------------------------------------ #
    def party_duplicates(self) -> set[str]:
        """Названия, которые встречаются у нескольких карточек (п. 4.1)."""
        counts: dict[str, int] = {}
        for rec in self.app.store.active_counterparties():
            base = storage.display_with_inn(rec)
            counts[base] = counts.get(base, 0) + 1
        return {base for base, count in counts.items() if count > 1}

    def party_display(self, rec: dict, duplicates: set[str] | None = None) -> str:
        """Имя в списке; при совпадении названий добавляется ID (п. 4.1)."""
        base = storage.display_with_inn(rec)
        if duplicates and base in duplicates:
            return f"{base} · ID {rec.get('id')}"
        return base

    def refresh_parties(self, keep_selection: bool = True) -> None:
        """Обновляет список контрагентов, сбрасывает устаревший выбор (п. 4.2)."""
        current = self.party_combo.currentText() if keep_selection else ""
        recs = self.app.store.active_counterparties()
        duplicates = self.party_duplicates()
        self.party_combo.blockSignals(True)
        try:
            self.party_combo.clear()
            values = [self.party_display(rec, duplicates) for rec in recs]
            if values:
                self.party_combo.addItems(values)
            else:
                self.party_combo.addItem("— контрагентов нет —")
            if current and current in values:
                self.party_combo.setCurrentText(current)
            elif recs:
                self.party_combo.setCurrentIndex(0)
        finally:
            self.party_combo.blockSignals(False)
        self.on_party_changed()

    def selected_party(self) -> dict | None:
        """Запись контрагента по выбранному значению списка (п. 4.2)."""
        value = self.party_combo.currentText().strip()
        if not value or value.startswith("—"):
            return None
        recs = self.app.store.active_counterparties()
        duplicates = self.party_duplicates()
        for rec in recs:
            if self.party_display(rec, duplicates) == value:
                return rec
        # Совместимость: значение могло быть установлено без ID.
        for rec in recs:
            if storage.display_with_inn(rec) == value:
                return rec
        return None

    def set_current_party(self, rec: dict, source: str | None = None) -> None:
        """Выбор контрагента (поиск по ИНН; M4+ — перенос между страницами)."""
        self.app.current_party_id = rec.get("id")
        display = self.party_display(rec, self.party_duplicates())
        self.party_combo.blockSignals(True)
        try:
            self.party_combo.setCurrentText(display)
        finally:
            self.party_combo.blockSignals(False)
        self.on_party_changed(source=source)

    def _sync_current_party(self) -> None:
        """Панель «Документ для: …» (п. 4.2; Tk update_current_party_labels)."""
        party_id = getattr(self.app, "current_party_id", None)
        if party_id is None:
            text = "Документ для: контрагент не выбран"
        else:
            rec = self.app.store.get_counterparty(party_id)
            if rec is None:
                text = "Документ для: выбранный контрагент не найден — выберите заново"
            else:
                text = f"Документ для: {self.party_display(rec, self.party_duplicates())}"
        self.current_party_label.setText(text)

    def edit_current_party(self) -> None:
        """Открывает окно правки текущего контрагента (M4, до M5 — реестр).

        Tk-точка входа — панель действий реестра; в Qt-сборке реестр
        станет живым в M5, поэтому окно правки карточки доступно отсюда.
        После M5 точка остаётся или уходит — решение этапа M5.
        """
        party_id = getattr(self.app, "current_party_id", None)
        rec = (
            self.app.store.get_counterparty(party_id)
            if party_id is not None else None
        )
        if rec is None:
            self.app.alert(
                self.app.window, "Контрагент не выбран",
                "Выберите контрагента в списке слева — редактировать "
                "будем его карточку.", "error",
            )
            return
        counterparty_page_qt.open_edit_dialog(self.app, rec)

    def find_party_by_inn(self) -> None:
        """Ищет контрагента по ИНН в реестре и подставляет его (U3).

        Контракт Tk find_party_by_inn (document_page_controller): чистый
        store.find_by_inn, честный алерт при отсутствии, поле очищается.
        """
        inn = digits_only(self.search_entry.text())
        rec = self.app.store.find_by_inn(inn)
        if rec is None:
            self.app.alert(
                self.app.window, "Контрагент не найден",
                f"В реестре программы нет контрагента с ИНН {inn or '—'}.\n"
                "Добавьте его через кнопку «Добавить контрагента».",
                "error",
            )
            return
        self.set_current_party(rec)
        self.search_entry.clear()

    def refresh_contracts(self) -> None:
        """Список договоров выбранного контрагента + автоподстановка (п. 5.1)."""
        rec = self.selected_party()
        contracts = self.app.store.contracts_for(rec["id"]) if rec else []
        self.contract_combo.blockSignals(True)
        try:
            self.contract_combo.clear()
            if contracts:
                self.contract_combo.addItems(
                    [self.contract_display(c) for c in contracts]
                )
            else:
                self.contract_combo.addItem("— нет сохранённых договоров —")
        finally:
            self.contract_combo.blockSignals(False)
        if contracts:
            self.contract_info.setText(
                f"Договор-основание: {self.contract_display(contracts[0])} — "
                "подставлен автоматически, проверьте."
            )
        else:
            self.contract_info.setText(
                "Договор-основание не выбран: сохраните договор поставки."
            )

    def selected_contract(self) -> dict | None:
        """Запись договора-основания по выбранному значению (п. 4.2)."""
        rec = self.selected_party()
        if rec is None:
            return None
        value = self.contract_combo.currentText().strip()
        if not value or value.startswith("—"):
            return None
        for contract in self.app.store.contracts_for(rec["id"]):
            if self.contract_display(contract) == value:
                return contract
        return None

    @staticmethod
    def contract_display(rec: dict) -> str:
        """Подпись договора: «№ 516 от 10 января» (п. 4.2)."""
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        return f"№ {number} от {date_text}" if date_text else f"№ {number}"

    def on_party_changed(self, *args, source: str | None = None) -> None:
        """Смена контрагента: запомнить текущего, обновить договоры/контекст.

        Контракт Tk on_party_selected (п. 4.2): выбор запоминается в
        общем контексте каркаса и переносится на остальные страницы
        документов (M6: договор и доп. соглашение); пустой выбор честно
        сбрасывает контекст. Перенос — через каркас (fan-out), рекурсия
        исключена guard'ом _syncing_party.
        """
        rec = self.selected_party()
        self.app.current_party_id = rec["id"] if rec else None
        self.refresh_contracts()
        self._sync_current_party()
        self.regenerate_spec()
        if rec is not None and not getattr(self.app, "_syncing_party", False):
            self.app.set_current_party(rec, source="spec")

    def on_contract_changed(self, *args) -> None:
        """Смена договора-основания: явная строка + предпросмотр (п. 4.2)."""
        value = self.contract_combo.currentText().strip()
        if not value or value.startswith("—"):
            self.contract_info.setText(
                "Договор-основание не выбран — укажите его явно."
            )
        else:
            self.contract_info.setText(f"Договор-основание: {value}")
        self.regenerate_spec()

    # ------------------------------------------------------------------ #
    #  Итог и предпросмотр (U12, п. 5.6, P-07)
    # ------------------------------------------------------------------ #
    def update_spec_totals(self) -> None:
        """Строка «Позиций: N · Итого: …» под таблицей (U12).

        Один источник суммы — document_builder.spec_total(rows): столько же
        пишется в реестр (extra.total) и печатную форму. Поведение Tk:
        считаются все строки таблицы (пустые дают нулевой вклад).
        """
        rows = self.table.row_values() if self.table is not None else []
        total = spec_total(rows)
        self.totals_label.setText(
            f"Позиций: {len(rows)} · Итого: {total:.2f}".replace(".", ",")
            if rows else "Позиции не добавлены"
        )

    def on_spec_changed(self) -> None:
        """Пересчёт итога + отложенный предпросмотр при каждом изменении."""
        self.update_spec_totals()
        self.regenerate_spec()

    def regenerate_spec(self) -> None:
        """Откладывает пересборку предпросмотра на 120 мс (п. 5.6)."""
        if self.preview_timer is not None:
            self.preview_timer.start()  # перезапуск заменяет предыдущий отсчёт

    def flush_spec_preview(self) -> None:
        """Синхронно дособирает отложенный предпросмотр (перед сохранением)."""
        if self.preview_timer is not None:
            self.preview_timer.stop()
        self.render_spec_preview()

    def build_document_record(self) -> dict | None:
        """Черновик записи спецификации из состояния страницы.

        Одна сборка для предпросмотра и сохранения: предпросмотр строит
        печатную форму из того же черновика, что уйдёт в файл и реестр —
        расхождение исключено по построению (раздел Е, критерии 3/5).
        """
        party = self.selected_party()
        if party is None:
            return None
        contract = self.selected_contract()
        data = dict(party.get("fields", {}))
        items = self.table.get_items()
        number = self.number_entry.text().strip() or "б/н"
        day = self.day_entry.text().strip()
        month = self.month_combo.currentText().strip()
        total = spec_total(items)
        return {
            "kind": storage.KIND_SPEC,
            "party_id": party["id"],
            "number": number,
            "date": f"{day} {month}".strip(),
            "data": data,
            "extra": {
                "contract_id": (contract or {}).get("id"),
                "contract_number": (contract or {}).get("number") or "",
                "items": [
                    {"name": str(i["name"]), "qty": float(i["qty"]),
                     "unit": str(i["unit"]), "price": float(i["price"])}
                    for i in items
                ],
                "total": float(total),
                "spec_day": day,
                "spec_month": month,
                "prices_until": self.prices_entry.text().strip(),
                "delivery_term": self.delivery_term_entry.text().strip(),
                "payment_terms": self.payment_terms_entry.text().strip(),
                "delivery_basis": self.delivery_basis_entry.text().strip(),
                "is_ip": party.get("type") == "ip",
            },
            "text": "",
        }

    def render_spec_preview(self) -> None:
        """Собирает предпросмотр печатной формы без записи на диск (раздел Е).

        Печатная форма — из шаблона spec_template*.docx через
        services.build_record_docx (тот же путь, что и запись файла);
        остаточные {{…}} показываются до попытки сохранения (DEF-2).
        """
        if self.app.current_page != "spec":
            # UPD-UI (P-07): предпросмотр — только при видимой странице;
            # при переключении на «Спецификацию» его пересобирает
            # on_page_open.
            return
        rec = self.build_document_record()
        if rec is None:
            self._set_preview_text(
                "Выберите контрагента и договор-основание слева."
            )
            return
        try:
            doc = services.build_record_docx(rec, self.app.store)
            text = docx_to_text(doc)
            tokens = unfilled_placeholders(doc)
            if tokens:
                shown = ", ".join(tokens[:5]) + (" …" if len(tokens) > 5 else "")
                text += (
                    "\n\nВ печатной форме остались незаполненные поля: "
                    + shown + " — сохранение будет отклонено."
                )
        except Exception as exc:
            # Сбор деталей не роняет предпросмотр (Tk-контракт DEF-2).
            self.log.warning(
                "Не удалось собрать предпросмотр спецификации: %s",
                type(exc).__name__,
            )
            text = (
                "Не удалось собрать печатную форму:\n"
                f"{type(exc).__name__}: {exc}"
            )
        self._set_preview_text(text)

    def _set_preview_text(self, text: str) -> None:
        self.preview_text.setPlainText(text)

    def on_page_open(self) -> None:
        """Переключение на видимую страницу: пересборка по P-07 (п. 5.6)."""
        self.flush_spec_preview()
        self.update_spec_totals()

    def _fill_today(self) -> None:
        """Кнопка «Сегодня»: день и месяц текущей даты (Tk _fill_date)."""
        today = datetime.now()
        self.day_entry.setText(str(today.day))
        self.month_combo.setCurrentText(MONTHS_GENITIVE[today.month - 1])

    # ------------------------------------------------------------------ #
    #  Ошибка поля номера (подсветка, Tk-паритет п. 4.4)
    # ------------------------------------------------------------------ #
    def set_number_error(self, message: str) -> None:
        self.number_error.setText(message)
        self.number_error.setObjectName(
            "fieldErrorActive" if message else "fieldError"
        )
        self.number_error.style().unpolish(self.number_error)
        self.number_error.style().polish(self.number_error)

    # ------------------------------------------------------------------ #
    #  Подтверждения и несохранённые данные (п. 4.4)
    # ------------------------------------------------------------------ #
    def _confirm_destructive(self, message: str) -> bool:
        """Подтверждение деструктивных действий таблицы (U7)."""
        return self.app.confirm(
            self.app.window, "Таблица позиций", message,
            ok_text="Продолжить", cancel_text="Отмена", destructive=True,
        )

    def form_signature(self) -> tuple | None:
        """Данные формы, потеря которых требует подтверждения (п. 4.4).

        Выбор контрагента/договора намеренно не входит (переносится между
        страницами). Сбой чтения — None: страница считается «грязной»
        (аудит 2.0.1), а не «чистой».
        """
        try:
            return (
                self.number_entry.text(),
                self.day_entry.text(),
                self.month_combo.currentText(),
                self.delivery_term_entry.text(),
                self.payment_terms_entry.text(),
                self.delivery_basis_entry.text(),
                self.prices_entry.text(),
                self.table.snapshot(),
            )
        except Exception as exc:
            self.log.warning(
                "Не удалось снять подпись формы «spec»: %s", type(exc).__name__,
            )
            return None

    def is_form_dirty(self) -> bool:
        """Сравнение подписи с baseline; сбой чтения (None) — «грязная»."""
        return self.form_signature() != self.baseline

    def reset_form_baseline(self) -> None:
        """Baseline страницы пересчитан (построение, после сохранения)."""
        self.baseline = self.form_signature()

    # ------------------------------------------------------------------ #
    #  Папки контрагентов и транзакция сохранения (п. 9.2)
    # ------------------------------------------------------------------ #
    def documents_dir(self) -> str:
        """Папка документов: настройка «documents_dir» либо «Документы/EW.DocHub».

        Non-interactive ветка settings_controller.documents_dir (п. 2.8):
        выбранная папка, иначе видимая папка «Документы» (shell_open).
        Интерактивные диалоги выбора — вместе со страницей «Настройки»
        (M6); сейчас недоступная папка честно отклоняет сохранение.
        """
        custom = str((self.app.settings.get("documents_dir") or "")).strip()
        if custom:
            try:
                os.makedirs(custom, exist_ok=True)
                return custom
            except OSError as exc:
                raise OSError(
                    "Папка документов недоступна:\n"
                    f"{custom}\n{exc}\n\n"
                    "Выберите другую папку в «Настройках»."
                ) from exc
        base = os.path.join(shell_open.documents_known_folder(), "EW.DocHub")
        try:
            os.makedirs(base, exist_ok=True)
        except OSError as exc:
            raise OSError(
                "Папка документов по умолчанию недоступна:\n"
                f"{base}\n{exc}\n\nПроверьте права доступа."
            ) from exc
        return base

    def party_documents_dir(self, party_rec: dict | None) -> str:
        """Индивидуальная папка контрагента внутри общей папки (п. 9.2).

        Имя папки санитизируется, итоговый путь проверяется на выход за
        пределы папки документов (п. 2.7, services.is_inside_directory).
        """
        base = self.documents_dir()
        if party_rec is None:
            return base
        folder = safe_filename(storage.display_name(party_rec), "Контрагент")
        path = os.path.normpath(os.path.join(base, folder))
        if (
            not services.is_inside_directory(base, path)
            or os.path.normcase(path) == os.path.normcase(os.path.normpath(base))
        ):
            self.log.error(
                "Имя папки контрагента выходит за пределы папки документов",
            )
            self.app.alert(
                self.app.window, "Недопустимое имя папки",
                "Имя контрагента не может быть использовано как папка. "
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        try:
            os.makedirs(path, exist_ok=True)
        except OSError as exc:
            self.log.warning(
                "Не удалось создать папку контрагента: %s", type(exc).__name__,
            )
            self.app.alert(
                self.app.window, "Не удалось создать папку контрагента",
                f"{path}\n\n{exc}\n\n"
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        return path

    # ------------------------------------------------------------------ #
    #  Сохранение (п. 4.1, 9.2)
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        """Точка входа Ctrl+S каркаса (Tk: DocumentContextController.save_current_page)."""
        self.save_spec()

    def save_spec(self) -> None:
        """Сохранение спецификации — Tk-паритет document_save_controller."""
        app = self.app
        try:
            party = self.selected_party()
            if party is None:
                app.alert(app.window, "Контрагент не выбран",
                          "Выберите контрагента из списка слева.", "error")
                return
            contract = self.selected_contract()
            if contract is None:
                app.alert(app.window, "Договор не выбран",
                          "Выберите договор-основание. Если договоров нет — "
                          "создайте договор поставки на соответствующей "
                          "странице.", "error")
                return

            number = self.number_entry.text().strip()
            if not number:
                self.set_number_error("Заполните поле")
                self.number_entry.setFocus()
                app.alert(app.window, "Не хватает данных",
                          "Укажите номер спецификации.", "error")
                return

            items = self.table.get_items()
            if not items:
                app.alert(app.window, "Нет позиций",
                          "Добавьте хотя бы одну позицию с наименованием.",
                          "error")
                return
            # U6 (UX-A): одна сводка на все ошибочные позиции вместо
            # алерта на каждую — остальные ошибки видны сразу.
            issues = position_issues_message(items)
            if issues:
                app.alert(app.window, "Проверьте позиции", issues, "error")
                return

            # Предпросмотр собирается с задержкой — перед сохранением
            # дособираем синхронно, чтобы текст записи был актуальным.
            self.flush_spec_preview()
            text = self.preview_text.toPlainText().strip()
            date_text = (
                f"{self.day_entry.text().strip()} "
                f"{self.month_combo.currentText().strip()}"
            ).strip()
            if not self.confirm_document_duplicate(party["id"], number, date_text):
                return

            rec = self.build_document_record()
            if rec is None:  # контрагент deselected между проверками — защита
                app.alert(app.window, "Контрагент не выбран",
                          "Выберите контрагента из списка слева.", "error")
                return
            rec["number"] = number
            rec["date"] = date_text
            rec["text"] = text
            # Decimal не сериализуется в JSON — кол-во/цены float'ами,
            # итог — общий расчёт с предпросмотром и печатью (п. 9.8).
            result = services.save_document(
                app.store, rec, self.party_documents_dir(party),
            )
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            self.reset_form_baseline()
            app.refresh_status()
            opened = app.open_file(path)
            app.notify(
                f"Спецификация № {number} к договору № "
                f"{contract.get('number') or 'б/н'} сохранена"
                + (" и открыта." if opened else "."),
                path,
            )
        except Exception as exc:
            app.alert(
                app.window, "Ошибка",
                f"Не удалось сохранить спецификацию:\n{exc}", "error",
            )

    def confirm_document_duplicate(self, party_id, number: str,
                                   date_text: str) -> bool:
        """Предупреждает о документе с тем же номером у того же контрагента (п. 4.1)."""
        number = str(number or "").strip()
        if not number:
            return True
        for rec in self.app.store.documents_for(party_id):
            if rec.get("kind") != storage.KIND_SPEC:
                continue
            if str(rec.get("number", "")).strip() != number:
                continue
            existing_date = str(rec.get("date", "")).strip()
            details = f"Спецификация № {number}"
            if existing_date:
                details += f" от {existing_date}"
            return self.app.confirm(
                self.app.window, "Повтор номера",
                f"У этого контрагента уже есть {details}.\n\n"
                "Создать ещё один документ с тем же номером?",
                ok_text="Создать ещё", cancel_text="Отмена",
            )
        return True

    def persist_document_result_alert(self, result: dict) -> None:
        """Честное сообщение при сбое транзакции сохранения (п. 9.2, DEF-4).

        Сервисный слой формирует исход отката: «Созданный файл удалён,
        запись в реестр не добавлена» — либо имя и путь остаточного файла
        (result['orphan_path']), когда удалить его не удалось. Сообщение
        «файл удалён» не показывается, если файл на диске.
        """
        message = result["error"]
        orphan = result.get("orphan_path")
        if orphan and os.path.basename(str(orphan)) not in message:
            message += (
                "\n\nФайл удалить не удалось — он остался на диске:\n"
                f"{orphan}"
            )
        elif not orphan and "запись в реестр не добавлена" not in message.lower():
            message += "\n\nФайл удалён, запись в реестр не добавлена."
        self.app.alert(
            self.app.window, "Не удалось сохранить документ", message, "error",
        )
