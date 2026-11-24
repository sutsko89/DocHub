"""Страницы «Договор поставки» и «Доп. соглашение» на Qt (этап M6).

Qt-перенос ``document_page_controller.DocumentPageController`` и сценариев
сохранения ``document_save_controller.save_contract / save_additional`` (Tk):
отложенные текстовые предпросмотры (п. 5.6), поиск контрагента по ИНН со
страницы (U3), единый «текущий контрагент» для всех страниц документов
(п. 4.2), подстановка договора-основания при выборе контрагента (п. 5.1),
поля «сумма/срок прописью» и «Действует до» с дефолтом 31 декабря текущего
года (2.18.0).

Контракты, переносимые дословно (Tk-паритет):

* предпросмотр собирается с задержкой 120 мс — не рендерить .docx на
  каждый символ (п. 5.6); рендер — только при видимой странице (P-07),
  при переключении на страницу его пересобирает ``on_page_open``;
* предпросмотр — печатная форма из шаблона без записи на диск;
  текст предпросмотра только для чтения (п. 4.3);
* LAZY-1: тело ленивого договора дочитывается единым getter'ом
  ``store.document_body`` — предпросмотр и сохранение доп. соглашения
  получают полный текст даже после ленивой загрузки;
* выбор контрагента/договора не входит в подпись формы (он переносится
  между страницами, п. 4.4); сбой чтения подписи — «грязная» страница
  (аудит 2.0.1);
* сохранение — транзакция сервисного слоя «файл + запись» (п. 9.2,
  DEF-4): честное сообщение при сбое, «Файл удалён, запись в реестр не
  добавлена» либо имя остаточного файла; подтверждение повтора номера
  у того же контрагента (п. 4.1);
* после сохранения: реестр обновляется, статус, тост «…сохранён» с
  «Открыть файл»/«Показать в папке», файл сразу открывается в Word.

Отличия от Tk-сборки (зафиксированы в docs/qt-migration-stage-notes.md):
виджеты Qt; курсор при переформатировании лимита восстанавливается
QLineEdit.setCursorPosition (Tk restore_caret_after_reformat); диалог
сохранения/подтверждения — app.alert/app.confirm (qt_dialogs).
"""

from __future__ import annotations

import os
from datetime import datetime

from PySide6.QtCore import QRegularExpression, QTimer
from PySide6.QtGui import QRegularExpressionValidator
from PySide6.QtWidgets import (
    QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit,
    QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

import app_common
import counterparty_page_qt  # noqa: F401 — окно правки карточки (M4)
import lifecycle_qt
import services
import storage
from document_builder import (
    F, MONTHS_GENITIVE, amount_to_words_ru, digits_only, docx_to_text,
    format_number_grouped, number_to_words_ru, parse_valid_until,
    valid_until_default, valid_until_iso,
)

# Отложенный предпросмотр (п. 5.6): не собирать .docx на каждый символ.
PREVIEW_DELAY_MS = 120

# Поиск по ИНН — только цифры, не длиннее 12 (Tk: _digits_vcmd_for(12), U3).
_INN_DIGITS_RE = QRegularExpression("^\\d{0,12}$")
# Лимит: цифры и пробелы-разделители (Tk validate_limit_input: ≤ 15 цифр).
_LIMIT_RE = QRegularExpression("^[0-9 ]*$")
# Отсрочка: целое число дней, не более четырёх цифр (Tk validate_term_input).
_TERM_RE = QRegularExpression("^\\d{0,4}$")

PAGE_LABELS = {
    "contract": "Договор поставки",
    "additional": "Доп. соглашение",
    "spec": "Спецификация",
    "counterparty": "Карточка контрагента",
}


class _DocPageBaseQt:
    """Общая часть страниц договора и доп. соглашения (п. 4.2, U3, 9.2).

    Состояние страницы — у контроллера (раздел Д плана); подпись формы и
    baseline — здесь же (как Tk ``_form_baseline[key]`` на приложении).
    """

    page_key = ""
    page_title = ""
    page_subtitle = ""

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        self.page: QWidget | None = None
        self.preview_timer: QTimer | None = None
        self.baseline: tuple | None = None

    # ------------------------------------------------------------------ #
    #  Каркас страницы: шапка, панель «Документ для», колонки, предпросмотр
    # ------------------------------------------------------------------ #
    def _build_scaffold(self, preview_title: str) -> None:
        """Заголовок, подзаголовок, панель «Документ для: …», колонки."""
        app = self.app
        page = QWidget()
        self.page = page
        outer = QVBoxLayout(page)
        outer.setContentsMargins(24, 18, 24, 12)
        outer.setSpacing(6)

        title = QLabel(self.page_title, page)
        title.setObjectName("pageTitle")
        subtitle = QLabel(self.page_subtitle, page)
        subtitle.setObjectName("pageSubtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(title)
        outer.addWidget(subtitle)

        # Панель «Документ для: …» — единый текущий контрагент (п. 4.2);
        # «Выбрать в реестре» — переход в реестр (Tk _make_page with_text).
        party_bar = QFrame(page)
        party_bar.setObjectName("partyBar")
        party_layout = QHBoxLayout(party_bar)
        party_layout.setContentsMargins(12, 6, 12, 6)
        self.current_party_label = QLabel(
            "Документ для: контрагент не выбран", party_bar,
        )
        self.current_party_label.setObjectName("currentPartyLabel")
        party_layout.addWidget(self.current_party_label, 1)
        registry_button = QPushButton("Выбрать в реестре", party_bar)
        registry_button.clicked.connect(
            lambda _checked=False: app.select_page("registry"),
        )
        party_layout.addWidget(registry_button)
        outer.addWidget(party_bar)

        body = QHBoxLayout()
        body.setSpacing(12)
        outer.addLayout(body, 1)

        # Левая колонка: карточки параметров (прокручиваемая).
        left_scroll = QScrollArea(page)
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setFixedWidth(app_common.LEFT_PANEL_WIDTH + 32)
        left_panel = QWidget(left_scroll)
        self._left_layout = QVBoxLayout(left_panel)
        self._left_layout.setContentsMargins(0, 0, 8, 0)
        self._left_layout.setSpacing(12)
        left_scroll.setWidget(left_panel)
        body.addWidget(left_scroll)

        # Правая колонка: предпросмотр печатной формы (только чтение, 4.3).
        preview_holder = QWidget(page)
        preview_layout = QVBoxLayout(preview_holder)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(4)
        preview_label = QLabel(preview_title, preview_holder)
        preview_label.setObjectName("previewTitle")
        preview_layout.addWidget(preview_label)
        self.preview_text = QPlainTextEdit(preview_holder)
        self.preview_text.setReadOnly(True)
        self.preview_text.setObjectName("previewText")
        preview_layout.addWidget(self.preview_text, 1)
        body.addWidget(preview_holder, 1)

        # Отложенный предпросмотр (п. 5.6): один QTimer, перезапуск заменой.
        self.preview_timer = QTimer(page)
        self.preview_timer.setSingleShot(True)
        self.preview_timer.setInterval(PREVIEW_DELAY_MS)
        self.preview_timer.timeout.connect(self.render_preview)

        self._install_page()

    def _install_page(self) -> None:
        """Заменяет плейсхолдер M2 живой страницей."""
        app = self.app
        placeholder = app.pages.get(self.page_key)
        index = app.stack.indexOf(placeholder) if placeholder is not None else -1
        if placeholder is not None:
            app.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        app.stack.insertWidget(max(0, index), self.page)
        app.pages[self.page_key] = self.page
        if app.current_page == self.page_key:
            app.stack.setCurrentWidget(self.page)

    def _card(self, caption: str) -> QFrame:
        """Карточка левой колонки: рамка + заголовок + сетка полей."""
        card = QFrame(self.page)
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(6)
        caption_label = QLabel(caption, card)
        caption_label.setObjectName("cardCaption")
        layout.addWidget(caption_label)
        grid = QVBoxLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(4)
        layout.addLayout(grid)
        card.grid = grid  # noqa: B010 — компактный доступ из _build_page
        self._left_layout.addWidget(card)
        return card

    def _field_row(self, card: QFrame, caption: str, entry: QLineEdit,
                   hint: str = "", error: bool = False) -> QLabel:
        """Строка поля: подпись, ввод, подсказка; возвращает метку ошибки."""
        label = QLabel(caption + (" *" if error else ""), card)
        card.grid.addWidget(label)
        card.grid.addWidget(entry)
        error_label = QLabel("", card)
        error_label.setObjectName("fieldError")
        card.grid.addWidget(error_label)
        if hint:
            hint_label = QLabel(hint, card)
            hint_label.setObjectName("fieldHint")
            hint_label.setWordWrap(True)
            card.grid.addWidget(hint_label)
        return error_label

    # ------------------------------------------------------------------ #
    #  Контрагенты: список, поиск по ИНН (U3), единый текущий (п. 4.2)
    # ------------------------------------------------------------------ #
    def _build_party_card(self, card: QFrame, search_attr: str) -> None:
        """Карточка «Контрагент»: комбо + поиск по ИНН (U3, UX-A)."""
        self.party_combo = QComboBox(card)
        self.party_combo.addItem("— выберите контрагента —")
        card.grid.addWidget(self.party_combo)
        self.party_combo.currentTextChanged.connect(self.on_party_changed)

        search_row = QHBoxLayout()
        setattr(self, search_attr, QLineEdit(card))
        search_entry = getattr(self, search_attr)
        search_entry.setPlaceholderText("ИНН целиком")
        search_entry.setMaxLength(12)
        search_entry.setValidator(
            QRegularExpressionValidator(_INN_DIGITS_RE, search_entry)
        )
        search_entry.returnPressed.connect(
            lambda: self.find_party_by_inn(search_entry)
        )
        search_row.addWidget(search_entry, 1)
        find_button = QPushButton("Найти", card)
        find_button.clicked.connect(
            lambda _checked=False: self.find_party_by_inn(search_entry)
        )
        search_row.addWidget(find_button)
        card.grid.addLayout(search_row)
        hint = QLabel(
            "Найдёт контрагента в реестре и подставит его во все страницы.",
            card,
        )
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        card.grid.addWidget(hint)

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
        self._sync_current_party()

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
        # Совместимость: значение могло быть установлено программно без ID.
        for rec in recs:
            if storage.display_with_inn(rec) == value:
                return rec
        return None

    def set_current_party(self, rec: dict, source: str | None = None) -> None:
        """Подстановка контрагента на страницу (fan-out каркаса, п. 4.2).

        Комбо обновляется с заблокированными сигналами — пользовательский
        маршрут (on_party_changed) не срабатывает повторно.
        """
        display = self.party_display(rec, self.party_duplicates())
        self.party_combo.blockSignals(True)
        try:
            self.party_combo.setCurrentText(display)
        finally:
            self.party_combo.blockSignals(False)
        self._local_party_changed()

    def _sync_current_party(self) -> None:
        """Панель «Документ для: …» (Tk update_current_party_labels)."""
        party_id = getattr(self.app, "current_party_id", None)
        if party_id is None:
            text = "Документ для: контрагент не выбран"
        else:
            rec = self.app.store.get_counterparty(party_id)
            if rec is None:
                text = "Документ для: выбранный контрагент не найден — выберите заново"
            else:
                text = f"Документ для: {self.party_display(rec, self.party_duplicates())}"
        try:
            self.current_party_label.setText(text)
        except RuntimeError:
            pass  # страница могла быть разрушена

    def on_party_changed(self, *args, source: str | None = None) -> None:
        """Смена контрагента: локальные обновления + перенос выбора (п. 4.2).

        Выбор запоминается в общем контексте каркаса и переносится на
        остальные страницы документов; пустой выбор честно сбрасывает
        контекст. Перенос идёт через каркас (fan-out) — обратного вызова
        в страницу нет, рекурсия исключена флагом _syncing_party.
        """
        rec = self.selected_party()
        self.app.current_party_id = rec["id"] if rec else None
        self._local_party_changed()
        if rec is not None and not getattr(self.app, "_syncing_party", False):
            self.app.set_current_party(rec, source=self.page_key)

    def _local_party_changed(self) -> None:
        """Локальные хвосты смены контрагента (договоры, предпросмотр)."""
        self.refresh_contracts()
        self._sync_current_party()
        self.preview_deferred()

    def find_party_by_inn(self, search_entry: QLineEdit) -> None:
        """Ищет контрагента по ИНН и подставляет во все страницы (U3).

        Контракт Tk find_party_by_inn: чистый store.find_by_inn, честный
        алерт при отсутствии, поле очищается.
        """
        inn = digits_only(search_entry.text())
        rec = self.app.store.find_by_inn(inn)
        if rec is None:
            self.app.alert(
                self.app.window, "Контрагент не найден",
                f"В реестре программы нет контрагента с ИНН {inn or '—'}.\n"
                "Добавьте его через кнопку «Добавить контрагента».",
                "error",
            )
            return
        self.app.set_current_party(rec, source=self.page_key)
        search_entry.clear()

    # ------------------------------------------------------------------ #
    #  Папки и транзакция сохранения (п. 9.2) — общий путь с M5
    # ------------------------------------------------------------------ #
    def party_documents_dir(self, party_rec: dict | None) -> str:
        """Индивидуальная папка контрагента (lifecycle_qt, M5)."""
        return lifecycle_qt.party_documents_dir(self.app, party_rec)

    def persist_document(self, draft: dict) -> dict:
        """Пишет .docx и регистрирует запись через сервисный слой (п. 9.2)."""
        docs_dir = self.party_documents_dir(
            self.app.store.get_counterparty(draft.get("party_id"))
        )
        return services.save_document(self.app.store, draft, docs_dir)

    def persist_document_result_alert(self, result: dict) -> None:
        """Честное сообщение при сбое транзакции сохранения (п. 9.2, DEF-4)."""
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

    def confirm_document_duplicate(self, party_id, kind: str, number: str,
                                   date_text: str) -> bool:
        """Предупреждает о документе с тем же номером у того же контрагента (п. 4.1)."""
        number = str(number or "").strip()
        if not number:
            return True
        for rec in self.app.store.documents_for(party_id):
            if rec.get("kind") != kind:
                continue
            if str(rec.get("number", "")).strip() != number:
                continue
            existing_date = str(rec.get("date", "")).strip()
            title = storage.KIND_LABELS.get(kind, "Документ")
            details = f"{title} № {number}"
            if existing_date:
                details += f" от {existing_date}"
            return self.app.confirm(
                self.app.window, "Повтор номера",
                f"У этого контрагента уже есть {details}.\n\n"
                "Создать ещё один документ с тем же номером?",
                ok_text="Создать ещё", cancel_text="Отмена",
            )
        return True

    # ------------------------------------------------------------------ #
    #  Несохранённые данные (п. 4.4)
    # ------------------------------------------------------------------ #
    def form_signature(self) -> tuple | None:
        """Данные формы, потеря которых требует подтверждения.

        Выбор контрагента/договора намеренно не входит (переносится между
        страницами). Сбой чтения — None: страница считается «грязной»
        (аудит 2.0.1).
        """
        raise NotImplementedError

    def is_form_dirty(self) -> bool:
        """Сравнение подписи с baseline; сбой чтения (None) — «грязная»."""
        return self.form_signature() != self.baseline

    def reset_form_baseline(self) -> None:
        """Baseline страницы пересчитан (построение, после сохранения)."""
        self.baseline = self.form_signature()

    def on_page_open(self) -> None:
        """Переключение на видимую страницу: пересборка предпросмотра (P-07)."""
        self.render_preview()


class ContractPageQt(_DocPageBaseQt):
    """Страница «Договор поставки»: контрагент, номер, дата, предпросмотр."""

    page_key = "contract"
    page_title = "Договор поставки"
    page_subtitle = (
        "Выберите контрагента, укажите номер и дату. «Сохранить договор» "
        "запишет его в реестр и сразу откроет готовый документ в Word."
    )

    def __init__(self, app) -> None:
        super().__init__(app)
        self._build_page()
        self.reset_form_baseline()

    # ------------------------------------------------------------------ #
    #  Построение страницы
    # ------------------------------------------------------------------ #
    def _build_page(self) -> None:
        self._build_scaffold("Предпросмотр договора (только чтение)")

        # Карточка «Контрагент» (п. 4.2 + поиск по ИНН U3).
        card_party = self._card("Контрагент")
        self._build_party_card(card_party, "search_entry")

        # Карточка «Договор»: номер и дата.
        card_doc = self._card("Договор")
        self.number_entry = QLineEdit(card_doc)
        self._number_error = self._field_row(
            card_doc, F.NUMBER, self.number_entry,
            hint="Номер поставочного договора.", error=True,
        )
        self.number_entry.textChanged.connect(
            lambda _text: self.set_number_error("")
        )

        date_label = QLabel(F.DAY, card_doc)
        date_row = QHBoxLayout()
        self.day_entry = QLineEdit(card_doc)
        self.day_entry.setPlaceholderText("число")
        self.day_entry.setMaxLength(2)
        self.day_entry.setFixedWidth(56)
        self.month_combo = QComboBox(card_doc)
        self.month_combo.addItems(list(MONTHS_GENITIVE))
        self.month_combo.setCurrentText(MONTHS_GENITIVE[datetime.now().month - 1])
        today_button = QPushButton("Сегодня", card_doc)
        today_button.clicked.connect(self._fill_today)
        date_row.addWidget(self.day_entry)
        date_row.addWidget(QLabel("число", card_doc))
        date_row.addWidget(self.month_combo, 1)
        date_row.addWidget(today_button)
        card_doc.grid.addWidget(date_label)
        card_doc.grid.addLayout(date_row)
        self._left_layout.addStretch(1)

        for entry in (self.number_entry, self.day_entry):
            entry.textChanged.connect(lambda _t: self.preview_deferred())
        self.month_combo.currentTextChanged.connect(
            lambda _t: self.preview_deferred()
        )

        self.refresh_parties(keep_selection=False)
        self.render_preview()

    def _fill_today(self) -> None:
        """Кнопка «Сегодня»: день и месяц текущей даты (Tk _fill_date)."""
        today = datetime.now()
        self.day_entry.setText(str(today.day))
        self.month_combo.setCurrentText(MONTHS_GENITIVE[today.month - 1])

    def set_number_error(self, message: str) -> None:
        """Подсветка ошибки поля номера (Tk-паритет п. 4.4)."""
        self._number_error.setText(message)
        self._number_error.setObjectName(
            "fieldErrorActive" if message else "fieldError"
        )
        self._number_error.style().unpolish(self._number_error)
        self._number_error.style().polish(self._number_error)

    # ------------------------------------------------------------------ #
    #  Договор-основание: у договора нет — карточка скрыта (Tk-паритет)
    # ------------------------------------------------------------------ #
    def refresh_contracts(self) -> None:
        """Единый контракт с базой: у страницы договора договоров нет."""

    # ------------------------------------------------------------------ #
    #  Предпросмотр (п. 5.6, P-07)
    # ------------------------------------------------------------------ #
    def preview_deferred(self) -> None:
        """Откладывает перерисовку предпросмотра (п. 5.6)."""
        if self.preview_timer is not None:
            self.preview_timer.start()  # перезапуск заменяет предыдущий отсчёт

    def render_preview(self) -> None:
        """Собирает текст предпросмотра договора из текущих полей."""
        if self.app.current_page != self.page_key:
            # UPD-UI (P-07): предпросмотр — только при видимой странице;
            # операции из других страниц не рендерят скрытый текст.
            return
        rec = self.selected_party()
        if rec is None:
            self.preview_text.setPlainText(
                "Выберите контрагента и заполните поля слева — "
                "здесь появится текст договора."
            )
            return
        data = dict(rec.get("fields", {}))
        data[F.NUMBER] = self.number_entry.text().strip()
        data[F.DAY] = self.day_entry.text().strip()
        data[F.MONTH] = self.month_combo.currentText().strip()
        is_ip = rec.get("type") == "ip"
        try:
            doc = services.render_contract(data, is_ip)
            self.preview_text.setPlainText(docx_to_text(doc))
        except Exception as exc:
            self.log.warning(
                "Не удалось собрать предпросмотр договора: %s", type(exc).__name__,
            )
            self.preview_text.setPlainText(f"Не удалось собрать предпросмотр: {exc}")

    # ------------------------------------------------------------------ #
    #  Сохранение (п. 4.1, 9.2) — Tk save_contract
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        """Точка входа Ctrl+S каркаса (Tk: save_current_page → save_contract)."""
        self.save_contract()

    def save_contract(self) -> None:
        """Сохранение договора — Tk-паритет document_save_controller."""
        app = self.app
        try:
            rec = self.selected_party()
            if rec is None:
                app.alert(app.window, "Контрагент не выбран",
                          "Выберите контрагента из списка или добавьте нового "
                          "через кнопку «Добавить контрагента».", "error")
                return

            data = dict(rec.get("fields", {}))
            data[F.NUMBER] = self.number_entry.text().strip()
            data[F.DAY] = self.day_entry.text().strip()
            data[F.MONTH] = self.month_combo.currentText().strip()
            is_ip = rec.get("type") == "ip"

            errors = services.validate_contract_fields(data, is_ip)
            if errors:
                self.set_number_error(errors.get(F.NUMBER, ""))
                # Фокус — первое ошибочное поле (Tk: F.NUMBER → F.DAY).
                if errors.get(F.NUMBER):
                    self.number_entry.setFocus()
                elif errors.get(F.DAY):
                    self.day_entry.setFocus()
                app.alert(
                    app.window, "Ошибки в данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            date_text = f"{data[F.DAY]} {data[F.MONTH]}".strip()
            if not self.confirm_document_duplicate(
                rec["id"], storage.KIND_CONTRACT, data[F.NUMBER], date_text
            ):
                return
            extra = {"day": data[F.DAY], "month": data[F.MONTH], "is_ip": is_ip}
            draft = {
                "kind": storage.KIND_CONTRACT,
                "party_id": rec["id"],
                "number": data[F.NUMBER],
                "date": date_text,
                "data": data,
                "extra": extra,
                "text": "",
            }
            # Файл и запись реестра — транзакция сервисного слоя (п. 9.2).
            result = self.persist_document(draft)
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            self.reset_form_baseline()
            if getattr(app, "registry_flow", None) is not None:
                app.registry_flow.refresh()
            app.refresh_status()
            opened = app.open_file(path)
            app.notify(
                f"Договор № {data[F.NUMBER]} от {date_text} для "
                f"{storage.display_name(rec)} сохранён"
                + (" и открыт." if opened else "."),
                path,
            )
        except Exception as exc:
            app.alert(app.window, "Ошибка",
                      f"Не удалось сохранить договор:\n{exc}", "error")

    # ------------------------------------------------------------------ #
    #  Несохранённые данные (п. 4.4) — подпись Tk contract-страницы
    # ------------------------------------------------------------------ #
    def form_signature(self) -> tuple | None:
        """Номер и дата договора (выбор контрагента не входит, п. 4.4)."""
        try:
            return (
                self.number_entry.text(),
                self.day_entry.text(),
                self.month_combo.currentText(),
            )
        except Exception as exc:
            self.log.warning(
                "Не удалось снять подпись формы «contract»: %s", type(exc).__name__,
            )
            return None


class AdditionalPageQt(_DocPageBaseQt):
    """Страница «Доп. соглашение»: договор-основание, лимит, отсрочка, срок."""

    page_key = "additional"
    page_title = "Доп. соглашение"
    page_subtitle = (
        "Выберите контрагента и договор-основание. «Сохранить» запишет "
        "документ в реестр и сразу откроет его в Word."
    )

    def __init__(self, app) -> None:
        super().__init__(app)
        self._build_page()
        self.reset_form_baseline()

    # ------------------------------------------------------------------ #
    #  Построение страницы
    # ------------------------------------------------------------------ #
    def _build_page(self) -> None:
        self._build_scaffold("Предпросмотр доп. соглашения (только чтение)")

        # Карточка «Контрагент» (п. 4.2 + поиск по ИНН U3).
        card_party = self._card("Контрагент")
        self._build_party_card(card_party, "search_entry")

        # Карточка «Договор-основание» (п. 4.2: явная строка выбранного).
        card_contract = self._card("Договор-основание")
        self.contract_combo = QComboBox(card_contract)
        self.contract_combo.addItem("— нет сохранённых договоров —")
        card_contract.grid.addWidget(self.contract_combo)
        self.contract_combo.currentTextChanged.connect(self.on_contract_changed)
        self.contract_info = QLabel("Договор-основание не выбран.", card_contract)
        self.contract_info.setObjectName("fieldHint")
        self.contract_info.setWordWrap(True)
        card_contract.grid.addWidget(self.contract_info)
        note = QLabel(
            "Договоры берутся из реестра: сначала создайте договор поставки.",
            card_contract,
        )
        note.setObjectName("fieldHint")
        note.setWordWrap(True)
        card_contract.grid.addWidget(note)

        # Карточка «Условия»: номер, лимит, отсрочка, срок действия (2.18.0).
        card_terms = self._card("Условия")
        self.number_entry = QLineEdit(card_terms)
        self._field_row(
            card_terms, "Номер доп. соглашения", self.number_entry,
            hint="Можно оставить пустым — сохранится как «б/н».",
        )

        limit_label = QLabel(F.LIMIT + " *", card_terms)
        limit_holder = QWidget(card_terms)
        limit_layout = QHBoxLayout(limit_holder)
        limit_layout.setContentsMargins(0, 0, 0, 0)
        self.limit_entry = QLineEdit(limit_holder)
        self.limit_entry.setValidator(
            QRegularExpressionValidator(_LIMIT_RE, self.limit_entry)
        )
        limit_layout.addWidget(self.limit_entry, 1)
        limit_layout.addWidget(QLabel("руб.", limit_holder))
        self._limit_error = QLabel("", card_terms)
        self._limit_error.setObjectName("fieldError")
        card_terms.grid.addWidget(limit_label)
        card_terms.grid.addWidget(limit_holder)
        card_terms.grid.addWidget(self._limit_error)
        limit_hint = QLabel(
            "Сумма по доп. соглашению к договору поставки.", card_terms,
        )
        limit_hint.setObjectName("fieldHint")
        limit_hint.setWordWrap(True)
        card_terms.grid.addWidget(limit_hint)

        self.limit_words_label = QLabel("Сумма прописью: —", card_terms)
        self.limit_words_label.setObjectName("fieldHint")
        card_terms.grid.addWidget(self.limit_words_label)

        term_label = QLabel(F.TERM + " *", card_terms)
        term_holder = QWidget(card_terms)
        term_layout = QHBoxLayout(term_holder)
        term_layout.setContentsMargins(0, 0, 0, 0)
        self.term_entry = QLineEdit(term_holder)
        self.term_entry.setValidator(
            QRegularExpressionValidator(_TERM_RE, self.term_entry)
        )
        term_layout.addWidget(self.term_entry, 1)
        term_layout.addWidget(QLabel("дней", term_holder))
        self._term_error = QLabel("", card_terms)
        self._term_error.setObjectName("fieldError")
        card_terms.grid.addWidget(term_label)
        card_terms.grid.addWidget(term_holder)
        card_terms.grid.addWidget(self._term_error)
        term_hint = QLabel("Отсрочка платежа по доп. соглашению.", card_terms)
        term_hint.setObjectName("fieldHint")
        term_hint.setWordWrap(True)
        card_terms.grid.addWidget(term_hint)

        self.words_label = QLabel("Срок прописью: —", card_terms)
        self.words_label.setObjectName("fieldHint")
        card_terms.grid.addWidget(self.words_label)

        # Срок действия (2.18.0): дефолт — 31 декабря текущего года (п. 5).
        self.valid_entry = QLineEdit(card_terms)
        self.valid_entry.setText(valid_until_default())
        self._valid_error = self._field_row(
            card_terms, F.VALID_UNTIL, self.valid_entry,
            hint="Срок действия по п. 5 (формат ДД.ММ.ГГГГ). По умолчанию — "
                 "31 декабря текущего года.",
            error=True,
        )
        self._left_layout.addStretch(1)

        self.limit_entry.textChanged.connect(self.on_limit_changed)
        self.term_entry.textChanged.connect(self.on_term_changed)
        self.number_entry.textChanged.connect(lambda _t: self.preview_deferred())
        self.valid_entry.textChanged.connect(lambda _t: self.preview_deferred())

        self.update_words()
        self.update_limit_words()
        self.refresh_parties(keep_selection=False)
        self.render_preview()

    # ------------------------------------------------------------------ #
    #  Договор-основание: список договоров контрагента (п. 5.1)
    # ------------------------------------------------------------------ #
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

    @staticmethod
    def contract_display(rec: dict) -> str:
        """Подпись договора: «№ 516 от 10 января» (п. 4.2)."""
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        return f"№ {number} от {date_text}" if date_text else f"№ {number}"

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

    def on_contract_changed(self, *args) -> None:
        """Смена договора-основания: явная строка + предпросмотр (п. 4.2)."""
        value = self.contract_combo.currentText().strip()
        if not value or value.startswith("—"):
            self.contract_info.setText(
                "Договор-основание не выбран — укажите его явно."
            )
        else:
            self.contract_info.setText(f"Договор-основание: {value}")
        self.preview_deferred()

    # ------------------------------------------------------------------ #
    #  Суммы прописью и формат лимита (Tk format_limit/update_words)
    # ------------------------------------------------------------------ #
    def on_limit_changed(self, raw: str) -> None:
        """Форматирует лимит с разделителями разрядов + сумма прописью.

        Курсор восстанавливается через setCursorPosition (Tk
        restore_caret_after_reformat): длина слева от курсора сохраняется.
        """
        formatted = format_number_grouped(raw)
        if raw != formatted:
            cursor = self.limit_entry.cursorPosition()
            self.limit_entry.setText(formatted)
            self.limit_entry.setCursorPosition(
                min(max(cursor + len(formatted) - len(raw), 0), len(formatted))
            )
        self.update_limit_words()
        self.preview_deferred()

    def on_term_changed(self, _text: str) -> None:
        """Отсрочка изменилась: срок прописью + предпросмотр."""
        self.update_words()
        self.preview_deferred()

    def update_words(self) -> None:
        """«Срок прописью» под полем отсрочки платежа."""
        value = self.term_entry.text().strip()
        if not value:
            self.words_label.setText("Срок прописью: —")
        elif value.isdigit():
            self.words_label.setText(
                f"Срок прописью: {number_to_words_ru(int(value))}")
        else:
            self.words_label.setText("Срок прописью: введите целое число дней")

    def update_limit_words(self) -> None:
        """«Сумма прописью» под полем лимита доп. соглашения."""
        digits = digits_only(self.limit_entry.text())
        if not digits:
            self.limit_words_label.setText("Сумма прописью: —")
            return
        try:
            self.limit_words_label.setText(
                f"Сумма прописью: {amount_to_words_ru(int(digits))}")
        except Exception:
            self.limit_words_label.setText("Сумма прописью: —")

    def set_field_error(self, which: str, message: str) -> None:
        """Подсветка ошибки поля (Tk _set_page_error, п. 4.4)."""
        label = {
            "limit": self._limit_error,
            "term": self._term_error,
            "valid": self._valid_error,
        }.get(which)
        if label is None:
            return
        label.setText(message)
        label.setObjectName("fieldErrorActive" if message else "fieldError")
        label.style().unpolish(label)
        label.style().polish(label)

    # ------------------------------------------------------------------ #
    #  Предпросмотр (п. 5.6, P-07, LAZY-1)
    # ------------------------------------------------------------------ #
    def preview_deferred(self) -> None:
        """Откладывает перерисовку предпросмотра (п. 5.6)."""
        if self.preview_timer is not None:
            self.preview_timer.start()  # перезапуск заменяет предыдущий отсчёт

    def render_preview(self) -> None:
        """Собирает текст предпросмотра доп. соглашения из текущих полей."""
        if self.app.current_page != self.page_key:
            # UPD-UI (P-07): предпросмотр — только при видимой странице.
            return
        party = self.selected_party()
        contract = self.selected_contract() if party else None
        if party is None or contract is None:
            self.preview_text.setPlainText(
                "Выберите контрагента и договор-основание слева — "
                "здесь появится текст документа."
            )
            return
        # LAZY-1: тело ленивого договора дочитывается единым getter'ом.
        if "data" not in contract:
            contract = {
                **contract,
                **self.app.store.document_body(contract.get("id")),
            }
        data = dict(contract.get("data") or {})
        is_ip = party.get("type") == "ip"
        try:
            doc = services.render_additional(
                data, is_ip,
                self.number_entry.text().strip(),
                self.limit_entry.text().strip(),
                self.term_entry.text().strip(),
                self.valid_entry.text().strip(),
            )
            self.preview_text.setPlainText(docx_to_text(doc))
        except Exception as exc:
            self.log.warning(
                "Не удалось собрать предпросмотр доп. соглашения: %s",
                type(exc).__name__,
            )
            self.preview_text.setPlainText(f"Не удалось собрать предпросмотр: {exc}")

    # ------------------------------------------------------------------ #
    #  Сохранение (п. 4.1, 9.2) — Tk save_additional
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        """Точка входа Ctrl+S каркаса (Tk: save_current_page → save_additional)."""
        self.save_additional()

    def save_additional(self) -> None:
        """Сохранение доп. соглашения — Tk-паритет document_save_controller."""
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

            limit_display = self.limit_entry.text().strip()
            limit = digits_only(limit_display)
            term = self.term_entry.text().strip()
            number = self.number_entry.text().strip() or "б/н"
            valid_text = self.valid_entry.text().strip()
            valid_date = parse_valid_until(valid_text)

            errors = {}
            if not limit:
                errors["limit"] = "Заполните поле"
            if not term:
                errors["term"] = "Заполните поле"
            elif not term.isdigit():
                errors["term"] = "Должно быть целое число дней"
            if not valid_text:
                errors["valid"] = "Заполните поле"
            elif valid_date is None:
                errors["valid"] = "Введите дату в формате ДД.ММ.ГГГГ"
            if errors:
                for which in ("limit", "term", "valid"):
                    self.set_field_error(which, errors.get(which, ""))
                app.alert(
                    app.window, "Не хватает данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            contract_number = contract.get("number") or ""
            date_text = contract.get("date") or ""
            if number and number != "б/н" and not self.confirm_document_duplicate(
                party["id"], storage.KIND_ADDITIONAL, number, date_text
            ):
                return
            extra = {
                "contract_id": contract["id"],
                "contract_number": contract_number,
                "limit": limit_display,
                "limit_digits": limit,
                "term": term,
                "is_ip": party.get("type") == "ip",
                # Срок действия (2.18.0): вид «31.12.2026» для реестра и
                # ISO «2026-12-31» для расчёта «истекло/истекает».
                "valid_until": valid_date.strftime("%d.%m.%Y"),
                "valid_until_iso": valid_until_iso(valid_date),
            }
            draft = {
                "kind": storage.KIND_ADDITIONAL,
                "party_id": party["id"],
                "number": number,
                "date": date_text,
                # LAZY-1: тело договора дочитывается единым getter'ом.
                "data": dict(
                    (contract if "data" in contract
                     else {**contract,
                           **app.store.document_body(contract.get("id"))}).get("data") or {}
                ),
                "extra": extra,
                "text": "",
            }
            result = self.persist_document(draft)
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            self.reset_form_baseline()
            if getattr(app, "registry_flow", None) is not None:
                app.registry_flow.refresh()
            app.refresh_status()
            opened = app.open_file(path)
            app.notify(
                f"Доп. соглашение № {number} к договору № "
                f"{contract_number or 'б/н'} сохранено"
                + (" и открыто." if opened else "."),
                path,
            )
        except Exception as exc:
            app.alert(app.window, "Ошибка",
                      f"Не удалось сохранить доп. соглашение:\n{exc}", "error")

    # ------------------------------------------------------------------ #
    #  Несохранённые данные (п. 4.4) — подпись Tk additional-страницы
    # ------------------------------------------------------------------ #
    def form_signature(self) -> tuple | None:
        """Номер, лимит, отсрочка и срок действия (выбор не входит, п. 4.4)."""
        try:
            return (
                self.number_entry.text(),
                self.limit_entry.text(),
                self.term_entry.text(),
                self.valid_entry.text(),
            )
        except Exception as exc:
            self.log.warning(
                "Не удалось снять подпись формы «additional»: %s",
                type(exc).__name__,
            )
            return None
