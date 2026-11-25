"""Страница «Настройки» на Qt (этап M6) — перенос settings_controller.

Qt-перенос ``settings_controller.SettingsController`` (Tk): тема
оформления, папка документов (выбор с немедленным сохранением), ключ
DaData (сохранённый ключ не подставляется в поле — значение не видно и
не копируется, аудит 2.0.1), состояние автокопий реестра (BKUP-1),
диагностика (папка журналов) и справка по горячим клавишам (U14).

Контракты, переносимые дословно (Tk-паритет):

* тема меняется через каркас (``app.set_theme``) — палитра/QSS +
  settings.json + слушатели; отказ записи сообщается алертом, тема в
  сеансе остаётся применённой;
* папка документов сохраняется сразу после выбора («сразу, а не только
  при закрытии: иначе при аварийном завершении выбранная папка
  теряется»); отказ записи — без success-алерта;
* ключ DaData хранится в профиле пользователя отдельно от реестра
  (dadata.save_api_key — DPAPI на Windows); сохранённый ключ НЕ
  подставляется в поле; удаление — с подтверждением (деструктивное, U7);
* состояние копий реестра — storage.backup_status_text; папка копий
  создаётся при открытии, если копий ещё не было.

Отличия от Tk-сборки (docs/qt-migration-stage-notes.md): переключатель
темы — пара радиокнопок вместо самописного ThemeToggle; карточка
«Обновления» появится вместе с UI обновлений (этап M7 — оркестрация
updater_controller на QTimer); окно правки констант спецификации
отсутствует в обеих сборках — константы задаются settings.json
(app_common.get_spec_settings, M1).
"""

from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QButtonGroup, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QRadioButton, QScrollArea, QVBoxLayout, QWidget,
)

import applog
import dadata_client as dadata
import qt_theme
import services
import shell_open
import storage
from settings_controller import HOTKEY_HINTS


class SettingsPageController:
    """Страница «Настройки»: тема, папка документов, ключ DaData, копии.

    Состояние страницы — у контроллера (раздел Д плана); настройки
    читаются и пишутся через каркас (``app.settings`` /
    ``app.save_setting`` / ``app.set_theme``).
    """

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        self.page: QWidget | None = None
        self._build_page()
        self._install_page()

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

        title = QLabel("Настройки", page)
        title.setObjectName("pageTitle")
        subtitle = QLabel(
            "Тема оформления, папка документов, ключ DaData и копии реестра.",
            page,
        )
        subtitle.setObjectName("pageSubtitle")
        outer.addWidget(title)
        outer.addWidget(subtitle)

        scroll = QScrollArea(page)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        panel = QWidget(scroll)
        self._panel_layout = QVBoxLayout(panel)
        self._panel_layout.setContentsMargins(0, 0, 8, 0)
        self._panel_layout.setSpacing(12)
        scroll.setWidget(panel)
        outer.addWidget(scroll, 1)

        # -- Тема -------------------------------------------------------- #
        card = self._card("Тема оформления")
        theme_row = QHBoxLayout()
        self._theme_group = QButtonGroup(card)
        self._theme_light = QRadioButton("Светлая", card)
        self._theme_dark = QRadioButton("Тёмная", card)
        self._theme_group.addButton(self._theme_light)
        self._theme_group.addButton(self._theme_dark)
        current = qt_theme.current_mode()
        self._theme_light.setChecked(current == "light")
        self._theme_dark.setChecked(current == "dark")
        # Один обработчик на пару (группа исключающая): две подписки давали
        # бы двойной set_theme на каждое переключение.
        self._theme_light.toggled.connect(self._on_theme_toggled)
        # Внешняя смена темы (если появится до M7) синхронизирует индикаторы
        # без обратного вызова set_theme (сигналы заблокированы).
        app.register_theme_listener(self._sync_theme_indicators, owner=self.page)
        theme_row.addWidget(self._theme_light)
        theme_row.addWidget(self._theme_dark)
        theme_row.addStretch(1)
        card.grid.addLayout(theme_row)

        # -- Папка документов --------------------------------------------- #
        card = self._card("Папка для сохранения документов")
        self._documents_dir_label = QLabel(self._documents_dir(), card)
        self._documents_dir_label.setWordWrap(True)
        card.grid.addWidget(self._documents_dir_label)
        buttons = QHBoxLayout()
        choose_btn = QPushButton("Выбрать папку…", card)
        choose_btn.clicked.connect(self.choose_documents_dir)
        buttons.addWidget(choose_btn)
        open_btn = QPushButton("Открыть папку", card)
        open_btn.clicked.connect(self._open_documents_folder)
        buttons.addWidget(open_btn)
        buttons.addStretch(1)
        card.grid.addLayout(buttons)
        hint = QLabel(
            "Здесь сохраняются договоры и доп. соглашения, созданные в "
            "программе. Для каждого контрагента заводится своя папка "
            "с его документами.",
            card,
        )
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        card.grid.addWidget(hint)

        # -- DaData ------------------------------------------------------- #
        card = self._card("DaData: автозаполнение по ИНН")
        info = QLabel(
            "Ключ не входит в поставку — укажите свой ключ с dadata.ru. "
            "Без ключа доступен только ручной ввод реквизитов. При "
            "автозаполнении сервису отправляется только ИНН — публичный "
            "идентификатор организации/ИП; реквизиты карточки и документы "
            "никуда не отправляются.",
            card,
        )
        info.setObjectName("fieldHint")
        info.setWordWrap(True)
        card.grid.addWidget(info)
        self._dadata_key_entry = QLineEdit(card)
        self._dadata_key_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self._dadata_key_entry.setPlaceholderText("Новый ключ (не отображается)")
        card.grid.addWidget(self._dadata_key_entry)
        # Аудит 2.0.1: сохранённый ключ НЕ подставляется в поле — реальное
        # значение не показывается и не копируется. Состояние — строкой ниже.
        self._dadata_key_status = QLabel(self.dadata_key_status_text(), card)
        self._dadata_key_status.setObjectName("fieldHint")
        self._dadata_key_status.setWordWrap(True)
        card.grid.addWidget(self._dadata_key_status)
        key_buttons = QHBoxLayout()
        save_key_btn = QPushButton("Сохранить ключ", card)
        save_key_btn.clicked.connect(self.save_dadata_key)
        key_buttons.addWidget(save_key_btn)
        clear_key_btn = QPushButton("Удалить ключ", card)
        clear_key_btn.clicked.connect(self.clear_dadata_key)
        key_buttons.addWidget(clear_key_btn)
        key_buttons.addStretch(1)
        card.grid.addLayout(key_buttons)
        key_hint = QLabel(
            "Ключ хранится в профиле пользователя отдельно от реестра.",
            card,
        )
        key_hint.setObjectName("fieldHint")
        key_hint.setWordWrap(True)
        card.grid.addWidget(key_hint)

        # -- Копии реестра (BKUP-1) ---------------------------------------- #
        card = self._card("Копии реестра")
        self._backup_status_label = QLabel(storage.backup_status_text(), card)
        self._backup_status_label.setWordWrap(True)
        card.grid.addWidget(self._backup_status_label)
        backup_buttons = QHBoxLayout()
        backups_btn = QPushButton("Открыть папку копий", card)
        backups_btn.clicked.connect(self.open_backups_folder)
        backup_buttons.addWidget(backups_btn)
        backup_buttons.addStretch(1)
        card.grid.addLayout(backup_buttons)
        backup_hint = QLabel(
            "Копия базы создаётся при запуске, если реестр менялся с прошлой "
            "копии (содержимое — состояние на конец предыдущего прогона). "
            "Восстановление вручную: закройте программу, переименуйте файл "
            "копии registry_…db в database.db рядом с текущей базой "
            "(сначала сохранив текущий database.db). Ключ DaData в копии "
            "настроек не включается. Копия лежит на том же диске — от "
            "отказа диска она не защищает.",
            card,
        )
        backup_hint.setObjectName("fieldHint")
        backup_hint.setWordWrap(True)
        card.grid.addWidget(backup_hint)

        # -- Диагностика (WINLOG-1) ----------------------------------------- #
        card = self._card("Диагностика")
        diag_hint = QLabel(
            "В журнале видны события программы, медленные операции и сбои "
            "Win32-вызовов с именем операции. Журнал не содержит реквизитов "
            "и раскрывает путь только внутри профиля пользователя — его "
            "можно переслать целиком.",
            card,
        )
        diag_hint.setObjectName("fieldHint")
        diag_hint.setWordWrap(True)
        card.grid.addWidget(diag_hint)
        logs_buttons = QHBoxLayout()
        logs_btn = QPushButton("Открыть папку журналов", card)
        logs_btn.clicked.connect(self.open_logs_folder)
        logs_buttons.addWidget(logs_btn)
        logs_buttons.addStretch(1)
        card.grid.addLayout(logs_buttons)

        # -- Горячие клавиши (U14, UX-A) --------------------------------- #
        card = self._card("Горячие клавиши")
        for keys, description in HOTKEY_HINTS:
            line = QLabel(f"{keys}  {description}", card)
            line.setWordWrap(True)
            card.grid.addWidget(line)

        self._panel_layout.addStretch(1)

    def _card(self, caption: str) -> QFrame:
        """Карточка настроек: рамка + заголовок + вертикальная сетка."""
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
        self._panel_layout.addWidget(card)
        return card

    def _install_page(self) -> None:
        """Заменяет плейсхолдер M2 живой страницей «Настройки»."""
        app = self.app
        placeholder = app.pages.get("settings")
        index = app.stack.indexOf(placeholder) if placeholder is not None else -1
        if placeholder is not None:
            app.stack.removeWidget(placeholder)
            placeholder.deleteLater()
        app.stack.insertWidget(max(0, index), self.page)
        app.pages["settings"] = self.page
        if app.current_page == "settings":
            app.stack.setCurrentWidget(self.page)

    # ------------------------------------------------------------------ #
    #  Тема (через каркас: палитра + settings.json + слушатели)
    # ------------------------------------------------------------------ #
    def _sync_theme_indicators(self, mode: str) -> None:
        """Индикаторы темы следуют за применённой темой (внешние смены)."""
        for button in (self._theme_light, self._theme_dark):
            button.blockSignals(True)
        try:
            self._theme_light.setChecked(mode == "light")
            self._theme_dark.setChecked(mode == "dark")
        finally:
            for button in (self._theme_light, self._theme_dark):
                button.blockSignals(False)

    def _on_theme_toggled(self, _checked: bool = False) -> None:
        """Смена темы со страницы «Настройки» (mode: "light"/"dark")."""
        mode = "dark" if self._theme_dark.isChecked() else "light"
        normalized = self.app.set_theme(mode)
        if normalized != mode:
            # Нормализация каркаса не совпала (не должно случаться) —
            # вернуть индикаторы в фактическое состояние.
            self._theme_light.setChecked(normalized == "light")
            self._theme_dark.setChecked(normalized == "dark")

    # ------------------------------------------------------------------ #
    #  Папка документов: выбор и немедленное сохранение
    # ------------------------------------------------------------------ #
    def _documents_dir(self) -> str:
        """Текущая папка документов (настройка либо «Документы/EW.DocHub»).

        Non-interactive ветка: страница показывает путь честно, но не
        создаёт недоступные папки молча — выбор доступной папки
        выполняет choose_documents_dir.
        """
        custom = str((self.app.settings.get("documents_dir") or "")).strip()
        if custom:
            return custom
        try:
            return os.path.join(
                shell_open.documents_known_folder(), "EW.DocHub",
            )
        except Exception:  # noqa: BLE001 — показ пути не должен ронять страницу
            return "—"

    def choose_documents_dir(self) -> None:
        """Выбор папки документов с немедленным сохранением настройки."""
        app = self.app
        chosen = QFileDialog.getExistingDirectory(
            app.window,
            "Папка для сохранения договоров и документов",
            self._documents_dir(),
        )
        if not chosen:
            return
        # Немедленное сохранение (Tk: «при аварийном завершении выбранная
        # папка не теряется»); save_setting сам сообщает об отказе записи.
        if app.save_setting("documents_dir", chosen):
            self._documents_dir_label.setText(chosen)
            app.alert(
                app.window, "Папка сохранена",
                f"Документы будут сохраняться в:\n{chosen}", "success",
            )

    def _open_documents_folder(self) -> None:
        """«Открыть папку» — системное открытие текущей папки документов."""
        self.app.open_file(self._documents_dir())

    # ------------------------------------------------------------------ #
    #  Ключ DaData
    # ------------------------------------------------------------------ #
    def dadata_key_status_text(self) -> str:
        """Статус сохранённого ключа — без подстановки самого значения."""
        if dadata.load_api_key():
            if dadata.dpapi_available():
                return "Ключ сохранён и зашифрован (DPAPI, профиль пользователя)."
            return "Ключ сохранён в профиле пользователя."
        return "Ключ не задан."

    def save_dadata_key(self) -> None:
        """Сохраняет введённый пользователем ключ DaData."""
        app = self.app
        key = self._dadata_key_entry.text().strip()
        if not key:
            app.alert(app.window, "Ключ не указан",
                      "Введите API-ключ DaData или оставьте поле пустым.",
                      "error")
            return
        if dadata.save_api_key(key):
            self._dadata_key_entry.clear()
            self._dadata_key_status.setText(self.dadata_key_status_text())
            app.alert(
                app.window, "Ключ сохранён",
                "Автозаполнение по ИНН будет использовать этот ключ."
                + (" Значение зашифровано (DPAPI)." if dadata.dpapi_available()
                   else ""),
                "success",
            )
        else:
            app.alert(
                app.window, "Не удалось сохранить ключ",
                "Проверьте права на запись в профиль пользователя:\n"
                f"{dadata.key_storage_path()}", "error",
            )

    def clear_dadata_key(self) -> None:
        """Удаляет сохранённый ключ DaData."""
        app = self.app
        if not app.confirm(
            app.window, "Удалить ключ DaData",
            "Удалить сохранённый ключ? Автозаполнение по ИНН перестанет "
            "работать.",
            ok_text="Удалить", destructive=True,  # U7: потеря ключа
        ):
            return
        dadata.clear_api_key()
        self._dadata_key_entry.clear()
        self._dadata_key_status.setText(self.dadata_key_status_text())
        app.alert(app.window, "Ключ удалён",
                  "Ключ DaData удалён из профиля пользователя.", "success")

    # ------------------------------------------------------------------ #
    #  Копии реестра (BKUP-1) и диагностика
    # ------------------------------------------------------------------ #
    def update_backup_status(self, result: "dict | None" = None) -> None:
        """Обновляет статус копий (после стартовой автокопии)."""
        try:
            if result is not None and result.get("status") == "failed":
                self._backup_status_label.setText(
                    f"Последняя копия не удалась: {result['message']}"
                )
            else:
                self._backup_status_label.setText(storage.backup_status_text())
        except RuntimeError:
            pass  # страница могла быть разрушена

    def open_backups_folder(self) -> None:
        """Открывает папку автокопий (создаёт её, если копий ещё не было)."""
        try:
            os.makedirs(storage.backups_dir(), exist_ok=True)
        except OSError:
            pass
        self.app.open_file(storage.backups_dir())

    def open_logs_folder(self) -> None:
        """Открывает папку журналов app.log/crash.log/perf.log (WINLOG-1).

        Папка создаётся при отсутствии: кнопка работает и на свежей
        установке, где журналов ещё нет.
        """
        try:
            os.makedirs(applog.log_dir(), exist_ok=True)
        except OSError:
            pass
        self.app.open_file(applog.log_dir())
