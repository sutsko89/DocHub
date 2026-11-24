"""Точка входа и каркас Qt-приложения EW.DocHub (этапы M2–M3 миграции Tk → PySide6).

Каркас по плану миграции (M2): QApplication, главное окно со
стандартной системной рамкой (нативные заголовок, «свернуть/развернуть/
закрыть», перенос, растягивание, прилипание и унос на второй монитор —
средствами Windows), панель разделов, футер со строкой статуса,
немодальный тост, модальные диалоги (qt_dialogs), темы light/dark
(qt_theme), журналы app.log и crash.log (faulthandler), защита от
второго экземпляра (single_instance, те же мьютексы, что у Tk-сборки),
загрузка settings.json (прежний ключ ``theme`` — профиль совместим с
Tk-сборкой), владелец DataStore, QTimer-инфраструктура опроса очередей
и perf-каналы ``startup.built`` / ``startup.ready``.

Этап M3 (контрольная точка CP1): живая страница «Спецификация»
(spec_page_qt.SpecPageController — заменяет плейсхолдер), Ctrl+S,
защита несохранённых данных при уходе со страницы и закрытии окна
(подтверждение, сбой чтения подписи = «грязная», п. 4.4), shell-хелперы
через GUI-free shell_open (точка M3 из докстринга M2).

Контракты из раздела Д плана, переносимые без изменений:

* любое касание виджетов — только главный поток Qt; фоновые воркеры
  (появятся в M4+ на месте Tk-контроллеров) кладут результаты в
  ``queue.Queue``, опрос — QTimer, не QThread («переписать заодно»
  запрещено планом, риск 6);
* владелец общего состояния — этот каркас (роль ``ContractApp`` в
  Tk-сборке): навигация, защита несохранённых, текущий контрагент;
  данные страниц — у их контроллеров; один ``DataStore`` на приложение,
  вызовы store — только из GUI-потока (STORE-1A);
* после CP1 реестр, карточка, импорт и настройки в Qt-сборке по-прежнему
  отсутствуют — разделы видны в панели, но сообщают «в разработке».

На Linux без X-сервера до импорта PySide6 выставляется платформа
``offscreen`` — каркасные тесты и smoke идут без дисплея, новые skip
для этого не вводятся.
"""

from __future__ import annotations

import os
import sys

# Linux/macOS без дисплея: платформа offscreen до создания QApplication.
# На Windows ничего не переопределяется — там каркас идёт с нативной рамкой.
if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QFrame, QHBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QMainWindow, QPushButton, QStackedWidget, QVBoxLayout, QWidget,
)

import app_common
import applog
import counterparty_page_qt
import dadata_qt
import lifecycle_qt
import perflog
import qt_dialogs
import qt_theme
import registry_page_qt
import shell_open
import single_instance
import storage
from services import resource_path
import spec_page_qt

# Признак интерфейса этой сборки (план И.2): app_common.APP_UI_FLAVOR остаётся
# "tk" для Tk-сборки; Qt-сборка объявляет себя здесь. Отображение версии —
# «2.28.2-qt»; сравнение версий updater идёт по числовой части (M7).
UI_FLAVOR = "qt"

# Геометрия окна Qt-каркаса: значения Tk-сборки (BASE_*/MIN_WINDOW_* из
# main.py) повторены здесь, чтобы каркас выглядел соразмерно. DPI
# масштабируется Qt нативно (К-4 — проверка на Windows, этап M2).
BASE_WIDTH = 1280
BASE_HEIGHT = 740
MIN_WINDOW_WIDTH = 1120
MIN_WINDOW_HEIGHT = 620

DEFAULT_PAGE = "spec"  # живая страница CP1 (M3); остальные — плейсхолдеры

# Разделы панели: те же ключи, что у Tk NAV_PAGES (main.py), плюс этап
# миграции, на котором раздел станет живым. Подписи — из Tk-версии.
NAV_PAGES = (
    ("counterparty", "Добавить контрагента", "M4"),
    ("contract", "Договор поставки", "M6"),
    ("additional", "Доп. соглашение", "M6"),
    ("spec", "Спецификация", "M3"),
    ("registry", "Реестр", "M5"),
    ("settings", "Настройки", "M6"),
)

TOAST_TIMEOUT_MS = 9000  # как в Tk-версии (main._notify)
STORE_WARNING_DELAY_MS = 300  # алерт после построения окна, как в Tk


class MainWindow(QMainWindow):
    """Главное окно со стандартной рамкой; closeEvent передаётся каркасу."""

    def __init__(self, controller: "AppController"):
        super().__init__()
        self._controller = controller

    def closeEvent(self, event) -> None:  # noqa: N802 — имя Qt
        self._controller._on_close(event)


class AppController:
    """Владелец общего состояния Qt-сборки (роль ContractApp в Tk-сборке).

    Создаётся поверх уже существующего ``QApplication`` (точка входа
    :func:`main`, тесты — fixture ``qapp`` pytest-qt). Строит окно,
    панель разделов, страницы-плейсхолдеры, футер и тост; владеет
    ``DataStore``, таймерами опроса и настройками.
    """

    def __init__(self, app: QApplication):
        self.app = app
        self.log = applog.get_logger("main_qt")
        self._settings = app_common.load_settings()
        self._theme_listeners: list = []
        # Единый текущий контрагент (п. 4.2; Tk: _current_party_id) —
        # общий контекст будущих страниц документов (M4+).
        self.current_party_id = None

        qt_theme.apply_theme(app, self._settings.get("theme", qt_theme.DEFAULT_THEME))

        # Один DataStore на приложение, владелец — каркас (STORE-1A).
        self.store = storage.DataStore()
        if self.store.load_warning:
            # Сообщаем после построения окна — алерт поверх интерфейса.
            QTimer.singleShot(
                STORE_WARNING_DELAY_MS, self._show_store_warning,
            )

        self.window = MainWindow(self)
        self.window.setWindowTitle(single_instance.WINDOW_TITLE)
        self.window.resize(BASE_WIDTH, BASE_HEIGHT)
        self.window.setMinimumSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
        icon = QIcon(resource_path("app_icon.ico"))
        if not icon.isNull():
            self.window.setWindowIcon(icon)

        self._timers: dict[str, QTimer] = {}
        self._toast_timer: QTimer | None = None
        self._toast_path: str | None = None
        self.current_page = DEFAULT_PAGE

        self._build_body()

        # Живая страница CP1 (M3): SpecPageController заменяет плейсхолдер
        # «Спецификации» и владеет её состоянием (раздел Д плана).
        self.spec_flow = spec_page_qt.SpecPageController(self)

        # Живая страница M4: «Добавить контрагента» + окно правки карточки;
        # поток «Заполнить по ИНН» — один на приложение (как Tk _dadata_flow).
        self.counterparty_flow = counterparty_page_qt.CounterpartyPageController(self)
        self.dadata_flow = dadata_qt.DadataFlowQt(self)

        # Живые страницы M5: реестр (TREE-1) и поток жизненного цикла записей
        # (экспорт Word, корзина) — Qt-аналог Tk _lifecycle_flow.
        self.lifecycle_flow = lifecycle_qt.LifecycleFlowQt(self)
        self.registry_flow = registry_page_qt.RegistryPageController(self)

        # Ctrl+S — сохранение формы текущей страницы (п. 4.4, Tk-контракт).
        save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self.window)
        save_shortcut.activated.connect(self._on_save_shortcut)

        self.log.info("Каркас Qt построен (%s-%s)", app_common.APP_VERSION, UI_FLAVOR)

    @property
    def settings(self) -> dict:
        """Настройки профиля (settings.json) — чтение для страниц."""
        return self._settings

    # ------------------------------------------------------------------ #
    #  Построение каркаса: сайдбар, страницы-плейсхолдеры, футер, тост
    # ------------------------------------------------------------------ #
    def _build_body(self) -> None:
        central = QWidget(self.window)
        central.setObjectName("centralBody")
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.nav = QListWidget(central)
        self.nav.setObjectName("navPanel")
        self.nav.setFixedWidth(210)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for key, title, _stage in NAV_PAGES:
            item = QListWidgetItem(title)
            item.setData(Qt.ItemDataRole.UserRole, key)
            self.nav.addItem(item)
        self.nav.currentItemChanged.connect(self._on_nav_changed)
        layout.addWidget(self.nav)

        self.stack = QStackedWidget(central)
        layout.addWidget(self.stack, 1)
        self.pages: dict[str, QWidget] = {}
        for key, title, stage in NAV_PAGES:
            self.pages[key] = self._build_placeholder(title, stage)
            self.stack.addWidget(self.pages[key])

        footer = QFrame(central)
        footer.setObjectName("footer")
        footer.setFixedHeight(48)
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(20, 8, 20, 8)
        self.status_label = QLabel(footer)
        self.status_label.setObjectName("statusLabel")
        footer_layout.addWidget(self.status_label, 1)
        layout.addWidget(footer)

        self.toast = QFrame(central)
        self.toast.setObjectName("toast")
        self.toast.setFixedSize(360, 64)
        toast_layout = QHBoxLayout(self.toast)
        toast_layout.setContentsMargins(14, 8, 10, 8)
        self.toast_label = QLabel(self.toast)
        self.toast_label.setWordWrap(True)
        toast_layout.addWidget(self.toast_label, 1)
        self.toast_open = QPushButton("Открыть", self.toast)
        self.toast_open.clicked.connect(self._toast_open)
        self.toast_folder = QPushButton("Показать в папке", self.toast)
        self.toast_folder.clicked.connect(self._toast_show_folder)
        toast_layout.addWidget(self.toast_folder)
        toast_layout.addWidget(self.toast_open)
        self.toast.hide()
        self._apply_toast_actions()

        self.window.setCentralWidget(central)
        self.refresh_status()
        self._select_page(self.current_page)

    def _build_placeholder(self, title: str, stage: str) -> QWidget:
        """Страница «в разработке»: навигационная модель не теряется."""
        page = QWidget(self.stack)
        page_layout = QVBoxLayout(page)
        page_layout.addStretch(1)
        label = QLabel(f"Раздел «{title}» в разработке", page)
        label.setObjectName("pagePlaceholder")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_layout.addWidget(label)
        hint = QLabel(
            f"Появится на этапе {stage} миграции Tk → PySide6", page,
        )
        hint.setObjectName("pagePlaceholderHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        page_layout.addWidget(hint)
        page_layout.addStretch(2)
        return page

    # ------------------------------------------------------------------ #
    #  Навигация и защита несохранённых данных (п. 4.4, M3)
    # ------------------------------------------------------------------ #
    def _on_nav_changed(self, current: QListWidgetItem, previous) -> None:
        if current is None:
            return
        key = current.data(Qt.ItemDataRole.UserRole)
        if key not in self.pages:
            return
        if not self.confirm_leave_page():
            # Отказ — остаёмся на прежнем разделе: выделение откатывается
            # без повторного срабатывания guard'а.
            if previous is not None:
                self.nav.blockSignals(True)
                self.nav.setCurrentItem(previous)
                self.nav.blockSignals(False)
            return
        self.current_page = key
        self.stack.setCurrentWidget(self.pages[key])
        if key == "spec" and getattr(self, "spec_flow", None):
            # UPD-UI (P-07): предпросмотр и итоги пересобираются при
            # переключении на видимую страницу.
            self.spec_flow.on_page_open()
        if key == "registry" and getattr(self, "registry_flow", None):
            # 2.19.0: провал в документы не переживает уход со страницы;
            # ленивый рендер реестра — при первом открытии.
            self.registry_flow.on_page_open()

    def select_page(self, key: str) -> None:
        """Переключение раздела по ключу (для клавиатуры и тестов).

        Уход со страницы с несохранёнными данными — через подтверждение
        (п. 4.4); при отказе раздел не меняется.
        """
        if key not in self.pages:
            return
        if key != self.current_page and not self.confirm_leave_page():
            return
        self.current_page = key
        self.stack.setCurrentWidget(self.pages[key])
        for row in range(self.nav.count()):
            item = self.nav.item(row)
            if item.data(Qt.ItemDataRole.UserRole) == key:
                self.nav.blockSignals(True)
                self.nav.setCurrentItem(item)
                self.nav.blockSignals(False)
                break
        if key == "spec" and getattr(self, "spec_flow", None):
            self.spec_flow.on_page_open()
        if key == "registry" and getattr(self, "registry_flow", None):
            self.registry_flow.on_page_open()

    # Синоним с именем Tk-каркаса (привычные имена для будущих страниц).
    _select_page = select_page

    def is_current_page_dirty(self) -> bool:
        """Есть ли несохранённые данные на текущей странице (п. 4.4).

        Живые страницы: «Спецификация» (CP1) и «Добавить контрагента»
        (M4); M5+ расширяет словарь. Сбой чтения подписи контроллер
        считает «грязной» страницей.
        """
        if self.current_page == "spec" and getattr(self, "spec_flow", None):
            try:
                return self.spec_flow.is_form_dirty()
            except Exception:
                return True
        if self.current_page == "counterparty" and getattr(
            self, "counterparty_flow", None,
        ):
            try:
                return self.counterparty_flow.is_form_dirty()
            except Exception:
                return True
        return False

    def confirm_leave_page(self) -> bool:
        """Уход со страницы с изменениями — только через подтверждение.

        Контракт Tk confirm_leave_page (document_context_controller):
        без изменений — True без диалога; с изменениями — confirm с
        деструктивной семантикой (U7: Enter не подтверждает потерю).
        """
        if not self.is_current_page_dirty():
            return True
        key = self.current_page
        label = {
            "spec": "Спецификация",
            # Tk-подпись document_context_controller.form_label (п. 4.4).
            "counterparty": "Карточка контрагента",
        }.get(key, key)
        return self.confirm(
            self.window, "Несохранённые изменения",
            f"На странице «{label}» есть несохранённые данные.\n\n"
            "Перейти и потерять изменения?",
            ok_text="Перейти", cancel_text="Остаться",
            destructive=True,
        )

    def save_current_page(self) -> None:
        """Ctrl+S: сохранение формы текущей страницы (п. 4.4)."""
        if self.current_page == "spec" and getattr(self, "spec_flow", None):
            self.spec_flow.save()
        if self.current_page == "counterparty" and getattr(
            self, "counterparty_flow", None,
        ):
            self.counterparty_flow.save()

    # ------------------------------------------------------------------ #
    #  Единый текущий контрагент (п. 4.2; Tk set_current_party)
    # ------------------------------------------------------------------ #
    def set_current_party(self, rec: dict) -> None:
        """Единый текущий контрагент для всех страниц (Tk _current_party_id).

        Делегирует странице «Спецификация» (список и панель «Документ
        для»); M5 реестр и M6 договор присоединятся к тому же контракту.
        Список контрагентов освежается до выбора: карточка, добавленная
        или выбранная по дубликату ИНН (п. 4.1), обязана быть в комбо —
        иначе контекстный выбор честно очистится (п. 4.2) — как в Tk
        refresh_party_combo_values внутри set_current_party.
        """
        if getattr(self, "spec_flow", None):
            self.spec_flow.refresh_parties()
            self.spec_flow.set_current_party(rec)
        else:
            self.current_party_id = rec.get("id")

    def refresh_after_party_change(self) -> None:
        """Обновление зависимых списков после добавления/правки карточки.

        Tk: refresh_party_combo_values + _update_status + реестр (M5):
        с M5 реестр живой и обновляется вместе с комбо страниц документов.
        """
        if getattr(self, "spec_flow", None):
            self.spec_flow.refresh_parties()
        if getattr(self, "registry_flow", None):
            self.registry_flow.refresh()
        self.refresh_status()

    def save_setting(self, key: str, value) -> None:
        """Сохранение настройки профиля (Tk ContractApp.save_setting).

        Атомарная запись settings.json (app_common.write_settings) — тот
        же файл, что у Tk-сборки; отказ записи сообщается алертом,
        значение в сеансе остаётся применённым.
        """
        self._settings[key] = value
        if not app_common.write_settings(self._settings):
            self.alert(
                self.window, "Настройка не сохранена",
                f"Не удалось записать файл:\n{app_common.settings_path()}",
                "error",
            )

    def _on_save_shortcut(self) -> None:
        self.save_current_page()

    # ------------------------------------------------------------------ #
    #  Статус и тост
    # ------------------------------------------------------------------ #
    def refresh_status(self) -> None:
        """Строка футера: сколько контрагентов и документов в реестре.

        Формат и источник совпадают с Tk (main._update_status): счётчики
        через store.status_counts() (STORE-1A, три COUNT(*)).
        """
        counts = self.store.status_counts()
        status = (
            f"Контрагентов: {counts['parties']} · "
            f"Документов: {counts['documents']} (договоров: {counts['contracts']})"
        )
        if counts["trash"]:
            status += f" · В корзине: {counts['trash']}"
        self.status_label.setText(status)

    def notify(self, message: str, path: str | None = None) -> None:
        """Немодальное уведомление об успехе (замена Tk-тоста main._notify).

        Не блокирует работу; при ``path`` предлагает «Открыть» и
        «Показать в папке». Автоскрытие через 9 секунд, повторный вызов
        перезапускает таймер.
        """
        self._toast_path = path or None
        self.toast_label.setText(message)
        self._apply_toast_actions()
        self.toast.adjustSize()
        central = self.window.centralWidget()
        if central is not None:
            x = max(0, central.width() - self.toast.width() - 20)
            y = max(0, central.height() - self.toast.height() - 20)
            self.toast.move(x, y)
        self.toast.show()
        self.toast.raise_()
        if self._toast_timer is None:
            self._toast_timer = QTimer(self.toast)
            self._toast_timer.setSingleShot(True)
            self._toast_timer.timeout.connect(self._hide_toast)
        self._toast_timer.start(TOAST_TIMEOUT_MS)

    def _apply_toast_actions(self) -> None:
        has_path = bool(self._toast_path)
        self.toast_open.setVisible(has_path)
        self.toast_folder.setVisible(has_path)

    def _hide_toast(self) -> None:
        self.toast.hide()

    def _toast_open(self) -> None:
        if self._toast_path:
            self._open_path(self._toast_path)

    def _toast_show_folder(self) -> None:
        if self._toast_path:
            self._show_path(self._toast_path)

    def open_file(self, file_path: str) -> bool:
        """Открывает файл системным приложением; честный отказ без файла.

        Точка M3 (докстринг M2): GUI-free ядро — shell_open.open_file_system,
        общий для страниц (CP1 «Открыть файл»); алерты — qt. Контракт
        main.open_file (п. 6.2): absent — сразу отказ с алертом, без
        ложного «открыт».
        """
        if not file_path or not os.path.exists(file_path):
            self.alert(
                self.window, "Файл не найден",
                f"Файл не найден на диске:\n{file_path or '—'}", "error",
            )
            return False
        if shell_open.open_file_system(file_path):
            return True
        self.alert(
            self.window, "Не удалось открыть файл",
            "Система не смогла открыть файл:\n"
            f"{file_path}\n\nПодробности — в app.log.", "error",
        )
        return False

    def show_in_folder(self, file_path: str) -> None:
        """Показ файла в проводнике без запуска (контракт main.show_in_folder)."""
        if not file_path or not os.path.exists(file_path):
            self.alert(
                self.window, "Файл не найден",
                file_path or "Путь к файлу не указан.", "error",
            )
            return
        if not shell_open.show_in_file_manager(file_path):
            self.alert(
                self.window, "Не удалось открыть папку",
                "Система не смогла показать папку:\n"
                f"{file_path}\n\nПодробности — в app.log.", "error",
            )

    # Имена из M2 (докстринг «Собственная копия до M3»): хелперы объединены
    # на GUI-free shell_open, имена сохранены для тоста и тестов.
    _open_path = open_file
    _show_path = show_in_folder

    # ------------------------------------------------------------------ #
    #  Диалоги — замены ui.alert/ui.confirm/ui.choice_dialog
    # ------------------------------------------------------------------ #
    def alert(self, parent, title: str, message: str, kind: str = "info",
              ok_text: str = "ОК") -> None:
        qt_dialogs.alert(parent, title, message, kind, ok_text)

    def confirm(self, parent, title: str, message: str, ok_text: str = "ОК",
                cancel_text: str = "Отмена", destructive: bool = False) -> bool:
        return qt_dialogs.confirm(
            parent, title, message, ok_text, cancel_text, destructive,
        )

    def choice(self, parent, title: str, message: str, options) -> str | None:
        return qt_dialogs.choice(parent, title, message, options)

    def _show_store_warning(self) -> None:
        self.alert(
            self.window, "Реестр загружен не полностью",
            self.store.load_warning, "error",
        )

    # ------------------------------------------------------------------ #
    #  Темы (ключ settings.json прежний — профиль совместим с Tk)
    # ------------------------------------------------------------------ #
    def set_theme(self, mode: str) -> str:
        """Переключение темы: палитра/QSS + settings.json + слушатели.

        Возвращает нормализованный режим. Сохранение — то же атомарное
        app_common.write_settings, что в Tk-сборке; отказ записи
        сообщается алертом, тема в сеансе остаётся применённой.
        """
        normalized = qt_theme.apply_theme(self.app, mode)
        self._settings["theme"] = normalized
        if not app_common.write_settings(self._settings):
            self.alert(
                self.window, "Настройки не сохранены",
                f"Не удалось записать файл:\n{app_common.settings_path()}", "error",
            )
        for callback in list(self._theme_listeners):
            try:
                callback(normalized)
            except Exception:  # noqa: BLE001 — сбой слушателя не роняет тему
                self.log.exception("Сбой слушателя темы")
        return normalized

    def register_theme_listener(self, callback, owner=None) -> None:
        """Подписка на смену темы (LEAK-1/D8: автоотписка по владельцу).

        ``owner`` — QObject (или None для вечных подписок): по
        ``destroyed`` подписка снимается сама, накапливания мёртвых
        замыканий нет (контракт Tk-версии).
        """
        self._theme_listeners.append(callback)
        if owner is not None and hasattr(owner, "destroyed"):
            owner.destroyed.connect(
                lambda _obj=None, cb=callback: self.unregister_theme_listener(cb),
            )

    def unregister_theme_listener(self, callback) -> None:
        try:
            self._theme_listeners.remove(callback)
        except ValueError:
            pass

    # ------------------------------------------------------------------ #
    #  QTimer-инфраструктура опроса (замена Tk after()-точек, В.4)
    # ------------------------------------------------------------------ #
    def add_poll_timer(self, name: str, interval_ms: int, callback) -> QTimer:
        """Периодический опрос очередей контроллера в главном потоке.

        Контракт раздела Д плана: воркер не трогает UI — результат
        кладётся в queue.Queue, опрос таймером применяет его здесь.
        Имена однотоковые латинские (по шаблону perf-каналов).
        Повторная регистрация с тем же именем замещает старый таймер
        (останов + отсоединение), накапливания таймеров нет.
        """
        timer = QTimer(self.window)
        timer.setObjectName(name)
        timer.setInterval(max(1, int(interval_ms)))
        timer.timeout.connect(callback)
        timer.start()
        old = self._timers.get(name)
        if old is not None:
            old.stop()
            old.timeout.disconnect()
        self._timers[name] = timer
        return timer

    def call_later(self, delay_ms: int, callback) -> None:
        """Одноразовый отложенный вызов в главном потоке (замена after())."""
        QTimer.singleShot(max(0, int(delay_ms)), callback)

    def stop_poll_timer(self, name: str) -> None:
        timer = self._timers.pop(name, None)
        if timer is not None:
            timer.stop()

    def _stop_timers(self) -> None:
        """Полный останов опроса (закрытие окна, дренаж перед выходом)."""
        for timer in self._timers.values():
            try:
                timer.stop()
            except RuntimeError:
                pass  # таймер уже уничтожен вместе с окном
        self._timers.clear()

    # ------------------------------------------------------------------ #
    #  Perf-каналы старта (E-4; имена парсит bench_store_1a.py --report)
    # ------------------------------------------------------------------ #
    def _emit_startup_perf(self) -> None:
        """``startup.built`` — до входа в цикл, ``ready`` — первый проход.

        Без EWDOCHUB_PERF_LOG запись — мгновенный no-op (perflog.record).
        """
        perflog.record("startup.built", 0.0, kind="event")
        QTimer.singleShot(0, self._record_ready)

    def _record_ready(self) -> None:
        perflog.record("startup.ready", 0.0, kind="event")

    # ------------------------------------------------------------------ #
    #  Запуск и закрытие
    # ------------------------------------------------------------------ #
    def run(self) -> int:
        """Показ окна и цикл событий. Возвращает код выхода QApplication."""
        self.window.show()
        self._emit_startup_perf()
        return self.app.exec()

    def _on_close(self, event) -> None:
        """Закрытие окна: защита несохранённых (M3), останов опроса, журнал.

        «Грязная» форма — только через подтверждение (п. 4.4, Tk-контракт
        защиты при закрытии); отказ — окно остаётся открытым
        (event.ignore). Подтверждение деструктивное: Enter не теряет
        данные (U7).
        """
        if not self._confirm_close_dirty():
            event.ignore()
            return
        self._stop_timers()
        if self.spec_flow.preview_timer is not None:
            self.spec_flow.preview_timer.stop()
        # Активный запрос DaData: сброс цели/очереди — результат в полёте
        # не должен применяться к уничтоженным полям (тихий отказ U8).
        if getattr(self, "dadata_flow", None) is not None:
            self.dadata_flow.abort()
        if self._toast_timer is not None:
            self._toast_timer.stop()
        self.log.info("Каркас Qt закрыт")
        perflog.flush()
        event.accept()

    def _confirm_close_dirty(self) -> bool:
        """Подтверждение закрытия с несохранёнными данными (п. 4.4).

        Подпись страницы — по разделу (исправление M4-хвоста: у закрытия
        с «грязной» карточкой текст говорил о «Спецификации»).
        """
        if not self.is_current_page_dirty():
            return True
        label = {
            "spec": "Спецификация",
            "counterparty": "Карточка контрагента",
        }.get(self.current_page, "текущей странице")
        return self.confirm(
            self.window, "Несохранённые изменения",
            f"На странице «{label}» есть несохранённые данные.\n\n"
            "Закрыть окно и потерять изменения?",
            ok_text="Закрыть", cancel_text="Остаться",
            destructive=True,
        )


def main() -> int:
    """Точка входа Qt-сборки: один экземпляр, логирование, освобождение блокировки.

    Порядок повторяет main.py::main (Tk): блокировка экземпляра до
    открытия интерфейса — второй процесс писал бы в ту же базу; далее
    журналы, perf-флаг, faulthandler, QApplication, каркас.
    """
    lock = single_instance.InstanceLock()
    if not lock.acquire():
        single_instance.activate_existing_window()
        single_instance.show_already_running()
        return 0

    applog.setup_logging(f"{app_common.APP_VERSION}-{UI_FLAVOR}")
    if perflog.enabled():
        # SEC-1 (D12): путь журнала маскируется — не раскрываем профиль.
        applog.get_logger("perf").info(
            "Perf-лог включён (EWDOCHUB_PERF_LOG): %s",
            applog.mask_path(perflog.perf_path()),
        )
    # Дамп критических сбоев: хард-падения нативного стека Qt (как и Tcl/Tk
    # в Tk-сборке) умирают без Python-traceback — C-стек идёт в crash.log.
    crash_log = applog.enable_crash_dump()
    log = applog.get_logger("main_qt")
    if crash_log:
        log.info("Дамп критических сбоев включён: %s", applog.mask_path(crash_log))

    exit_code = 0
    try:
        qt_app = QApplication.instance() or QApplication(sys.argv)
        qt_app.setApplicationName("EW.DocHub")
        qt_app.setApplicationVersion(app_common.APP_VERSION)
        controller = AppController(qt_app)
        exit_code = controller.run()
    except Exception:
        log.exception("Необработанная ошибка при работе приложения")
        raise
    finally:
        lock.release()
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
