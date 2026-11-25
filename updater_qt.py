"""Проверка обновлений и автообновление на Qt (этап M7) — Г #17.

Qt-перенос ``updater_controller.UpdateController`` (Tk-контракты 2.15.0 /
2.18.0 / 2.19.2 сохраняются дословно):

* тихая автопроверка при старте (раз в сутки, если включена пользователем
  и не заглушена средой) — ``maybe_auto_check`` вызывается стартовым
  таймером каркаса ``QTimer.singleShot(2200, …)`` — точка Tk
  ``root.after(2200, …)``;
* ручная проверка кнопкой «Проверить обновления» на странице «Настройки»
  — ``check_now`` с guard повторных кликов и блокировкой кнопки;
* диалог результата: «Обновить сейчас» / «Открыть страницу» / «Позже»;
  при «Обновить сейчас» — окно скачивания с прогрессом и отменой,
  контроль SHA-256 по SHA256SUMS.txt релиза, подтверждение, тихий запуск
  установщика Inno Setup и закрытие приложения;
* честное «версия последняя» / ошибка сети при ручной проверке;
  автопроверка молчит, если обновления нет или сеть недоступна.

Отличие от Tk-контроллера (И.2 плана миграции): ассет релиза выбирается
ПО ФЛЕЙВОРУ сборки — Qt-сборка скачивает только ``*-qt-setup.exe``,
чужой Tk-установщик не скачивается; при отсутствии своего ассета
предлагается страница релиза (CI мог ещё не выложить ассет). Статус
версии на «Настройках» — с суффиксом флейвора («2.28.2-qt»).

Паттерн потоков — контракт раздела Д плана (как dadata_qt/import_qt):
HTTP в рабочем потоке, результат через ``queue.Queue``, опрос из главного
потока именованными poll-таймерами каркаса (QTimer, не QThread-сигналы —
«переписать заодно на QThread» без измеримой необходимости запрещено).
Фактическая сетевая работа и запуск установщика — в GUI-free updater.py;
контроллер только маршрутизирует результат в интерфейс.

Подменяемые маршруты (для smoke_qt и юнит-тестов):

* ``updater.check_for_update`` / ``updater.download_file`` /
  ``updater.fetch_text`` — сетевые функции подменяются целиком;
* ``threading.Thread`` — через фабрику потока (синхронный прогон);
* ``webbrowser.open`` — открытие страницы релиза подменяется в тестах;
* ``updater.run_installer_command`` — модульная функция запуска setup.exe;
* ``app.alert`` / ``app.confirm`` / ``app.choice`` — маршруты каркаса.

Модуль не импортирует updater_controller (там импорт tkinter, который в
Qt-сборке исключён): делегируемая функция живёт в updater.py с M7.
"""

from __future__ import annotations

import queue
import threading
import webbrowser
from typing import Any, Callable

import app_common
import applog
import updater
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
)

# Интервалы опроса очередей (Tk-паритет: after(80, …) в обоих циклах).
RESULT_INTERVAL_MS = 80
DOWNLOAD_INTERVAL_MS = 80

# Имена poll-таймеров каркаса (однотоковые латинские — perf-гигиена).
CHECK_TIMER = "update_check"
DOWNLOAD_TIMER = "update_download"


def _default_thread_factory() -> Callable:
    """Фабрика потока разрешается в момент вызова (подмена в тестах)."""
    return threading.Thread


class UpdateProgressDialogQt(QDialog):
    """Окно скачивания обновления: заголовок, статус, прогресс, «Отмена».

    Замена CTkToplevel Tk-контроллера. Не модальное (как Tk-окно):
    пользователь может свернуть программу, скачивание продолжается.
    Закрытие окна любым способом (крестик, Alt+F4) = отмена — контракт
    Tk ``WM_DELETE_WINDOW → cancel_download``.
    """

    def __init__(self, parent, release: dict, on_cancel: Callable) -> None:
        super().__init__(parent)
        self._on_cancel = on_cancel
        newest = str((release or {}).get("version_str")
                     or (release or {}).get("tag") or "")
        self.setWindowTitle("Обновление EW.DocHub")
        self.setModal(False)
        self.setFixedSize(430, 150)
        body = QVBoxLayout(self)
        body.setContentsMargins(24, 20, 24, 16)
        body.setSpacing(8)
        title = QLabel(f"Скачивание обновления {newest}")
        title.setStyleSheet("font-weight: 600;")
        body.addWidget(title)
        self.status_label = QLabel("Подключение к github.com…")
        self.status_label.setWordWrap(True)
        body.addWidget(self.status_label)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        body.addWidget(self.progress)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Отмена")
        cancel.clicked.connect(self.cancel)
        buttons.addWidget(cancel)
        body.addLayout(buttons)

    def set_status(self, text: str) -> None:
        """Строка статуса (фаза подключения/скачивания/проверки)."""
        self.status_label.setText(text)

    def set_progress(self, done: int, total: int) -> None:
        """Полоса прогресса и статус «X МБ из Y МБ» (Tk-паритет текста)."""
        if total > 0:
            self.progress.setRange(0, 100)
            self.progress.setValue(int(min(1.0, done / total) * 100))
            mb_done = done / (1024 * 1024)
            mb_total = total / (1024 * 1024)
            if mb_total >= 1:
                self.set_status(f"{mb_done:.1f} МБ из {mb_total:.1f} МБ")
            else:
                self.set_status(f"{done / 1024:.0f} КБ из {total / 1024:.0f} КБ")
        else:
            self.progress.setRange(0, 0)  # неопределённый режим
            self.set_status(f"Скачано: {done / (1024 * 1024):.1f} МБ")

    def cancel(self) -> None:
        """Кнопка «Отмена» / закрытие окна: остановка и закрытие."""
        self._on_cancel()
        self.close()

    def closeEvent(self, event) -> None:  # noqa: N802 — имя Qt
        # Крестик окна = «Отмена» (Tk-контракт WM_DELETE_WINDOW);
        # cancel идемпотентен: повторная установка события безвредна.
        self._on_cancel()
        super().closeEvent(event)


class UpdateControllerQt:
    """Жизненный цикл обновления: guard → проверка → диалог → скачивание → установка.

    Владелец состояния — главный поток Qt. Рабочие потоки кладут результат
    в очереди и не трогают интерфейс; устаревших результатов нет: каждая
    проверка/скачивание имеет свою очередь, терминальное событие гасит
    таймер опроса.
    """

    def __init__(self, app: Any, flavor: str | None = None,
                 thread_factory: Callable | None = None) -> None:
        self.app = app
        self.busy = False              # guard повторных проверок
        self.result_queue: "queue.Queue | None" = None
        # Скачивание: своя очередь, окно прогресса и флаг отмены.
        self.download_queue: "queue.Queue | None" = None
        self.cancel_event: "threading.Event | None" = None
        self.download_window: "UpdateProgressDialogQt | None" = None
        self._flavor = flavor          # None → лениво из каркаса (UI_FLAVOR)
        self._thread_factory = thread_factory

    # ------------------------------------------------------------------ #
    #  Флейвор сборки (И.2): «qt» — выбор ассета и суффикс версии
    # ------------------------------------------------------------------ #
    @property
    def flavor(self) -> str:
        if self._flavor is None:
            import main_qt  # поздний импорт: каркас импортирует этот модуль

            self._flavor = getattr(
                main_qt, "UI_FLAVOR", app_common.APP_UI_FLAVOR)
        return self._flavor

    # ------------------------------------------------------------------ #
    #  Запуск проверки
    # ------------------------------------------------------------------ #
    def maybe_auto_check(self) -> None:
        """Стартовый таймер каркаса: тихая проверка, если положено.

        Три независимых рубильника: env-глушилка (тесты/офлайн), отключение
        пользователем в настройках, суточный троттлинг. Любой False — выход
        без сетевых вызовов и без интерфейсных изменений.
        """
        if not updater.check_enabled():
            return
        if not updater.auto_check_enabled(self.app.settings):
            return
        if not updater.auto_check_due(self.app.settings):
            return
        self._start(auto=True)

    def check_now(self) -> None:
        """Кнопка «Проверить обновления»: ручная проверка с guard."""
        if self.busy:
            return
        self._start(auto=False)

    def _start(self, auto: bool) -> None:
        self.busy = True
        self._set_busy_ui(True)
        # Очередь результата + опрос из главного потока: рабочий поток
        # не трогает Qt (контракт раздела Д плана).
        result_queue: queue.Queue = queue.Queue()
        self.result_queue = result_queue
        factory = self._thread_factory or _default_thread_factory()
        factory(
            target=self._worker,
            args=(app_common.APP_VERSION, auto, result_queue),
            daemon=True,
        ).start()
        self.app.add_poll_timer(CHECK_TIMER, RESULT_INTERVAL_MS, self.result_tick)

    def _worker(self, current_version: str, auto: bool,
                result_queue: "queue.Queue") -> None:
        """Рабочий поток: только HTTP и очередь — никаких вызовов Qt."""
        try:
            result = updater.check_for_update(current=current_version)
        except Exception as exc:  # сетевой слой не должен ронять очередь
            result = {"status": "error", "message": str(exc), "release": None}
        result_queue.put((auto, result))

    # ------------------------------------------------------------------ #
    #  Опрос результата
    # ------------------------------------------------------------------ #
    def result_tick(self) -> None:
        """Опрашивает очередь результата из главного потока Qt."""
        result_queue = self.result_queue
        if result_queue is None:
            self.app.stop_poll_timer(CHECK_TIMER)
            return
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            return  # QTimer периодический: перезапускать не нужно
        except RuntimeError:
            return  # окно закрывается — цикл остановится вместе с ним
        self.app.stop_poll_timer(CHECK_TIMER)
        self.result_queue = None
        self._finish(item[0], item[1])

    # ------------------------------------------------------------------ #
    #  Результат: интерфейс
    # ------------------------------------------------------------------ #
    def _set_busy_ui(self, busy: bool) -> None:
        """Блокирует кнопку проверки на время запроса (если она создана)."""
        button = getattr(self.app, "_update_check_button", None)
        if button is None:
            return
        try:
            button.setEnabled(not busy)
            button.setText("Проверяем…" if busy else "Проверить обновления")
        except RuntimeError:
            pass  # виджет уже уничтожен — проверка всё равно завершится

    def _finish(self, auto: bool, result: dict) -> None:
        """Разбирает результат проверки: диалог или тихая фиксация.

        Время проверки отмечается и при авто-, и при ручной проверке —
        суточный троттлинг считает от любой удачной попытки.
        """
        self.busy = False
        self._set_busy_ui(False)
        settings = self.app.settings
        updater.mark_checked(settings)
        self.app.save_setting(updater.SETTINGS_KEY_TS,
                              settings[updater.SETTINGS_KEY_TS])
        self._refresh_status_text()

        status = result.get("status")
        release = result.get("release") or {}
        if status == "update":
            self._show_update_dialog(release)
            return
        if auto:
            return  # тихо: нет обновления или сеть недоступна — не тревожим
        if status == "latest":
            self.app.alert(
                self.app.window, "Обновлений нет",
                "Вы используете последнюю версию EW.DocHub.", "success",
            )
            return
        self.app.alert(
            self.app.window, "Не удалось проверить обновления",
            str(result.get("message") or "Неизвестная ошибка проверки.")
            + "\n\nСтраницу релизов можно открыть вручную:\n" + updater.PAGE_URL,
            "error",
        )

    def _refresh_status_text(self) -> None:
        """Обновляет строку статуса на странице «Настройки», если она есть."""
        label = getattr(self.app, "_update_status_label", None)
        if label is None:
            return
        try:
            label.setText(self.status_text())
        except RuntimeError:
            pass

    # ------------------------------------------------------------------ #
    #  Диалог обновления: «Обновить сейчас» / страница / позже
    # ------------------------------------------------------------------ #
    def _flavor_asset(self, release: dict) -> dict:
        """Ассет СВОЕГО флейвора (И.2); чужой флейвор не скачивается."""
        return updater.installer_asset_for_flavor(release, self.flavor)

    def _show_update_dialog(self, release: dict) -> None:
        newest = str(release.get("version_str") or release.get("tag") or "")
        has_installer = bool(self._flavor_asset(release).get("url"))
        if has_installer:
            message = (
                f"Установлена версия: {app_common.APP_VERSION}\n"
                f"Новая версия: {newest}\n\n"
                "«Обновить сейчас» скачает установщик, проверит его целостность "
                "и запустит обновление — приложение закроется и запустится "
                "уже новой версией. Реестр и настройки сохранятся."
            )
            options = [
                ("later", "Позже", False),
                ("page", "Открыть страницу", False),
                ("install", "Обновить сейчас", True),
            ]
        else:
            # Установщика СВОЕГО флейвора в релизе нет — обновиться
            # программой нельзя, остаётся страница релиза. Чужой ассет
            # (другой сборки) не предлагается и не скачивается (И.2).
            message = (
                f"Установлена версия: {app_common.APP_VERSION}\n"
                f"Новая версия: {newest}\n\n"
                "Установщик ещё не выложен в релиз — открыть страницу релиза, "
                "чтобы скачать его вручную?"
            )
            options = [
                ("later", "Позже", False),
                ("page", "Открыть страницу", True),
            ]
        choice = self.app.choice(
            self.app.window, "Доступна новая версия EW.DocHub", message, options,
        )
        if choice == "install":
            self.start_download(release)
        elif choice == "page":
            self.open_release_page(release)

    def open_release_page(self, release: dict | None = None) -> None:
        """Открывает страницу релиза в браузере (не блокирует интерфейс)."""
        url = str((release or {}).get("url") or updater.PAGE_URL)
        try:
            webbrowser.open(url)
        except Exception:
            self.app.alert(
                self.app.window, "Не удалось открыть браузер",
                "Откройте страницу релиза вручную:\n" + url, "info",
            )

    # ------------------------------------------------------------------ #
    #  Скачивание и запуск установки
    # ------------------------------------------------------------------ #
    def start_download(self, release: dict) -> None:
        """Открывает окно прогресса и запускает рабочий поток скачивания.

        Скачивается установщик СВОЕГО флейвора, затем SHA256SUMS.txt и
        сверяется SHA-256; всё это в одном рабочем потоке, результат —
        через очередь ``download_queue`` и опрос ``download_tick``.
        """
        if self.download_queue is not None:
            return  # скачивание уже идёт
        asset = self._flavor_asset(release)
        url = str(asset.get("url") or "")
        if not url:
            self.open_release_page(release)
            return
        self.cancel_event = threading.Event()
        self.download_queue = queue.Queue()
        self._build_progress_window(release)
        factory = self._thread_factory or _default_thread_factory()
        factory(
            target=self._download_worker,
            args=(release, url, str(asset.get("name") or ""),
                  self.cancel_event, self.download_queue),
            daemon=True,
        ).start()
        self.app.add_poll_timer(
            DOWNLOAD_TIMER, DOWNLOAD_INTERVAL_MS, self.download_tick)

    def _build_progress_window(self, release: dict) -> None:
        """Окно скачивания: заголовок, статус, прогресс, отмена."""
        self.download_window = UpdateProgressDialogQt(
            self.app.window, release, self.cancel_download)
        self.download_window.show()

    def _download_worker(self, release: dict, url: str, filename: str,
                         cancel_event: "threading.Event",
                         result_queue: "queue.Queue") -> None:
        """Рабочий поток: скачивание установщика + проверка SHA-256.

        Никаких вызовов Qt; ошибки сети/проверки складываются в очередь
        как ("error", …), успех — как ("done", путь). Отмена —
        ("cancelled", …).
        """
        import os
        import tempfile

        name = filename or "EW.DocHub-setup.exe"
        dest_dir = os.path.join(tempfile.gettempdir(), "EW.DocHub-update")
        dest_path = os.path.join(dest_dir, name)
        try:
            def progress(done: int, total: int) -> None:
                result_queue.put(("progress", done, total))

            updater.download_file(
                url, dest_path, progress_cb=progress, cancel_event=cancel_event,
            )
            result_queue.put(("phase", "Проверка целостности файла…"))
            sums_url = str(release.get("sums_url") or "")
            expected = ""
            if sums_url:
                expected = updater.match_sums_line(
                    updater.fetch_text(sums_url), name)
            if not expected:
                result_queue.put(("error",
                    "Не удалось получить контрольную сумму релиза "
                    "(SHA256SUMS.txt) — проверка целостности невозможна."))
                return
            if not updater.verify_sha256(dest_path, expected):
                try:
                    os.remove(dest_path)
                except OSError:
                    pass
                result_queue.put(("error",
                    "Скачанный файл не прошёл проверку целостности "
                    "(SHA-256 не совпал). Файл удалён — повторите обновление "
                    "позже или скачайте установщик со страницы релиза."))
                return
            result_queue.put(("done", dest_path))
        except updater.DownloadCancelled:
            result_queue.put(("cancelled", None))
        except updater.DownloadNotAllowed as exc:
            # SEC-1 (D11): чужой хост / превышение лимита — честный отказ
            # пользователю, а не «ошибка сети».
            result_queue.put(("error", str(exc)))
        except Exception as exc:  # сеть/диск: пользовательская ошибка
            result_queue.put(("error", updater.human_error(exc)))

    def download_tick(self) -> None:
        """Опрашивает очередь скачивания из главного потока Qt.

        Регрессия 2.18.0 (Tk: зависание «0,0 МБ из 19,5 МБ»): за один тик
        очередь осушается пачкой; QTimer периодический и в перезапуске не
        нуждается — событие progress не может «убить» цикл опроса, как
        это было с after()-цепочкой. Остановка — только на терминальном
        done/cancelled/error.
        """
        result_queue = self.download_queue
        if result_queue is None:
            self.app.stop_poll_timer(DOWNLOAD_TIMER)
            return
        while True:
            try:
                item = result_queue.get_nowait()
            except queue.Empty:
                break  # событий больше нет — ждём следующий тик
            except RuntimeError:
                return  # окно закрывается — цикл остановится вместе с ним
            kind = item[0]
            if kind == "progress":
                if self.download_window is not None:
                    self.download_window.set_progress(item[1], item[2])
                continue  # не терминальное — цикл продолжается
            if kind == "phase":
                if self.download_window is not None:
                    self.download_window.set_status(item[1])
                continue
            # Терминальные события: done / cancelled / error.
            self.download_queue = None
            self.app.stop_poll_timer(DOWNLOAD_TIMER)
            self._close_progress_window()
            if kind == "done":
                self._confirm_install(item[1])
            elif kind == "cancelled":
                pass  # окно закрыто — ничего не делаем
            else:
                self.app.alert(
                    self.app.window, "Обновление не удалось",
                    str(item[1] or "Неизвестная ошибка скачивания.")
                    + "\n\nСтраницу релизов можно открыть вручную:\n"
                    + updater.PAGE_URL,
                    "error",
                )
            return

    def _close_progress_window(self) -> None:
        """Закрывает окно прогресса (если оно живо)."""
        window = self.download_window
        self.download_window = None
        if window is None:
            return
        try:
            window.close()
        except RuntimeError:
            pass  # окно уже уничтожено

    def cancel_download(self) -> None:
        """«Отмена» / закрытие окна: останавливает рабочий поток."""
        if self.cancel_event is not None:
            self.cancel_event.set()
        self.download_queue = None
        self.app.stop_poll_timer(DOWNLOAD_TIMER)

    def _confirm_install(self, path: str) -> None:
        """Финальное подтверждение: закрыть программу и запустить установщик."""
        proceed = self.app.confirm(
            self.app.window, "Обновление скачано",
            "Установщик проверен и готов к запуску.\n\n"
            "EW.DocHub закроется, установится новая версия и запустится "
            "автоматически. Реестр, настройки и ключ DaData сохранятся.\n\n"
            "Запустить установку сейчас?",
            ok_text="Обновить",
            destructive=True,  # U7: программа закроется — случайный Enter недопустим
        )
        if proceed:
            self.run_installer(path)

    def run_installer(self, path: str) -> None:
        """Запускает setup.exe в тихом режиме и закрывает приложение.

        Inno Setup с /VERYSILENT ставит обновление в ту же папку без
        вопросов (профиль в %APPDATA% не трогается), /NORESTART отменяет
        перезагрузку. Программа закрывается сразу после запуска — так
        установщик не встречает занятые файлы.
        """
        self._snapshot_registry_before_update()
        try:
            updater.run_installer_command(path)
        except Exception as exc:
            self.app.alert(
                self.app.window, "Не удалось запустить установщик",
                f"{exc}\n\nУстановщик сохранён здесь:\n{path}",
                "error",
            )
            return
        # Небольшая задержка: установщик стартует, затем приложение
        # закрывается штатно (очистки, сохранение настроек). Tk-паритет:
        # root.after(500, app.close) — подтверждение «грязных» страниц
        # сработает только при реально несохранённых данных (по контракту
        # STORE-1A их нет — все изменения пишутся сразу).
        self.app.call_later(500, self.app.window.close)

    def _snapshot_registry_before_update(self) -> None:
        """Снимок реестра перед установкой новой версии (BKUP-1, D-карточка).

        Хук перед run_installer_command: обновление — необратимая на уровне
        данных операция, снимок сохраняет состояние базы с меткой текущей
        версии. Снимок ≠ откат версии приложения: восстановление базы —
        ручная операция. Отказ снимка обновление не блокирует — факт виден
        в журнале.
        """
        try:
            import storage

            storage.create_update_snapshot(
                self.app.store, app_common.APP_VERSION)
        except Exception as exc:  # noqa: BLE001 — обновление не блокируем
            applog.get_logger("updater").warning(
                "Снимок реестра перед обновлением не создан: %s", type(exc).__name__,
            )

    # ------------------------------------------------------------------ #
    #  Текст статуса для страницы «Настройки»
    # ------------------------------------------------------------------ #
    def status_text(self) -> str:
        """Строка «Установлена версия X.Y.Z-флейвор» — без сетевых вызовов.

        Суффикс флейвора — И.2 (M7): сборка объявляет себя в интерфейсе;
        сравнение версий updater идёт только по числовой части APP_VERSION.
        """
        return f"Установлена версия {app_common.APP_VERSION}-{self.flavor}"
