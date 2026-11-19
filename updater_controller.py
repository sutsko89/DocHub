"""Проверка обновлений — восьмой контроллер ContractApp (2.15.0).

Владеет всей оркестрацией проверки обновлений:

* тихая автопроверка при старте (раз в сутки, если включена пользователем
  и не заглушена средой) — ``maybe_auto_check``, вызывается стартовым
  таймером ``after(2200, …)`` ContractApp;
* ручная проверка кнопкой «Проверить обновления» на странице «Настройки»
  — ``check_now`` с guard повторных кликов и блокировкой кнопки;
* диалог результата: «Доступна новая версия — открыть страницу релиза?»,
  открытие страницы в браузере, честное «версия последняя» / ошибка сети
  при ручной проверке; автопроверка молчит, если обновления нет или сеть
  недоступна (пользователя не нужно тревожить неудавшейся фоновой задачей).

Паттерн потоков — тот же, что у dadata_controller (аудит 2.0.1): HTTP в
рабочем потоке, результат через ``queue.Queue``, опрос из главного потока
Tk; рабочий поток не трогает Tk и не вызывает ``root.after()``.
Фактическая сетевая работа — в updater.py (без Tk), контроллер только
маршрутизирует результат в интерфейс.

Подменяемые маршруты (для smoke-теста и юнит-тестов):

* ``updater.check_for_update`` — сетевая функция подменяется целиком;
* ``threading.Thread`` — общий модуль, синхронная подмена из smoke-теста
  действует и на этот контроллер;
* ``webbrowser.open`` — открытие страницы релиза подменяется в тестах;
* ``ui.alert`` / ``ui.confirm`` — атрибуты общего модуля ui_macos.

Зависимость от main (APP_VERSION, подмена main.threading в smoke) берётся
поздним импортом внутри методов — main импортирует этот контроллер.
"""

from __future__ import annotations

import queue
import threading
import webbrowser
from typing import Any

import ui_macos as ui
import updater


class UpdateController:
    """Жизненный цикл проверки обновлений: guard → HTTP → диалог."""

    def __init__(self, app: Any) -> None:
        self.app = app
        self.busy = False              # guard повторных проверок
        self.result_queue: "queue.Queue | None" = None
        self.result_after = None       # id таймера опроса для after_cancel

    # ------------------------------------------------------------------ #
    #  Запуск проверки
    # ------------------------------------------------------------------ #
    def maybe_auto_check(self) -> None:
        """Стартовый таймер ContractApp: тихая проверка, если положено.

        Три независимых рубильника: env-глушилка (тесты/офлайн), отключение
        пользователем в настройках, суточный троттлинг. Любой False — выход
        без сетевых вызовов и без интерфейсных изменений.
        """
        import main  # поздний импорт: цикл main → контроллер

        if not updater.check_enabled():
            return
        if not updater.auto_check_enabled(self.app._settings):
            return
        if not updater.auto_check_due(self.app._settings):
            return
        self._start(auto=True)

    def check_now(self) -> None:
        """Кнопка «Проверить обновления»: ручная проверка с guard."""
        if self.busy:
            return
        self._start(auto=False)

    def _start(self, auto: bool) -> None:
        import main  # поздний импорт: APP_VERSION читается на момент проверки

        self.busy = True
        self._set_busy_ui(True)
        # Очередь результата + опрос из главного потока (паттерн
        # dadata_controller): рабочий поток не трогает Tk.
        result_queue: queue.Queue = queue.Queue()
        self.result_queue = result_queue
        threading.Thread(
            target=self._worker,
            args=(main.APP_VERSION, auto, result_queue),
            daemon=True,
        ).start()
        self.result_after = self.app.root.after(80, self.result_tick)

    def _worker(self, current_version: str, auto: bool, result_queue: "queue.Queue") -> None:
        """Рабочий поток: только HTTP и очередь — никаких вызовов Tk."""
        try:
            result = updater.check_for_update(current=current_version)
        except Exception as exc:  # сетевой слой не должен ронять очередь
            result = {"status": "error", "message": str(exc), "release": None}
        result_queue.put((auto, result))

    # ------------------------------------------------------------------ #
    #  Опрос результата
    # ------------------------------------------------------------------ #
    def result_tick(self) -> None:
        """Опрашивает очередь результата из главного потока Tk."""
        result_queue = self.result_queue
        if result_queue is None:
            return  # проверка завершена или окно закрывается
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            try:
                self.result_after = self.app.root.after(80, self.result_tick)
            except Exception:
                pass  # окно закрывается — цикл прервётся вместе с ним
            return
        except (RuntimeError, AttributeError):
            return
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
            if busy:
                button.configure(text="Проверяем…", state="disabled")
            else:
                button.configure(text="Проверить обновления", state="normal")
        except Exception:
            pass  # виджет уже разрушен — проверка всё равно завершится

    def _finish(self, auto: bool, result: dict) -> None:
        """Разбирает результат проверки: диалог или тихая фиксация.

        Время проверки отмечается и при авто-, и при ручной проверке —
        суточный троттлинг считает от любой удачной попытки.
        """
        self.busy = False
        self._set_busy_ui(False)
        updater.mark_checked(self.app._settings)
        self.app._persist_settings()
        self._refresh_status_text()

        status = result.get("status")
        release = result.get("release") or {}
        if status == "update":
            self._show_update_dialog(release)
            return
        if auto:
            return  # тихо: нет обновления или сеть недоступна — не тревожим
        if status == "latest":
            ui.alert(
                self.app.root, "Обновлений нет",
                "Вы используете последнюю версию EW.DocHub.", "success",
            )
            return
        ui.alert(
            self.app.root, "Не удалось проверить обновления",
            str(result.get("message") or "Неизвестная ошибка проверки.")
            + "\n\nСтраницу релизов можно открыть вручную:\n" + updater.PAGE_URL,
            "error",
        )

    def _refresh_status_text(self) -> None:
        """Обновляет строку статуса на странице «Настройки», если она есть."""
        status_var = getattr(self.app, "_update_status_var", None)
        if status_var is None:
            return
        try:
            status_var.set(self.status_text())
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    #  Диалог обновления и открытие страницы релиза
    # ------------------------------------------------------------------ #
    def _show_update_dialog(self, release: dict) -> None:
        import main  # поздний импорт: APP_VERSION (см. докстринг класса)

        newest = str(release.get("version_str") or release.get("tag") or "")
        open_now = ui.confirm(
            self.app.root, "Доступна новая версия EW.DocHub",
            f"Установлена версия: {main.APP_VERSION}\n"
            f"Новая версия: {newest}\n\n"
            "Открыть страницу релиза, чтобы скачать новый установщик?",
            ok_text="Открыть страницу",
        )
        if open_now:
            self.open_release_page(release)

    def open_release_page(self, release: dict | None = None) -> None:
        """Открывает страницу релиза в браузере (фоново — не блокирует Tk)."""
        url = str((release or {}).get("url") or updater.PAGE_URL)
        try:
            webbrowser.open(url)
        except Exception:
            ui.alert(
                self.app.root, "Не удалось открыть браузер",
                "Откройте страницу релиза вручную:\n" + url, "info",
            )

    # ------------------------------------------------------------------ #
    #  Текст статуса для страницы «Настройки»
    # ------------------------------------------------------------------ #
    def status_text(self) -> str:
        """Строка «Установлена версия X.Y.Z» — без сетевых вызовов."""
        import main

        return f"Установлена версия {main.APP_VERSION}"
