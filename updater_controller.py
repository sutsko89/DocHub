"""Проверка обновлений и автообновление — восьмой контроллер ContractApp.

Владеет всей оркестрацией обновления приложения:

* тихая автопроверка при старте (раз в сутки, если включена пользователем
  и не заглушена средой) — ``maybe_auto_check``, вызывается стартовым
  таймером ``after(2200, …)`` ContractApp;
* ручная проверка кнопкой «Проверить обновления» на странице «Настройки»
  — ``check_now`` с guard повторных кликов и блокировкой кнопки;
* диалог результата (2.18.0): «Обновить сейчас» / «Открыть страницу» /
  «Позже»; при выборе «Обновить сейчас» — окно скачивания с прогрессом
  и отменой, контроль SHA-256 по SHA256SUMS.txt релиза, запуск
  установщика Inno Setup в тихом режиме и закрытие приложения;
* честное «версия последняя» / ошибка сети при ручной проверке;
  автопроверка молчит, если обновления нет или сеть недоступна.

Паттерн потоков — тот же, что у dadata_controller (аудит 2.0.1): HTTP в
рабочем потоке, результат через ``queue.Queue``, опрос из главного потока
Tk; рабочий поток не трогает Tk и не вызывает ``root.after()``.
Фактическая сетевая работа — в updater.py (без Tk), контроллер только
маршрутизирует результат в интерфейс.

Подменяемые маршруты (для smoke-теста и юнит-тестов):

* ``updater.check_for_update`` / ``updater.download_file`` /
  ``updater.fetch_text`` — сетевые функции подменяются целиком;
* ``threading.Thread`` — общий модуль, синхронная подмена из smoke-теста
  действует и на этот контроллер;
* ``webbrowser.open`` — открытие страницы релиза подменяется в тестах;
* ``run_installer_command`` — модульная функция запуска setup.exe;
* ``ui.alert`` / ``ui.confirm`` — атрибуты общего модуля ui_macos.

Зависимость от main (APP_VERSION, подмена main.threading в smoke) берётся
поздним импортом внутри методов — main импортирует этот контроллер.
"""

from __future__ import annotations

import applog
import os
import queue
import subprocess
import sys
import threading
import webbrowser
from typing import Any

import tkinter as tk

import ui_macos as ui
import updater


class UpdateController:
    """Жизненный цикл обновления: guard → проверка → диалог → скачивание → установка."""

    def __init__(self, app: Any) -> None:
        self.app = app
        self.busy = False              # guard повторных проверок
        self.result_queue: "queue.Queue | None" = None
        self.result_after = None       # id таймера опроса для after_cancel
        # Скачивание (2.18.0): своя очередь, окно прогресса и флаг отмены.
        self.download_queue: "queue.Queue | None" = None
        self.download_after = None
        self.cancel_event: "threading.Event | None" = None
        self.download_window = None
        self._download_progress = None
        self._download_status_var: "tk.StringVar | None" = None

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
    #  Диалог обновления: «Обновить сейчас» / страница / позже
    # ------------------------------------------------------------------ #
    def _show_update_dialog(self, release: dict) -> None:
        import main  # поздний импорт: APP_VERSION (см. докстринг класса)

        newest = str(release.get("version_str") or release.get("tag") or "")
        has_installer = bool(release.get("installer_url"))
        if has_installer:
            message = (
                f"Установлена версия: {main.APP_VERSION}\n"
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
            # Установщика в релизе нет — обновиться программой нельзя,
            # остаётся страница релиза (CI мог ещё не выложить ассет).
            message = (
                f"Установлена версия: {main.APP_VERSION}\n"
                f"Новая версия: {newest}\n\n"
                "Установщик ещё не выложен в релиз — открыть страницу релиза, "
                "чтобы скачать его вручную?"
            )
            options = [
                ("later", "Позже", False),
                ("page", "Открыть страницу", True),
            ]
        choice = self.app._choice_dialog(
            "Доступна новая версия EW.DocHub", message, options,
        )
        if choice == "install":
            self.start_download(release)
        elif choice == "page":
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
    #  Скачивание и запуск установки (2.18.0)
    # ------------------------------------------------------------------ #
    def start_download(self, release: dict) -> None:
        """Открывает окно прогресса и запускает рабочий поток скачивания.

        Скачивается setup.exe релиза, затем SHA256SUMS.txt и сверяется
        SHA-256; всё это в одном рабочем потоке, результат — через
        очередь ``download_queue`` и опрос ``download_tick`` из Tk.
        """
        if self.download_queue is not None:
            return  # скачивание уже идёт
        url = str(release.get("installer_url") or "")
        if not url:
            self.open_release_page(release)
            return
        self.cancel_event = threading.Event()
        self.download_queue = queue.Queue()
        self._build_progress_window(release)
        threading.Thread(
            target=self._download_worker,
            args=(release, url, self.cancel_event, self.download_queue),
            daemon=True,
        ).start()
        self.download_after = self.app.root.after(80, self.download_tick)

    def _build_progress_window(self, release: dict) -> None:
        """Модальное окно скачивания: заголовок, прогресс, отмена."""
        import customtkinter as ctk

        app = self.app
        newest = str(release.get("version_str") or release.get("tag") or "")
        win = ctk.CTkToplevel(app.root)
        win.title("Обновление EW.DocHub")
        win.configure(fg_color=ui.COLORS["window"])
        win.transient(app.root)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self.cancel_download)
        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True,
                  padx=int(24 * app.scale), pady=int(20 * app.scale))
        ctk.CTkLabel(
            body, text=f"Скачивание обновления {newest}",
            font=ui.font(15, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w",
        ).pack(anchor="w")
        self._download_status_var = tk.StringVar(
            value="Подключение к github.com…")
        ctk.CTkLabel(
            body, textvariable=self._download_status_var,
            font=ui.font(12), text_color=ui.COLORS["text_secondary"],
            anchor="w", justify="left", wraplength=int(380 * app.scale),
        ).pack(anchor="w", pady=(int(8 * app.scale), int(12 * app.scale)))
        self._download_progress = ctk.CTkProgressBar(
            body, width=int(380 * app.scale),
            progress_color=ui.COLORS["accent"],
        )
        self._download_progress.pack(fill="x")
        self._download_progress.set(0.0)
        ctk.CTkButton(
            body, text="Отмена", command=self.cancel_download,
            width=int(120 * app.scale), height=int(30 * app.scale),
            fg_color=ui.COLORS["card"], hover_color=ui.COLORS["button_secondary_hover"],
            border_width=1, border_color=ui.COLORS["card_border"],
            text_color=ui.COLORS["text"],
        ).pack(anchor="e", pady=(int(14 * app.scale), 0))
        self.download_window = win

    def _download_worker(self, release: dict, url: str,
                         cancel_event: "threading.Event",
                         result_queue: "queue.Queue") -> None:
        """Рабочий поток: скачивание установщика + проверка SHA-256.

        Никаких вызовов Tk; ошибки сети/проверки складываются в очередь
        как ("error", …), успех — как ("done", путь). Отмена —
        ("cancelled", …).
        """
        import tempfile

        filename = str(release.get("installer_name") or "EW.DocHub-setup.exe")
        dest_dir = os.path.join(tempfile.gettempdir(), "EW.DocHub-update")
        dest_path = os.path.join(dest_dir, filename)
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
                    updater.fetch_text(sums_url), filename)
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
        except Exception as exc:  # сеть/диск: пользовательская ошибка
            result_queue.put(("error", updater.human_error(exc)))

    def download_tick(self) -> None:
        """Опрашивает очередь скачивания из главного потока Tk.

        Регрессия 2.18.0 (зависание «0,0 МБ из 19,5 МБ»): после события
        ``progress`` цикл опроса не перевешивал ``after`` — опрос умирал
        на первом же пакете прогресса, окно застывало, а финальное
        ``done`` никто не забирал, хотя скачивание в рабочем потоке
        продолжалось. Теперь за один тик очередь осушается пачкой, а
        ``after`` перевешивается после КАЖДОГО не-терминального события;
        остановка цикла — только на терминальном ``done``/``cancelled``/
        ``error`` или при разрушении окна.
        """
        result_queue = self.download_queue
        if result_queue is None:
            return
        while True:
            try:
                item = result_queue.get_nowait()
            except queue.Empty:
                break  # событий больше нет — планируем следующий опрос
            except (RuntimeError, AttributeError):
                return  # окно закрывается — цикл прервётся вместе с ним
            kind = item[0]
            if kind == "progress":
                self._show_progress(item[1], item[2])
                continue  # не терминальное — цикл продолжается
            if kind == "phase":
                # Промежуточная фаза (проверка суммы): окно остаётся
                # открытым, меняется только строка статуса.
                self._set_download_status(item[1])
                continue
            # Терминальные события: done / cancelled / error.
            self.download_queue = None
            self._close_progress_window()
            if kind == "done":
                self._confirm_install(item[1])
            elif kind == "cancelled":
                pass  # окно закрыто — ничего не делаем
            else:
                ui.alert(
                    self.app.root, "Обновление не удалось",
                    str(item[1] or "Неизвестная ошибка скачивания.")
                    + "\n\nСтраницу релизов можно открыть вручную:\n"
                    + updater.PAGE_URL,
                    "error",
                )
            return
        try:
            self.download_after = self.app.root.after(80, self.download_tick)
        except Exception:
            pass  # окно закрывается — цикл прервётся вместе с ним

    def _set_download_status(self, text: str) -> None:
        """Обновляет строку статуса окна скачивания (если окно живо)."""
        var = self._download_status_var
        if var is None:
            return
        try:
            var.set(text)
        except Exception:
            pass  # окно уже разрушено

    def _show_progress(self, done: int, total: int) -> None:
        """Обновляет полосу и статус скачивания."""
        if self._download_progress is None:
            return
        try:
            if total > 0:
                self._download_progress.set(min(1.0, done / total))
                mb_done = done / (1024 * 1024)
                mb_total = total / (1024 * 1024)
                if mb_total >= 1:
                    text = (f"{mb_done:.1f} МБ из {mb_total:.1f} МБ")
                else:
                    text = f"{done / 1024:.0f} КБ из {total / 1024:.0f} КБ"
            else:
                self._download_progress.set(0.0)
                text = f"Скачано: {done / (1024 * 1024):.1f} МБ"
            self._download_status_var.set(text)
        except Exception:
            pass  # окно уже закрыто

    def _close_progress_window(self) -> None:
        """Закрывает окно прогресса (если оно живо)."""
        window = self.download_window
        self.download_window = None
        if window is None:
            return
        try:
            if window.winfo_exists():
                window.destroy()
        except Exception:
            pass

    def cancel_download(self) -> None:
        """Кнопка «Отмена»: останавливает рабочий поток и закрывает окно."""
        if self.cancel_event is not None:
            self.cancel_event.set()
        self.download_queue = None
        self._close_progress_window()

    def _confirm_install(self, path: str) -> None:
        """Финальное подтверждение: закрыть программу и запустить установщик."""
        proceed = ui.confirm(
            self.app.root, "Обновление скачано",
            "Установщик проверен и готов к запуску.\n\n"
            "EW.DocHub закроется, установится новая версия и запустится "
            "автоматически. Реестр, настройки и ключ DaData сохранятся.\n\n"
            "Запустить установку сейчас?",
            ok_text="Обновить",
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
            run_installer_command(path)
        except Exception as exc:
            ui.alert(
                self.app.root, "Не удалось запустить установщик",
                f"{exc}\n\nУстановщик сохранён здесь:\n{path}",
                "error",
            )
            return
        # Небольшая задержка: установщик стартует, затем приложение
        # закрывается штатно (очистки, сохранение настроек).
        self.app.root.after(500, self.app.close)

    def _snapshot_registry_before_update(self) -> None:
        """Снимок реестра перед установкой новой версии (BKUP-1, D-карточка).

        Хук перед run_installer_command: обновление — необратимая на уровне
        данных операция, снимок сохраняет состояние базы с меткой текущей
        версии. Снимок ≠ откат версии приложения: восстановление базы —
        ручная операция. Отказ снимка обновление не блокирует — факт виден
        в журнале.
        """
        try:
            import main  # поздний импорт: APP_VERSION без цикла
            import storage

            storage.create_update_snapshot(self.app.store, main.APP_VERSION)
        except Exception as exc:  # noqa: BLE001 — обновление не блокируем
            applog.get_logger("updater").warning(
                "Снимок реестра перед обновлением не создан: %s", type(exc).__name__,
            )

    # ------------------------------------------------------------------ #
    #  Текст статуса для страницы «Настройки»
    # ------------------------------------------------------------------ #
    def status_text(self) -> str:
        """Строка «Установлена версия X.Y.Z» — без сетевых вызовов."""
        import main

        return f"Установлена версия {main.APP_VERSION}"


def run_installer_command(path: str) -> None:
    """Стартует установщик Inno Setup в тихом режиме (модульная для тестов).

    /VERYSILENT — без вопросов (обновление в ту же папку, профиль в
    %APPDATA% не трогается); /NORESTART — без перезагрузки;
    /SUPPRESSMSGBOXES и /FORCECLOSEAPPLICATIONS — тихо закрывают
    запущенную программу, если она не успела выйти (все данные
    сохраняются сразу при изменениях, риск минимальный); после установки
    приложение перезапускается само ([Run] skipifnotsilent). На других ОС
    функция не запускается (обновление — Windows-сценарий).
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Установщик не найден: {path}")
    if not sys.platform.startswith("win"):
        raise OSError("Автообновление доступно только в Windows.")
    # creationflags: DETACHED_PROCESS — установщик живёт после закрытия
    # приложения; без консольного окна.
    detached = getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(
        [path, "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES",
         "/FORCECLOSEAPPLICATIONS"],
        close_fds=True,
        creationflags=detached,
    )
