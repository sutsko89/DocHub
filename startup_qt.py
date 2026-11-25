"""Стартовые дисковые задачи в рабочем потоке на Qt (этап M6) — START-1.

Qt-перенос ``startup_controller.StartupFlow`` (Tk): сверка «реестр ↔ папка
документов», очистки (.docx.tmp, просроченная корзина) и сбор сроков доп.
соглашений после запуска выполняются в рабочем потоке, результаты
применяются ТОЛЬКО в главном потоке Qt (опрос очереди QTimer'ом — контракт
раздела Д плана: воркер не трогает UI, не QThread-сигналы).

Контракт владения состоянием (этап 5 roadmap) переносится дословно:

* владелец — главный поток; работник делает только read-only работу:
  скан файлов, SQL-чтение с собственным соединением (LAZY-1:
  bulk_source_files), сбор сроков по снимку списка и лёгкой карте сроков;
* мутации реестра (очистка корзины) и любые вызовы UI применяет главный
  поток; устаревший результат отбрасывается вместе с очередью
  завершившегося запуска (флаг остановки проверяется между шагами);
* TRASH-WARN (этап 2): корзина не удаляет записи молча — перед очисткой
  сводка «будет удалено N записей» в журнале, после — итог в статусе и
  немодальном уведомлении; при пустой сводке очистка не запускается;
* сверка (DEF-3): список файлов вне реестра попадает в баннер страницы
  «Реестр» с кнопкой «Разобрать» — восстановить в реестр или перенести
  в служебную папку _trash;
* сроки доп. соглашений (2.18.0): раз в сутки диалог «Открыть реестр» /
  «Позже»; отметка «показано» пишется только когда диалог показывается;
* каждый шаг воркера и каждое главнопоточное применение обёрнуты в
  ``applog.timed`` — perf-каналы ``startup.*`` совпадают с Tk-сборкой.

Автокопия BKUP-1 остаётся в прежней точке запуска (main_qt, отдельный
одноразовый таймер, как root.after(400) в Tk); сетевая проверка
обновлений — M7.
"""

from __future__ import annotations

import os
import queue
import time
from collections.abc import Callable

import agreements
import applog
import reconcile
import threading

TEMP_SUFFIX = ".docx.tmp"

# Суточный гейт уведомления о сроках (ключ settings.json — Tk-паритет).
EXPIRY_NOTICE_SETTINGS_KEY = "last_expiry_notice"
EXPIRY_NOTICE_INTERVAL = 24 * 60 * 60


def _default_thread_factory() -> type:
    """Фабрика потока разрешается в момент вызова (подмена в тестах)."""
    return threading.Thread


class StartupFlowQt:
    """Владелец стартовых дисковых задач (START-1).

    Состояние потока (``running``/``_queue``/``_stopped``) живёт здесь.
    Вызовы, которые подменяет smoke/тест (``app.cleanup_expired_trash``,
    ``app.apply_expiry_notice``, ``registry_flow.scan_consistency``),
    идут через ``self.app.*`` — подмена instance-атрибута работает.
    """

    def __init__(self, app, thread_factory: Callable | None = None):
        self.app = app
        self.running = False
        self._stopped = False
        self._queue: queue.Queue | None = None
        self._thread_factory = thread_factory

    # ------------------------------------------------------------------ #
    #  Запуск и остановка
    # ------------------------------------------------------------------ #
    def run(self, include_notice: bool = True) -> bool:
        """Запускает стартовые задачи в фоне. False — уже выполняются.

        ``include_notice=False`` — суточный гейт уведомления о сроках не
        прошёл (гейт читается в главном потоке до запуска, запись отметки
        делает главный поток после показа диалога).
        """
        if self.running:
            return False
        self.running = True
        self._stopped = False
        result_queue: queue.Queue = queue.Queue()
        self._queue = result_queue
        factory = self._thread_factory or _default_thread_factory()
        factory(target=self.worker, args=(result_queue, include_notice),
                daemon=True).start()
        self.app.add_poll_timer("startup_tick", 80, self.tick)
        return True

    def stop(self) -> None:
        """Флаг остановки: работник проверяет его между шагами и файлами."""
        self._stopped = True

    # ------------------------------------------------------------------ #
    #  Работник (только чтение: файлы, SQL с собственным соединением)
    # ------------------------------------------------------------------ #
    def worker(self, result_queue: queue.Queue, include_notice: bool = True) -> None:
        """Выполняет задачи по порядку; результаты — в очередь запуска."""
        app = self.app
        store = app.store
        try:
            with applog.timed("startup.temp", 500, "startup"):
                temp_removed = self._cleanup_temp_documents()
            self._emit(result_queue, "temp", temp_removed)
            if self._stopped:
                return
            with applog.timed("startup.trash_count", 500, "startup"):
                try:
                    planned = store.count_expired_trash()
                except Exception as exc:  # noqa: BLE001 — задача старта не роняет запуск
                    applog.get_logger("storage").warning(
                        "Подсчёт корзины перед автоочисткой не удался: %s",
                        type(exc).__name__,
                    )
                    planned = 0
            self._emit(result_queue, "trash_plan", planned)
            if self._stopped:
                return
            with applog.timed("startup.orphans", 500, "startup"):
                orphans = self._find_orphans()
            self._emit(result_queue, "orphans", orphans)
            if self._stopped or not include_notice:
                if not self._stopped:
                    self._emit(result_queue, "notice", None)
                return
            # Сбор сроков: снимок списка (не итерация живого хранилища),
            # сроки — из лёгкой карты хранилища (LAZY-1); describe читает
            # имя контрагента из карточки (только чтение).
            with applog.timed("startup.notice", 500, "startup"):
                try:
                    text = agreements.notice_text(
                        store, docs=list(store.active_documents()),
                    )
                except Exception:  # noqa: BLE001
                    applog.get_logger("agreements").exception(
                        "Сбор сроков доп. соглашений не удался",
                    )
                    text = None
            self._emit(result_queue, "notice", text)
        finally:
            self._emit(result_queue, "done", None)

    def _documents_dir_noninteractive(self) -> str:
        """Папка документов без диалогов (фоновая задача, Tk-паритет).

        Выбранная папка при недоступности не подменяется молча: raising
        ветка — та же, что у импорта; сбой гасится на уровне задач.
        """
        custom = str((self.app.settings.get("documents_dir") or "")).strip()
        if custom:
            os.makedirs(custom, exist_ok=True)
            return custom
        import shell_open

        base = os.path.join(shell_open.documents_known_folder(), "EW.DocHub")
        os.makedirs(base, exist_ok=True)
        return base

    def _cleanup_temp_documents(self) -> int:
        """Удаляет *.docx.tmp после аварийного завершения записи (чистый диск)."""
        try:
            base = self._documents_dir_noninteractive()
            directories = [base]
            with os.scandir(base) as entries:
                directories.extend(
                    entry.path for entry in entries if entry.is_dir()
                )
        except OSError:
            return 0
        removed = 0
        for directory in directories:
            if self._stopped:
                break
            try:
                with os.scandir(directory) as entries:
                    for entry in entries:
                        if self._stopped:
                            break
                        if entry.is_file() and entry.name.endswith(TEMP_SUFFIX):
                            try:
                                os.remove(entry.path)
                                removed += 1
                            except OSError:
                                pass
            except OSError:
                continue
        return removed

    def _find_orphans(self) -> list:
        """Сверка папки с реестром: bulk-карта source_file (LAZY-1)."""
        app = self.app
        try:
            base = self._documents_dir_noninteractive()
            if not base:
                return []
            return reconcile.find_orphans(
                base, app.store, source_files=app.store.bulk_source_files(),
            )
        except Exception as exc:  # noqa: BLE001 — задача старта не роняет запуск
            applog.get_logger("storage").warning(
                "Сверка реестра с папкой не удалась: %s", type(exc).__name__,
            )
            return []

    # ------------------------------------------------------------------ #
    #  Применение результатов — только в главном потоке
    # ------------------------------------------------------------------ #
    def _emit(self, result_queue: queue.Queue, kind: str, payload) -> None:
        result_queue.put((kind, payload))

    def tick(self) -> None:
        """Опрос очереди QTimer'ом: дренаж + перепланирование до 'done'."""
        finished = self._drain()
        if not finished:
            return
        self.running = False
        self.app.stop_poll_timer("startup_tick")

    def _drain(self) -> bool:
        """Применяет накопленные результаты; True — получен 'done'."""
        assert self._queue is not None
        while True:
            try:
                kind, payload = self._queue.get_nowait()
            except queue.Empty:
                return False
            if kind == "done":
                return True
            self._apply(kind, payload)

    def _apply(self, kind: str, payload) -> None:
        if kind == "temp":
            if payload:
                self.app.log.info("Удалено незавершённых .docx.tmp: %d", payload)
        elif kind == "trash_plan":
            if payload:
                # Мутация реестра — только главный поток (TRASH-WARN: сводка
                # в журнал, итог в статус, refresh реестра). E-4: длительность
                # — запись startup.apply_trash (главный поток).
                with applog.timed("startup.apply_trash", 500, "startup"):
                    self.app.cleanup_expired_trash(planned=payload)
        elif kind == "orphans":
            # Применение результата сверки + refresh при открытой странице.
            with applog.timed("startup.apply_scan", 500, "startup"):
                self.app.registry_flow.scan_consistency(orphans=payload)
        elif kind == "notice":
            with applog.timed("startup.apply_notice", 500, "startup"):
                self.app.apply_expiry_notice(payload)
