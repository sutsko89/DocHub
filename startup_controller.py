"""Стартовые дисковые задачи в рабочем потоке — START-1 (R-07 / P-04).

Сверка «реестр ↔ папка документов», очистки (.docx.tmp, просроченная
корзина) и сбор сроков доп. соглашений после запуска выполнялись
after-цепочкой на UI-потоке (600/900/1500/1900 мс) и забирали главный
поток дисковыми операциями в первые секунды (П-аудит P-04). Здесь они
выполняются в рабочем потоке по шаблону import_controller: guard
повторного входа, собственная очередь у каждого запуска (воркеру через
args), флаг остановки между шагами и файлами, результаты применяются
ТОЛЬКО в главном потоке (tick из Tk after).

Контракт владения состоянием (этап 5 roadmap): владелец — главный
поток; работник делает только read-only работу — скан файлов, SQL-чтение
с собственным соединением (LAZY-1: bulk_source_files), сбор сроков по
снимку списка и лёгкой карте сроков. Мутации реестра (очистка корзины)
и любые вызовы UI применяет главный поток; устаревший результат
отбрасывается вместе с очередью завершившегося запуска.

Автокопия BKUP-1 и сетевая проверка обновлений остаются в прежних
точках запуска: первая не входит в объём START-1, вторая уже
глушится env-флагом и ходит в сеть по собственному таймеру.
"""

from __future__ import annotations

import os
import queue
from collections.abc import Callable

import agreements
import applog
import reconcile
import threading


TEMP_SUFFIX = ".docx.tmp"


def _default_thread_factory() -> type:
    """Фабрика потока разрешается в момент вызова (подмена в тестах)."""
    return threading.Thread


class StartupFlow:
    """Владелец стартовых дисковых задач (START-1).

    Состояние потока (``running``/``_queue``/``_stopped``) живёт здесь;
    приложение публикует его через свойство с прежним именем. Вызовы,
    которые подменяет smoke-тест (``_cleanup_expired_trash``,
    ``_apply_expiry_notice``), идут через ``self.app.*`` — подмена
    instance-атрибута продолжает работать.
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
        factory(target=self.worker, args=(result_queue, include_notice), daemon=True).start()
        self.app.root.after(80, self.tick)
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
            self._emit(result_queue, "temp", self._cleanup_temp_documents())
            if self._stopped:
                return
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
            self._emit(result_queue, "orphans", self._find_orphans())
            if self._stopped or not include_notice:
                if not self._stopped:
                    self._emit(result_queue, "notice", None)
                return
            # Сбор сроков: снимок списка (не итерация живого хранилища),
            # сроки — из лёгкой карты хранилища (LAZY-1); describe читает
            # имя контрагента из карточки (только чтение).
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

    def _cleanup_temp_documents(self) -> int:
        """Удаляет *.docx.tmp после аварийного завершения записи (чистый диск)."""
        app = self.app
        try:
            base = app._settings_flow.documents_dir(interactive=False)
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
            base = app._settings_flow.documents_dir(interactive=False)
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
        """Опрос очереди из Tk after: дренаж + перепланирование до 'done'."""
        finished = self._drain()
        if not finished:
            self.app.root.after(80, self.tick)
        else:
            self.running = False

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
                self.app._log.info("Удалено незавершённых .docx.tmp: %d", payload)
        elif kind == "trash_plan":
            if payload:
                # Мутация реестра — только главный поток (TRASH-WARN: сводка
                # в журнал, итог в статус, refresh реестра).
                self.app._cleanup_expired_trash(planned=payload)
        elif kind == "orphans":
            # Применение результата сверки + refresh при открытой странице.
            self.app._registry_flow.scan_consistency(orphans=payload)
        elif kind == "notice":
            self.app._apply_expiry_notice(payload)
