"""Пакетный импорт документов на Qt (этап M6) — перенос import_controller.

Qt-перенос ``import_controller.ImportController`` (Tk): guard повторного
входа, проверки пакета до запуска (лимиты п. 3.4), фоновый воркер,
опрос очереди из главного потока (QTimer — контракт раздела Д плана:
воркер не трогает UI), пакетная сводка решений U9, регистрация
результатов одним ``save()`` (п. 5.2) и отчёт с «Открыть реестр» (U10).

Контракты, переносимые дословно (Tk-паритет):

* guard стоит в начале запуска — повторный запуск во время работающего
  импорта не сбрасывает очередь/результаты первого (аудит 2.0.1, P1);
* очередь у каждого запуска своя и передаётся воркеру через args;
* импорт принимает только .docx/.xlsx (allowlist): .exe/.bat/.lnk и
  прочее не анализируется и не копируется;
* карточка создаётся автоматически только при найденном ИНН; без ИНН —
  решение пользователя (привязать/создать/пропустить); U9: для серии
  файлов без ИНН и серии конфликтов названий — ОДИН диалог сводки на
  серию, а не диалог на каждый файл (критерий: импорт 20 файлов без
  ИНН — 1–2 диалога вместо 20+);
* документ, который уже есть в реестре (тип + номер + дата), повторно
  не добавляется;
* отклонённые файлы переносятся в папку unknown (п. 3.5);
* пакетная запись реестра одним ``save()`` в конце (п. 5.2);
* отмена: остаток пакета не анализируется и не регистрируется (п. 5.3);
* папка документов для импорта — non-interactive ветка: импорт идёт в
  фоне, где диалоги недопустимы; недоступная папка — ошибка в отчёте.

Отличия от Tk-сборки (docs/qt-migration-stage-notes.md): окно прогресса
и отчёт — QDialog; диалоги решения — app.confirm/app.choice
(qt_dialogs); выбор файлов — QFileDialog.getOpenFileNames; перетаскивание
файлов в окно — нативный Qt drag&drop (MainWindow), Win32-хуки Tk
(WM_DROPFILES) не переносятся.
"""

from __future__ import annotations

import os
import queue
import shutil
import threading
from collections.abc import Callable
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFileDialog, QHBoxLayout, QLabel, QProgressBar, QPushButton,
    QTextBrowser, QVBoxLayout,
)

import applog
import document_importer as importer
import services
import storage
from document_builder import F

# Импорт принимает только документы, которые программа умеет разбирать.
# Всё остальное (в т. ч. .exe/.bat/.lnk) не копируется и не открывается.
ALLOWED_EXTENSIONS = {".docx", ".xlsx"}


class ImportProgressDialog(QDialog):
    """Модальное окно прогресса: статус, полоса, «Отменить» (Tk-паритет)."""

    def __init__(self, parent, on_cancel: Callable[[], None]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Загрузка документов")
        self.setModal(True)
        self.setFixedSize(430, 240)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 22, 28, 18)
        title = QLabel("Загрузка документов", self)
        title.setObjectName("pageTitle")
        layout.addWidget(title)
        self.status_label = QLabel("Подготовка…", self)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        layout.addSpacing(6)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 1000)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        hint = QLabel(
            "Документы анализируются в фоне. Не закрывайте окно программы.",
            self,
        )
        hint.setObjectName("fieldHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        layout.addStretch(1)
        self.cancel_btn = QPushButton("Отменить", self)
        self.cancel_btn.clicked.connect(on_cancel)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(self.cancel_btn)
        layout.addLayout(row)

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def set_fraction(self, fraction: float) -> None:
        self.progress_bar.setValue(int(max(0.0, min(1.0, fraction)) * 1000))

    def set_cancel_enabled(self, enabled: bool) -> None:
        self.cancel_btn.setEnabled(enabled)

    def keyPressEvent(self, event) -> None:  # noqa: N802 — имя Qt
        # Escape не закрывает окно прогресса: отмена — только кнопкой,
        # закрытие окна не прерывает фоновый анализ (Tk-контракт grab_set).
        if event.key() == Qt.Key.Key_Escape:
            event.ignore()
            return
        super().keyPressEvent(event)


class ImportReportDialog(QDialog):
    """Отчёт импорта: по каждому файлу тип, результат и причина (п. 3.5)."""

    def __init__(self, parent, entries: list[dict], summary: str,
                 kind_labels: dict, open_registry: Callable[[], None]) -> None:
        super().__init__(parent)
        self.setWindowTitle("Отчёт импорта")
        self.setModal(True)
        self.resize(560, 420)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 14)
        self.summary_label = QLabel(summary, self)
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)
        layout.addSpacing(6)
        self.details = QTextBrowser(self)
        lines: list[str] = []
        for entry in entries:
            kind = kind_labels.get(entry.get("kind"), "Документ")
            status = {
                "added": "добавлен",
                "duplicate": "дубликат",
                "skip": "пропущен",
            }.get(entry.get("status"), str(entry.get("status") or "—"))
            line = f"• {entry.get('name')} — {kind}, {status}"
            if entry.get("detail"):
                line += f": {entry['detail']}"
            if entry.get("warning"):
                line += f" ({entry['warning']})"
            lines.append(line)
        self.details.setPlainText("\n".join(lines) if lines else "—")
        self.details.setReadOnly(True)
        layout.addWidget(self.details, 1)
        buttons = QHBoxLayout()
        self.open_registry_btn = QPushButton("Открыть реестр", self)
        self.open_registry_btn.setDefault(True)
        self.open_registry_btn.clicked.connect(open_registry)
        buttons.addWidget(self.open_registry_btn)
        close_btn = QPushButton("Закрыть", self)
        close_btn.clicked.connect(self.accept)
        buttons.addWidget(close_btn)
        buttons.addStretch(1)
        layout.addLayout(buttons)


class ImportFlowQt:
    """Жизненный цикл пакетного импорта: guard → воркер → опрос → отчёт.

    Состояние потока импорта принадлежит контроллеру и не смешивается с
    состоянием приложения: повторный запуск не может сбросить очередь
    работающего первого (аудит 2.0.1, P1 «re-entry импорта»).
    """

    def __init__(self, app) -> None:
        self.app = app
        self.log = app.log
        # Состояние активного импорта.
        self.running = False
        self.queue: "queue.Queue | None" = None
        self.results: list[dict] = []
        self.total = 1
        self.cancelled = False
        self.cancelled_count = 0
        # Окно прогресса (существует, пока идёт импорт).
        self.win: "ImportProgressDialog | None" = None

    # ------------------------------------------------------------------ #
    #  Запуск
    # ------------------------------------------------------------------ #
    def start(self, files: list[str],
              thread_factory: Callable[..., Any] | None = None) -> None:
        """Проверяет лимиты пакета и запускает фоновый анализ файлов.

        До запуска отклоняются: повторный запуск во время работающего
        импорта, слишком большой пакет, слишком большой суммарный объём
        и ситуация «на диске не хватит места» (п. 3.4). Размер каждого
        файла дополнительно проверяется в analyze_file().

        ``thread_factory`` разрешается в момент вызова (по умолчанию —
        threading.Thread): тесты подменяют её синхронным исполнителем,
        и подмена должна действовать в момент запуска воркера.
        """
        if self.running:
            self.app.notify("Импорт уже выполняется — дождитесь завершения.")
            return
        files = [str(path) for path in files if path]
        if not files:
            return
        problem = self._pack_problem(files)
        if problem is not None:
            self.app.alert(self.app.window, *problem)
            return

        self.running = True
        # Очередь у каждого запуска своя и передаётся воркеру через args:
        # воркер не зависит от того, что лежит в self.queue (аудит 2.0.1).
        result_queue: queue.Queue = queue.Queue()
        self.queue = result_queue
        self.results = []
        self.total = max(1, len(files))
        self.cancelled = False
        self.cancelled_count = 0

        self.win = ImportProgressDialog(
            self.app.window, self.cancel,
        )
        self.win.show()

        factory = thread_factory or threading.Thread
        factory(
            target=self.worker, args=(files, result_queue), daemon=True,
        ).start()
        self.app.add_poll_timer("import_tick", 80, self.tick)

    def _pack_problem(self, files: list[str]) -> "tuple[str, str, str] | None":
        """Проверяет пакет до запуска; возвращает (заголовок, текст, вид) или None.

        Порядок проверок прежний: количество файлов, суммарный объём,
        доступность папки документов, свободное место на диске.
        """
        if len(files) > importer.MAX_FILES_PER_BATCH:
            return (
                "Слишком много файлов",
                f"За один раз можно добавить не более "
                f"{importer.MAX_FILES_PER_BATCH} файлов, выбрано {len(files)}.\n\n"
                "Разбейте загрузку на несколько частей.",
                "warning",
            )

        total_size = 0
        for path in files:
            try:
                total_size += os.path.getsize(path)
            except OSError:
                pass
        if total_size > importer.MAX_TOTAL_SIZE:
            return (
                "Слишком большой пакет",
                f"Суммарный объём файлов {total_size / (1024 * 1024):.0f} МБ "
                f"превышает допустимые "
                f"{importer.MAX_TOTAL_SIZE / (1024 * 1024):.0f} МБ.\n\n"
                "Разбейте загрузку на несколько частей.",
                "warning",
            )

        try:
            target_dir = self.documents_dir_for_import()
        except OSError as exc:
            return ("Папка документов недоступна", str(exc), "error")

        try:
            free = shutil.disk_usage(target_dir).free
        except OSError:
            free = None
        if free is not None and free < total_size + importer.MIN_FREE_SPACE:
            return (
                "Недостаточно места на диске",
                "Для импорта нужно около "
                f"{(total_size + importer.MIN_FREE_SPACE) / (1024 * 1024):.0f} МБ "
                f"свободного места, доступно {free / (1024 * 1024):.0f} МБ.\n\n"
                "Освободите место или выберите другую папку документов "
                "в «Настройках».",
                "warning",
            )
        return None

    def documents_dir_for_import(self) -> str:
        """Папка документов для фонового импорта: без диалогов.

        Импорт выполняется в отдельном потоке, где нельзя показывать
        окна. Если выбранная пользователем папка недоступна,
        возвращается ошибка — она попадёт в отчёт импорта, а не
        приведёт к молчаливой подмене пути (Tk documents_dir_for_import).
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
                    "Выберите другую папку в «Настройках» и повторите импорт."
                ) from exc
        import shell_open

        base = os.path.join(shell_open.documents_known_folder(), "EW.DocHub")
        try:
            os.makedirs(base, exist_ok=True)
        except OSError:
            pass
        return base

    # ------------------------------------------------------------------ #
    #  Окно прогресса
    # ------------------------------------------------------------------ #
    def cancel(self) -> None:
        """Отменяет импорт: остаток пакета не анализируется (п. 5.3)."""
        if not self.running:
            return
        self.cancelled = True
        try:
            self.win.set_status("Отмена…")
            self.win.set_cancel_enabled(False)
        except RuntimeError:
            pass  # окно могло быть разрушено при закрытии программы

    # ------------------------------------------------------------------ #
    #  Фоновый воркер
    # ------------------------------------------------------------------ #
    def worker(self, files: list[str], result_queue: "queue.Queue") -> None:
        """Фоновый поток: анализ каждого файла и копирование.

        После нажатия «Отменить» остаток пакета не анализируется (п. 5.3).
        Результаты кладутся в свою очередь (args), а не в общий атрибут —
        повторный запуск импорта не может смешать два пакета (аудит 2.0.1).
        """
        for index, path in enumerate(files):
            if self.cancelled:
                break
            name = os.path.basename(path)
            result_queue.put(
                ("status", f"Анализ файла {index + 1} из {len(files)}:\n{name}")
            )
            ext = os.path.splitext(path)[1].lower()
            if ext not in ALLOWED_EXTENSIONS:
                # Allowlist: .exe/.bat/.lnk и прочее не анализируем и не копируем.
                result_queue.put(("file", {
                    "path": path,
                    "name": name,
                    "kind": storage.KIND_UNKNOWN,
                    "number": "",
                    "date": "",
                    "party": None,
                    "text": "",
                    "error": f"Тип файла не поддерживается: {ext or 'без расширения'}",
                    "warning": "",
                    "recognized": False,
                    "ok": False,
                    "copied": "",
                }))
                continue
            result = importer.analyze_file(path)
            try:
                result["copied"] = self.copy_imported_file(result)
            except Exception as exc:
                result["copied"] = ""
                if not result.get("error"):
                    result["error"] = f"Не удалось скопировать файл: {exc}"
            result_queue.put(("file", result))
        result_queue.put(("done",))

    def copy_imported_file(self, result: dict) -> str:
        """Копирует в папку контрагента только разобранные .docx/.xlsx.

        Неопознанные и потенциально опасные файлы не копируются: раньше
        они попадали в папку документов и открывались по клику
        (запуск кода).
        """
        if not result.get("ok") or not result.get("party"):
            return ""
        src = str(result.get("path") or "")
        ext = os.path.splitext(src)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return ""
        base = self.documents_dir_for_import()
        folder = services.safe_filename(
            importer.party_display_name(result["party"]), "Контрагент"
        )
        dest_dir = os.path.normpath(os.path.join(base, folder))
        if not services.is_inside_directory(base, dest_dir):
            self.log.error("Папка импортированного контрагента вне папки документов")
            return ""
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except OSError as exc:
            self.log.warning(
                "Не удалось создать папку для импорта: %s", type(exc).__name__,
            )
            raise

        stem = services.safe_filename(
            os.path.splitext(os.path.basename(src))[0], "Документ"
        )
        dest = services.fit_document_path(dest_dir, stem, ext)
        counter = 2
        while os.path.exists(dest):
            dest = services.fit_document_path(dest_dir, f"{stem} ({counter})", ext)
            counter += 1
        services.copy_with_zone_identifier(src, dest)
        try:
            size = os.path.getsize(dest)
        except OSError:
            size = 0
        self.log.info(
            "Импортированный файл скопирован: тип=%s, размер=%d байт", ext, size,
        )
        return dest

    # ------------------------------------------------------------------ #
    #  Опрос очереди из главного потока
    # ------------------------------------------------------------------ #
    def tick(self) -> None:
        """Опрос очереди фонового потока QTimer'ом (главный поток)."""
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "status":
                    self.win.set_status(item[1])
                elif kind == "file":
                    self.results.append(item[1])
                    self.win.set_fraction(len(self.results) / self.total)
                elif kind == "done":
                    results = self.results
                    win = self.win
                    self.queue = None
                    self.results = []
                    self.win = None
                    self.running = False
                    # Часть пакета могла быть пропущена по «Отмене» (п. 5.3).
                    self.cancelled_count = max(0, self.total - len(results))
                    self.app.stop_poll_timer("import_tick")
                    try:
                        win.deleteLater()
                    except RuntimeError:
                        pass
                    self.finish(results)
                    return
        except queue.Empty:
            pass
        except (RuntimeError, AttributeError):
            return  # окно программы закрывается — импорт прекращаем

    # ------------------------------------------------------------------ #
    #  U9: пакетные решения импорта (сводка вместо серии диалогов)
    # ------------------------------------------------------------------ #
    def plan_batch(self, results: list[dict]) -> "dict[int, tuple]":
        """U9 (UX-B): сводка решений перед регистрацией пакета.

        Импорт «старых» файлов без ИНН раньше превращался в серию
        модальных окон — по одному на каждый файл (UX-аудит U9). Теперь
        перед регистрацией собираются два контура решений и по каждому
        показывается ОДИН диалог сводки:

        * файлы без ИНН — «создать карточки по названиям» / «привязать к
          выбранному контрагенту» / «решить по каждому» / «пропустить все»;
        * файлы с распознанным ИНН, чьё название совпало с карточкой
          другого ИНН, — «привязать к найденным» / «создать новые» /
          «по каждому» / «пропустить».

        Проверка по названию обязательна (карточка U9): «создать по
        названиям» переиспользует существующую карточку с тем же
        названием и группирует файлы одного названия в ОДНУ карточку —
        дубли без ИНН не плодятся. Возвращается карта
        ``id(result) -> решение``; файлы без решения решаются по-старому
        (ask_party / подтверждение по файлу).
        """
        no_inn: list[dict] = []
        conflicts: list[tuple[dict, dict]] = []
        for result in results:
            if not result.get("ok") or not result.get("party"):
                continue  # нераспознанные файлы не спрашивают (п. 3.3)
            party = result["party"]
            fields = party.get("fields") or {}
            inn = str(fields.get(F.INN, "")).strip()
            if result.get("recognized") and inn:
                if self.app.store.find_by_inn(inn) is None:
                    same_name = self.app.store.find_by_name(party.get("display"))
                    if same_name is not None:
                        conflicts.append((result, same_name))
            else:
                no_inn.append(result)

        decisions: "dict[int, tuple]" = {}
        cache: "dict[str, dict]" = {}
        if no_inn:
            choice = self.ask_batch_no_inn(no_inn)
            if choice == "create_by_name":
                for result in no_inn:
                    display = str(
                        (result.get("party") or {}).get("display") or ""
                    ).strip()
                    if display:
                        decisions[id(result)] = ("create_by_name", cache)
                    # Название не распознано: создавать «пустую» карточку
                    # нельзя — файл решается по-старому (ask_party).
            elif choice == "attach_current":
                rec = self.current_party()
                for result in no_inn:
                    decisions[id(result)] = (
                        ("attach", rec) if rec is not None else ("skip",)
                    )
            elif choice == "skip_all":
                for result in no_inn:
                    decisions[id(result)] = ("skip",)
            # "per_file"/закрытие окна: решений нет — ask_party по каждому.

        if conflicts:
            choice = self.ask_batch_conflicts(conflicts)
            if choice == "attach_all":
                for result, same_name in conflicts:
                    decisions[id(result)] = ("attach", same_name)
            elif choice == "create_all":
                for result, _same in conflicts:
                    decisions[id(result)] = ("create",)
            elif choice == "skip_all":
                for result, _same in conflicts:
                    decisions[id(result)] = ("skip",)
            # "per_file": подтверждение «Похожее название уже есть» по файлу.
        return decisions

    def ask_batch_no_inn(self, results: list[dict]) -> "str | None":
        """Один диалог сводки для всех файлов без ИНН (критерий U9: 1–2
        диалога вместо 20+). Возвращает выбор или None (окно закрыто)."""
        groups: "dict[str, int]" = {}
        unnamed = 0
        for result in results:
            display = str((result.get("party") or {}).get("display") or "").strip()
            if display:
                groups[display] = groups.get(display, 0) + 1
            else:
                unnamed += 1
        lines = [f"• {name} — файлов: {count}" for name, count in groups.items()]
        if unnamed:
            lines.append(f"• название не распознано — файлов: {unnamed}")
        message = (
            f"Импорт: {len(results)} файлов без ИНН.\n\n"
            + "\n".join(lines)
            + "\n\nПри создании по названиям существующие карточки с тем же "
            "названием переиспользуются — дубли не создаются; файлы одного "
            "названия попадают в одну карточку."
        )
        options: "list[tuple[str, str, bool]]" = [
            ("create_by_name", "Создать карточки по названиям", True),
        ]
        current = self.current_party()
        if current is not None:
            options.append((
                "attach_current",
                f"Привязать все к «{storage.display_name(current)}»", False,
            ))
        options.append(("per_file", "Решить по каждому файлу", False))
        options.append(("skip_all", "Пропустить все", False))
        return self.app.choice(
            self.app.window, "Импорт без ИНН", message, options,
        )

    def ask_batch_conflicts(self, conflicts: "list[tuple[dict, dict]]") -> "str | None":
        """Один диалог для файлов, чьё название занято карточкой с другим ИНН."""
        lines = [
            f"• {result.get('name')} — распознано "
            f"«{str((result.get('party') or {}).get('display') or '')}», "
            f"в реестре «{storage.display_name(same_name)}»"
            for result, same_name in conflicts
        ]
        message = (
            f"Импорт: {len(conflicts)} файлов с совпадающим названием.\n\n"
            + "\n".join(lines)
            + "\n\nПривязать документы к найденным карточкам или создать новые?"
        )
        options: "list[tuple[str, str, bool]]" = [
            ("attach_all", "Привязать к найденным", True),
            ("create_all", "Создать новые карточки", False),
            ("per_file", "Решить по каждому файлу", False),
            ("skip_all", "Пропустить все", False),
        ]
        return self.app.choice(
            self.app.window, "Совпадение названий", message, options,
        )

    def _apply_decision(self, decision: tuple, party: dict) -> "tuple[dict | None, str]":
        """Преобразует решение сводки U9 в (запись, действие) регистрации.

        ``create_by_name`` держит кэш названий: два файла одного названия
        получают одну карточку; совпадение с реестром переиспользуется
        (проверка по названию — обязательное условие карточки U9).
        """
        kind = decision[0]
        if kind == "attach":
            return decision[1], "attach"
        if kind == "skip":
            return None, "skip"
        if kind == "create":
            return None, "create"
        if kind == "create_by_name":
            cache = decision[1]
            display = str(party.get("display") or "").strip()
            if not display:
                return None, "skip"
            if display in cache:
                return cache[display], "attach"
            existing = self.app.store.find_by_name(display)
            if existing is not None:
                cache[display] = existing
                return existing, "attach"
            return None, "create"  # создание — в общей ветке register_result
        return None, "skip"

    # ------------------------------------------------------------------ #
    #  Регистрация результатов
    # ------------------------------------------------------------------ #
    def finish(self, results: list[dict]) -> None:
        """Регистрирует результаты анализа в реестре и показывает отчёт.

        Правила (п. 3.3, 3.5):
          * карточка создаётся автоматически только при найденном ИНН;
          * без ИНН — диалог «привязать / создать / пропустить» (U9 —
            пакетная сводка);
          * документ, который уже есть в реестре, повторно не добавляется;
          * по каждому файлу в отчёте видны тип, результат и причина.

        Пакетная запись (п. 5.2): все изменения пакета пишутся на диск
        ОДНИМ ``save()`` в конце.
        """
        stats = {
            "added": 0,
            "new_parties": 0,
            "skipped": 0,
            "duplicates": 0,
            "cancelled": max(0, int(self.cancelled_count)),
        }
        save_errors: list[str] = []
        entries: list[dict] = []
        store_dirty = False

        # U9: сводка пакетных решений перед регистрацией — один диалог
        # вместо серии (файлы без ИНН, конфликты названий).
        decisions = self.plan_batch(results)

        for result in results:
            entry = self.new_entry(result)
            entries.append(entry)
            if self.register_result(result, entry, stats, decisions):
                store_dirty = True

        if store_dirty:
            with applog.timed("Запись реестра (импорт)", 500, "storage"):
                if not self.app.store.save():
                    save_errors.append(
                        self.app.store.last_error or "реестр не записан"
                    )
            self.log.info("Пакет импорта записан одним save() (п. 5.2)")

        if stats["added"]:
            # Списки страниц документов + реестр + статус одним хвостом (M6-a).
            self.app.refresh_after_party_change()
        elif getattr(self.app, "registry_flow", None) is not None:
            self.app.registry_flow.refresh()
            self.app.refresh_status()

        message = (
            "Загрузка завершена.\n"
            f"Добавлено документов: {stats['added']} "
            f"(новых контрагентов: {stats['new_parties']})\n"
            f"Пропущено: {stats['skipped']}"
        )
        if stats["duplicates"]:
            message += f", уже были в реестре: {stats['duplicates']}"
        if stats["cancelled"]:
            message += f"\nОтменено пользователем: {stats['cancelled']}"
        if save_errors:
            message += (
                "\n\nВНИМАНИЕ: часть записей не попала на диск:\n"
                + "\n".join(f"• {err}" for err in dict.fromkeys(save_errors))
            )
        self.log.info(
            "Импорт завершён: добавлено=%d, новых контрагентов=%d, "
            "пропущено=%d, дубликатов=%d, ошибок записи=%d",
            stats["added"], stats["new_parties"], stats["skipped"],
            stats["duplicates"], len(save_errors),
        )
        self.show_report(entries, message)
        self.cancelled_count = 0

    @staticmethod
    def new_entry(result: dict) -> dict:
        """Заготовка строки отчёта импорта для одного файла."""
        return {
            "name": result.get("name") or "файл",
            "kind": result.get("kind") or storage.KIND_UNKNOWN,
            "status": "skip",
            "detail": "",
            "warning": "",
        }

    def register_result(self, result: dict, entry: dict, stats: dict,
                        decisions: "dict[int, tuple] | None" = None) -> bool:
        """Регистрирует один результат разбора. Возвращает True, если реестр менялся.

        Изменения остаются в памяти: сохранение на диск делает вызывающий
        код одним ``save()`` на весь пакет (п. 5.2). ``decisions`` — карта
        решений сводки U9 (plan_batch): файл с решением диалогов по
        одному не задаёт.
        """
        copied = result.get("copied") or ""
        party = result.get("party") or {}

        if not result.get("ok") or not party:
            stats["skipped"] += 1
            entry["detail"] = result.get("error") or "документ не распознан"
            result["copied"] = self.move_to_unknown(copied)
            return False

        fields = dict(party.get("fields") or {})
        inn = str(fields.get(F.INN, "")).strip()
        rec = None
        action = ""
        decision = (decisions or {}).get(id(result))
        if result.get("recognized") and inn:
            rec = self.app.store.find_by_inn(inn)
            if rec is not None:
                action = "attach"
            elif decision is not None:
                # U9: решение сводки вместо подтверждения «Похожее название».
                rec, action = self._apply_decision(decision, party)
            else:
                # ИНН новый, но карточка с таким названием уже может быть
                # (например, заведена вручную без ИНН) — не плодим дубль.
                same_name = self.app.store.find_by_name(party.get("display"))
                if same_name is not None:
                    attach = self.app.confirm(
                        self.app.window, "Похожее название уже есть",
                        f"Файл: {result.get('name')}\n"
                        f"Распознано: {party.get('display')}\n\n"
                        f"В реестре есть «{storage.display_name(same_name)}» "
                        "с другим ИНН.\n\n"
                        "Привязать документ к этой карточке?",
                        ok_text="Привязать", cancel_text="Создать новую",
                    )
                    if attach:
                        rec, action = same_name, "attach"
                if action != "attach":
                    action = "create"
        elif decision is not None:
            # U9: файл без ИНН с предрешённым исходом сводки.
            rec, action = self._apply_decision(decision, party)
        else:
            rec, action = self.ask_party(result)

        if action == "create":
            rec = self.app.store.add_counterparty(
                party.get("type") or "ooo", fields, save=False
            )
            stats["new_parties"] += 1
            entry["detail"] = "создана новая карточка"
            # U9: свежесозданная карточка попадает в кэш названий сводки —
            # следующие файлы того же названия привяжутся к ней (одна
            # карточка на название, дубли без ИНН не плодятся).
            if decision is not None and decision[0] == "create_by_name":
                name_cache = decision[1]
                display = str(party.get("display") or "").strip()
                if display and display not in name_cache:
                    name_cache[display] = rec
        elif action == "attach" and rec is not None:
            entry["detail"] = f"привязан к «{storage.display_name(rec)}»"

        if action == "skip" or rec is None:
            stats["skipped"] += 1
            entry["detail"] = entry["detail"] or "контрагент не выбран"
            result["copied"] = self.move_to_unknown(copied)
            return False

        if not copied:
            stats["skipped"] += 1
            entry["detail"] = result.get("error") or "файл не скопирован"
            return False

        if self.document_is_duplicate(rec["id"], result):
            stats["duplicates"] += 1
            entry["status"] = "duplicate"
            entry["detail"] = "такой документ уже есть в реестре"
            result["copied"] = self.move_to_unknown(copied)
            return False

        self.app.store.add_document(
            result.get("kind") or storage.KIND_UNKNOWN,
            rec["id"],
            result.get("number") or "",
            result.get("date") or "",
            data=dict(rec.get("fields") or {}),
            extra={
                "imported": True,
                "source_file": copied,
                "is_ip": rec.get("type") == "ip",
            },
            text=result.get("text") or "",
            save=False,
        )
        stats["added"] += 1
        entry["status"] = "added"
        if not result.get("recognized"):
            entry["warning"] = (
                result.get("warning") or "карточка создана без ИНН"
            )
        return True

    def ask_party(self, result: dict) -> "tuple[dict | None, str]":
        """Спрашивает, к какому контрагенту отнести документ без ИНН (п. 3.3).

        Возвращает ``(запись, действие)``: действие — ``attach`` /
        ``create`` / ``skip``. Сначала предлагается найденный по названию
        контрагент, затем — выбранный на страницах документов, затем
        создание карточки.
        """
        party = result.get("party") or {}
        fields = party.get("fields") or {}
        display = str(party.get("display") or "").strip() or "название не распознано"
        missing = party.get("missing") or "ИНН"

        existing = self.app.store.find_by_name(display) if display else None
        if existing is not None:
            existing_inn = str(
                (existing.get("fields") or {}).get(F.INN, "")
            ).strip() or "не указан"
            attach = self.app.confirm(
                self.app.window, "Похожий контрагент уже есть",
                f"Файл: {result.get('name')}\n"
                f"Распознано: {display}\n\n"
                f"В реестре есть «{storage.display_name(existing)}» "
                f"(ИНН {existing_inn}).\n\n"
                "Привязать документ к этой карточке?",
                ok_text="Привязать", cancel_text="Создать новую",
            )
            if attach:
                return existing, "attach"
            choice = self.app.choice(
                self.app.window,
                "Новый контрагент",
                f"Создать новую карточку «{display}» без {missing}?",
                [("create", "Создать карточку", True),
                 ("skip", "Пропустить", False)],
            )
            return (None, "create") if choice == "create" else (None, "skip")

        current = self.current_party()
        options: "list[tuple[str, str, bool]]" = []
        if current is not None:
            options.append((
                "attach", f"Привязать к «{storage.display_name(current)}»", False,
            ))
        options.append(("create", "Создать карточку", True))
        options.append(("skip", "Пропустить", False))
        choice = self.app.choice(
            self.app.window,
            "Не найден ИНН",
            f"Файл: {result.get('name')}\n"
            f"Распознано: {display}\n\n"
            f"Не найден {missing}: автоматически карточку создать нельзя.\n"
            "Выберите, что сделать с документом:",
            options,
        )
        if choice == "attach":
            return current, "attach"
        if choice == "create":
            return None, "create"
        return None, "skip"

    def current_party(self) -> "dict | None":
        """Контрагент, выбранный на страницах документов (Tk-порядок).

        Порядок Tk current_party: договор → доп. соглашение →
        спецификация; у каждого потока свой комбо контрагента.
        """
        for flow in (getattr(self.app, "contract_flow", None),
                     getattr(self.app, "additional_flow", None),
                     getattr(self.app, "spec_flow", None)):
            if flow is None:
                continue
            try:
                rec = flow.selected_party()
            except Exception:
                continue
            if rec is not None:
                return rec
        return None

    def document_is_duplicate(self, party_id, result: dict) -> bool:
        """Такой документ уже есть у контрагента (тип + номер + дата)."""
        kind = result.get("kind") or ""
        number = str(result.get("number") or "").strip()
        date_text = str(result.get("date") or "").strip()
        if not kind or (not number and not date_text):
            return False  # без номера и даты сравнивать нечего
        for rec in self.app.store.documents_for(party_id):
            if rec.get("kind") != kind:
                continue
            if str(rec.get("number") or "").strip() != number:
                continue
            if str(rec.get("date") or "").strip() != date_text:
                continue
            return True
        return False

    def move_to_unknown(self, path: str) -> str:
        """Переносит отклонённый файл в папку «unknown» (п. 3.5).

        Документ, который не попал в реестр, не должен оставаться в папке
        контрагента — иначе он выглядит как осиротевший файл.
        """
        if not path:
            return path
        try:
            if not os.path.exists(path):
                return path
            base = self.documents_dir_for_import()
            unknown_dir = os.path.join(base, importer.UNKNOWN_FOLDER)
            os.makedirs(unknown_dir, exist_ok=True)
            ext = os.path.splitext(path)[1]
            stem = services.safe_filename(
                os.path.splitext(os.path.basename(path))[0], "Документ"
            )
            dest = services.fit_document_path(unknown_dir, stem, ext)
            counter = 2
            while os.path.exists(dest):
                dest = services.fit_document_path(
                    unknown_dir, f"{stem} ({counter})", ext
                )
                counter += 1
            os.replace(path, dest)
            return dest
        except OSError as exc:
            self.log.warning(
                "Не удалось перенести файл в unknown: %s", type(exc).__name__,
            )
            return path

    # ------------------------------------------------------------------ #
    #  Отчёт
    # ------------------------------------------------------------------ #
    def show_report(self, entries: list[dict], summary: str) -> None:
        """Окно отчёта: по каждому файлу тип, результат и причина (п. 3.5).

        Кнопка «Открыть реестр» (U10) ведёт в обновлённый реестр.
        Подменяется в smoke/тестах на уровне контроллера (маршруты
        подмен — риски M6 плана).
        """
        dialog = ImportReportDialog(
            self.app.window, entries, summary, storage.KIND_LABELS,
            self.open_registry,
        )
        dialog.exec()

    def open_registry(self) -> None:
        """U10: «Открыть реестр» из отчёта — обновить реестр и показать его."""
        if getattr(self.app, "registry_flow", None) is not None:
            self.app.registry_flow.refresh()
        self.app.refresh_status()
        self.app.select_page("registry")
