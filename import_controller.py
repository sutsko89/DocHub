"""Оркестрация пакетного импорта документов — первый контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет логикой потока импорта: guard
повторного входа, проверки пакета до запуска, фоновый воркер, опрос очереди
из главного потока Tk, регистрация результатов в реестре и отчёт.

Контроллер получает ссылку на приложение и обращается к общим службам через
него (``self.app.store``, ``self.app.root`` …), а вызовы, которые подменяет
smoke-тест (``show_report``, ``ask_party``, ``_choice_dialog``),
всегда идут через ``self.app.*`` — подмена instance-атрибута продолжает
работать. Состояние потока импорта (``running``/``queue``/``win``…) живёт
здесь; ContractApp публикует его через свойства с прежними именами.

Поведение один в один повторяет код, вынесенный из main.py версии 2.1.0,
включая исправления аудита 2.0.1: guard в начале запуска, своя очередь
результатов у каждого запуска (передаётся воркеру через args), пакетная
запись реестра одним ``save()``.
"""

from __future__ import annotations

import os
import queue
import threading
import threading
import threading
import shutil
from collections.abc import Callable
from typing import Any

import customtkinter as ctk
import tkinter as tk

import applog
import document_importer as importer
import services
import storage
import ui_macos as ui
from document_builder import F

# Импорт принимает только документы, которые программа умеет разбирать.
# Всё остальное (в т. ч. .exe/.bat/.lnk) не копируется и не открывается.
# Канонический источник — здесь; main.py публикует алиас с прежним именем.
ALLOWED_EXTENSIONS = {".docx", ".xlsx"}


class ImportController:
    """Жизненный цикл пакетного импорта: guard → воркер → опрос → отчёт.

    Состояние потока импорта принадлежит контроллеру и не смешивается с
    состоянием приложения: повторный запуск не может сбросить очередь
    работающего первого (аудит 2.0.1, P1 «re-entry импорта»).
    """

    def __init__(self, app: Any) -> None:
        self.app = app
        # Состояние активного импорта.
        self.running = False
        self.queue: "queue.Queue | None" = None
        self.results: list[dict] = []
        self.total = 1
        self.cancelled = False
        self.cancelled_count = 0
        # Виджеты окна прогресса (существуют, пока идёт импорт).
        self.win: "ctk.CTkToplevel | None" = None
        self.progress: "ctk.CTkProgressBar | None" = None
        self.status_var: "tk.StringVar | None" = None
        self.cancel_btn: "ctk.CTkButton | None" = None

    # ------------------------------------------------------------------ #
    #  Запуск
    # ------------------------------------------------------------------ #
    def start(self, files: list[str], thread_factory: Callable[..., Any] | None = None) -> None:
        """Проверяет лимиты пакета и запускает фоновый анализ файлов.

        До запуска отклоняются: повторный запуск во время работающего
        импорта, слишком большой пакет, слишком большой суммарный объём и
        ситуация «на диске не хватит места» (п. 3.4). Размер каждого файла
        дополнительно проверяется в analyze_file().

        Аудит 2.0.1: guard стоит в начале запуска — drag-and-drop вызывает
        метод напрямую, и раньше второй запуск сбрасывал очередь/результаты
        работающего первого (риск порчи отчёта и задвоения записей).

        ``thread_factory`` разрешается в момент вызова (по умолчанию —
        ``threading.Thread`` из модуля threading): smoke-тест подменяет
        ``threading.Thread`` синхронным исполнителем, и подмена должна
        действовать в момент запуска воркера.
        """
        if self.running:
            self.app._notify("Импорт уже выполняется — дождитесь завершения.")
            return
        files = [str(path) for path in files if path]
        if not files:
            return
        problem = self._pack_problem(files)
        if problem is not None:
            ui.alert(self.app.root, *problem)
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

        win = ctk.CTkToplevel(self.app.root)
        win.title("Загрузка документов")
        win.resizable(False, False)
        win.transient(self.app.root)
        self.win = win

        scale = self.app.scale
        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True,
                   padx=int(28 * scale), pady=int(22 * scale))
        ctk.CTkLabel(
            frame, text="Загрузка документов", font=ui.font(16, "bold"),
            text_color=ui.COLORS["text"], anchor="w",
        ).pack(anchor="w")
        self.status_var = tk.StringVar(value="Подготовка…")
        ctk.CTkLabel(
            frame, textvariable=self.status_var, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left",
            wraplength=int(360 * scale),
        ).pack(anchor="w", pady=(int(8 * scale), int(12 * scale)))
        self.progress = ctk.CTkProgressBar(frame, width=int(360 * scale))
        self.progress.set(0)
        self.progress.pack(fill="x")
        ctk.CTkLabel(
            frame,
            text="Документы анализируются в фоне. Не закрывайте окно программы.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left", wraplength=int(360 * scale),
        ).pack(anchor="w", pady=(int(10 * scale), 0))
        self.cancel_btn = ui.ghost_button(
            frame, "Отменить", self.cancel,
            width=int(120 * scale), height=int(28 * scale),
        )
        self.cancel_btn.pack(anchor="e", pady=(int(8 * scale), 0))

        # Окно по центру родительского, поверх остальных.
        win.update_idletasks()
        width = int(430 * scale)
        height = int(240 * scale)
        x = self.app.root.winfo_x() + (self.app.root.winfo_width() - width) // 2
        y = self.app.root.winfo_y() + (self.app.root.winfo_height() - height) // 2
        win.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
        try:
            win.grab_set()
        except tk.TclError:
            pass

        factory = thread_factory or threading.Thread
        factory(
            target=self.worker, args=(files, result_queue), daemon=True,
        ).start()
        self.app.root.after(80, self.tick)

    def _pack_problem(
        self, files: list[str],
    ) -> "tuple[str, str, str] | None":
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
            target_dir = self.app._settings_flow.documents_dir_for_import()
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

    # ------------------------------------------------------------------ #
    #  Окно прогресса
    # ------------------------------------------------------------------ #
    def cancel(self) -> None:
        """Отменяет импорт: остаток пакета не анализируется и не регистрируется (п. 5.3)."""
        if not self.running:
            return
        self.cancelled = True
        try:
            self.status_var.set("Отмена…")
            self.cancel_btn.configure(state="disabled")
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    #  Фоновый воркер
    # ------------------------------------------------------------------ #
    def worker(self, files: list[str], result_queue: "queue.Queue") -> None:
        """Фоновый поток: анализ каждого файла и копирование в папку контрагента.

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

        Неопознанные и потенциально опасные файлы не копируются: раньше они
        попадали в папку документов и открывались по клику (запуск кода).
        """
        if not result.get("ok") or not result.get("party"):
            return ""
        src = str(result.get("path") or "")
        ext = os.path.splitext(src)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return ""
        base = self.app._settings_flow.documents_dir_for_import()
        folder = services.safe_filename(
            importer.party_display_name(result["party"]), "Контрагент"
        )
        dest_dir = os.path.normpath(os.path.join(base, folder))
        if not self.app._settings_flow.is_inside_directory(base, dest_dir):
            self.app._log.error("Папка импортированного контрагента вне папки документов")
            return ""
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except OSError as exc:
            self.app._log.warning(
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
        self.app._log.info(
            "Импортированный файл скопирован: тип=%s, размер=%d байт", ext, size,
        )
        return dest

    # ------------------------------------------------------------------ #
    #  Опрос очереди из главного потока
    # ------------------------------------------------------------------ #
    def tick(self) -> None:
        """Опрос очереди фонового потока из главного потока Tk."""
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "status":
                    self.status_var.set(item[1])
                elif kind == "file":
                    self.results.append(item[1])
                    self.progress.set(len(self.results) / self.total)
                elif kind == "done":
                    results = self.results
                    win = self.win
                    self.queue = None
                    self.results = []
                    self.win = None
                    self.progress = None
                    self.running = False
                    # Часть пакета могла быть пропущена по «Отмене» (п. 5.3).
                    self.cancelled_count = max(0, self.total - len(results))
                    try:
                        win.grab_release()
                        win.destroy()
                    except Exception:
                        pass
                    self.finish(results)
                    return
        except queue.Empty:
            pass
        except (tk.TclError, AttributeError):
            return  # окно программы закрывается — импорт прекращаем
        self.app.root.after(80, self.tick)

    # ------------------------------------------------------------------ #
    #  Регистрация результатов
    # ------------------------------------------------------------------ #
    def finish(self, results: list[dict]) -> None:
        """Регистрирует результаты анализа в реестре и показывает отчёт.

        Правила (п. 3.3, 3.5):
          * карточка создаётся автоматически только при найденном ИНН;
          * без ИНН — диалог «привязать к существующему / создать / пропустить»;
          * документ, который уже есть в реестре, повторно не добавляется;
          * по каждому файлу в отчёте видны тип, результат и причина.

        Пакетная запись (п. 5.2): все изменения пакета пишутся на диск ОДНИМ
        ``save()`` в конце. Раньше каждое «добавить» полностью перезаписывало
        JSON с fsync и ротацией копий — импорт 200 файлов давал 200 полных
        записей реестра.
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

        for result in results:
            entry = self.new_entry(result)
            entries.append(entry)
            if self.register_result(result, entry, stats):
                store_dirty = True

        if store_dirty:
            with applog.timed("Запись реестра (импорт)", 500, "storage"):
                if not self.app.store.save():
                    save_errors.append(
                        self.app.store.last_error or "реестр не записан"
                    )
            self.app._log.info("Пакет импорта записан одним save() (п. 5.2)")

        if stats["added"]:
            self.app._doc_context_flow.refresh_party_combo_values()
            self.app._doc_page_flow.preview_contract()
            self.app._doc_page_flow.preview_additional()
        self.app._registry_flow.refresh()
        self.app._update_status()

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
        self.app._log.info(
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

    def register_result(self, result: dict, entry: dict, stats: dict) -> bool:
        """Регистрирует один результат разбора. Возвращает True, если реестр менялся.

        Изменения остаются в памяти: сохранение на диск делает вызывающий код
        одним ``save()`` на весь пакет (п. 5.2).
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
        if result.get("recognized") and inn:
            rec = self.app.store.find_by_inn(inn)
            if rec is not None:
                action = "attach"
            else:
                # ИНН новый, но карточка с таким названием уже может быть
                # (например, заведена вручную без ИНН) — не плодим дубль.
                same_name = self.app.store.find_by_name(party.get("display"))
                if same_name is not None:
                    attach = ui.confirm(
                        self.app.root, "Похожее название уже есть",
                        f"Файл: {result.get('name')}\n"
                        f"Распознано: {party.get('display')}\n\n"
                        f"В реестре есть «{storage.display_name(same_name)}» "
                        "с другим ИНН.\n\n"
                        "Привязать документ к этой карточке?",
                        "Привязать", "Создать новую",
                    )
                    if attach:
                        rec, action = same_name, "attach"
                if action != "attach":
                    action = "create"
        else:
            rec, action = self.ask_party(result)

        if action == "create":
            rec = self.app.store.add_counterparty(
                party.get("type") or "ooo", fields, save=False
            )
            stats["new_parties"] += 1
            entry["detail"] = "создана новая карточка"
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

        Возвращает ``(запись, действие)``: действие — ``attach`` / ``create`` /
        ``skip``. Сначала предлагается найденный по названию контрагент,
        затем — выбранный на страницах документов, затем создание карточки.
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
            attach = ui.confirm(
                self.app.root, "Похожий контрагент уже есть",
                f"Файл: {result.get('name')}\n"
                f"Распознано: {display}\n\n"
                f"В реестре есть «{storage.display_name(existing)}» "
                f"(ИНН {existing_inn}).\n\n"
                "Привязать документ к этой карточке?",
                "Привязать", "Создать новую",
            )
            if attach:
                return existing, "attach"
            choice = self.app._choice_dialog(
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
        choice = self.app._choice_dialog(
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
        """Контрагент, выбранный на страницах договора/доп./спецификации."""
        for attr in ("contract_party_var", "addl_party_var", "spec_party_var"):
            var = getattr(self.app, attr, None)
            if var is None:
                continue
            rec = self.app._doc_context_flow.party_from_var(var)
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
            base = self.app._settings_flow.documents_dir_for_import()
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
            self.app._log.warning(
                "Не удалось перенести файл в unknown: %s", type(exc).__name__,
            )
            return path

    # ------------------------------------------------------------------ #
    #  Отчёт
    # ------------------------------------------------------------------ #
    def show_report(self, entries: list[dict], summary: str) -> None:
        """Окно отчёта: по каждому файлу тип, результат и причина (п. 3.5).

        Тело — в ui_macos.import_report_window (примитив слоя оформления,
        перенос по рекомендации аудита 2.0.1); подписи типов документов
        передаются параметром, чтобы слой оформления не зависел от storage.
        """
        ui.import_report_window(
            self.app.root, entries, summary, self.app.scale,
            kind_labels=storage.KIND_LABELS,
        )
