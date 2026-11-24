"""Жизненный цикл записей реестра на Qt (этап M5).

Qt-перенос сценариев ``lifecycle_controller.LifecycleController`` (Tk),
которые нужны странице «Реестр»: экспорт/открытие документа в Word и
корзина — перемещение в корзину, восстановление, окончательная очистка.
Логика сервиса не копируется: подтверждения с теми же текстами, ходы
DataStore — те же ``trash_document`` / ``trash_counterparty`` /
``restore_*`` / ``purge_*`` / ``purge_all_trash``; сборка документа —
``services.build_record_docx`` (тот же путь, что у сохранения CP1),
атомарная запись — ``services.save_docx_atomic``.

Контракты, переносимые дословно (Tk-паритет):

* удаление без подтверждения невозможно (аудит 2.0.1, п. 2.5/2.6);
  подтверждения — деструктивная семантика (Enter — «Отмена», U7);
* у контрагента с более чем 20 документами — второе подтверждение;
* файлы на диске при удалении/очистке корзины остаются, тексты говорят
  об этом честно;
* утерянный оригинал импортированного документа не подменяется молча:
  пользователь выбирает «Показать папку» / «Пересобрать из шаблона»;
  пересборка — с предупреждением сверить с оригиналом (п. 2.9);
* «неопознанные» не открываются через системное открытие файла — только
  папка (п. 2.9);
* после каждой успешной операции обновляются реестр, статус и списки
  контрагентов страниц документов (Tk refresh_party_combo_values →
  Qt spec_flow.refresh_parties; предпросмотры договоров/доп. соглашений
  появятся вместе со страницами M6).

Отличия от Tk-сборки (зафиксированы в docs/qt-migration-stage-notes.md):
диалог сохранения файла — QFileDialog вместо tkinter.filedialog; папка
документов — non-interactive ветка (как у страницы «Спецификация», M3);
интерактивный выбор папки — вместе со страницей «Настройки» (M6).
"""

from __future__ import annotations

import os

from PySide6.QtWidgets import QFileDialog

import services
import shell_open
import storage
from services import safe_filename


def documents_dir(app) -> str:
    """Папка документов: настройка либо «Документы/EW.DocHub» (п. 2.8).

    Non-interactive ветка — та же, что у страницы «Спецификация» (M3);
    интерактивный выбор папки — вместе со страницей «Настройки» (M6).
    """
    custom = str((app.settings.get("documents_dir") or "")).strip()
    if custom:
        try:
            os.makedirs(custom, exist_ok=True)
            return custom
        except OSError as exc:
            raise OSError(
                "Папка документов недоступна:\n"
                f"{custom}\n{exc}\n\n"
                "Выберите другую папку в «Настройках»."
            ) from exc
    base = os.path.join(shell_open.documents_known_folder(), "EW.DocHub")
    try:
        os.makedirs(base, exist_ok=True)
    except OSError as exc:
        raise OSError(
            "Папка документов по умолчанию недоступна:\n"
            f"{base}\n{exc}\n\nПроверьте права доступа."
        ) from exc
    return base


def party_documents_dir(app, party_rec: dict | None) -> str:
    """Индивидуальная папка контрагента внутри папки документов (п. 9.2).

    Имя папки санитизируется, итоговый путь проверяется на выход за
    пределы папки документов (п. 2.7, services.is_inside_directory).
    """
    base = documents_dir(app)
    if party_rec is None:
        return base
    folder = safe_filename(storage.display_name(party_rec), "Контрагент")
    path = os.path.normpath(os.path.join(base, folder))
    if (
        not services.is_inside_directory(base, path)
        or os.path.normcase(path) == os.path.normcase(os.path.normpath(base))
    ):
        app.log.error("Имя папки контрагента выходит за пределы папки документов")
        app.alert(
            app.window, "Недопустимое имя папки",
            "Имя контрагента не может быть использовано как папка. "
            "Документ будет сохранён в общую папку документов.",
            "warning",
        )
        return base
    try:
        os.makedirs(path, exist_ok=True)
    except OSError as exc:
        app.log.warning(
            "Не удалось создать папку контрагента: %s", type(exc).__name__,
        )
        app.alert(
            app.window, "Не удалось создать папку контрагента",
            f"{path}\n\n{exc}\n\n"
            "Документ будет сохранён в общую папку документов.",
            "warning",
        )
        return base
    return path


class LifecycleFlowQt:
    """Экспорт/открытие документов реестра и корзина (Qt-поток M5).

    Без собственного состояния: контроллер только оркестрирует
    подтверждения и вызовы DataStore; после каждой успешной операции
    запрашивает обновление реестра/статуса/списков контрагентов.
    """

    def __init__(self, app) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Реестр: экспорт и открытие документа
    # ------------------------------------------------------------------ #
    def export_document(self, rec: dict) -> None:
        """Открывает документ из реестра (Tk-паритет export_document).

        Импортированные документы открываются как исходный файл; созданные
        в программе — собираются из шаблона .docx (путь сохранения CP1).
        Утерянный оригинал — явный выбор пользователя (п. 2.9).
        """
        app = self.app
        try:
            extra = rec.get("extra") or {}
            source = str(extra.get("source_file") or "")
            if rec.get("kind") == storage.KIND_UNKNOWN:
                # «Неопознанные» не открываем системным открытием — папка.
                app.show_in_folder(source)
                return
            if extra.get("imported"):
                if source and os.path.exists(source):
                    app.open_file(source)
                    return
                if not self._confirm_missing_source(rec, source):
                    return
                # Пользователь разрешил пересобрать документ из шаблона.

            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            party_rec = app.store.get_counterparty(rec.get("party_id"))
            party_name = app.store.party_name(rec.get("party_id"))
            number = rec.get("number") or "б/н"
            default_filename = (
                f"{kind_label} №{number} {services.safe_filename(party_name)}.docx"
            )
            party_dir = party_documents_dir(app, party_rec)
            suggested = os.path.join(party_dir, default_filename)

            file_path, _filter = QFileDialog.getSaveFileName(
                app.window,
                "Сохранить документ Word",
                suggested,
                "Документы Word (*.docx);;Все файлы (*.*)",
            )
            if not file_path:
                return

            doc = services.build_record_docx(rec, app.store)
            services.save_docx_atomic(doc, file_path)
            app.open_file(file_path)
            if extra.get("imported"):
                app.alert(
                    app.window, "Документ собран из шаблона",
                    "Оригинал импортированного документа не найден, поэтому "
                    "документ собран из шаблона программы.\n\n"
                    "Сверьте его с оригиналом перед отправкой контрагенту.",
                    "warning",
                )
        except Exception as exc:
            app.log.warning(
                "Экспорт документа не удался: id=%s, %s",
                rec.get("id"), type(exc).__name__,
            )
            app.alert(
                app.window, "Не удалось сохранить документ", str(exc), "error",
            )

    def _confirm_missing_source(self, rec: dict, source: str) -> bool:
        """Диалог при утере оригинала импортированного документа (п. 2.9)."""
        app = self.app
        app.log.warning(
            "Оригинал импортированного документа не найден: id=%s", rec.get("id"),
        )
        choice = app.choice(
            app.window, "Оригинал не найден",
            "Исходный файл импортированного документа отсутствует:\n"
            f"{source or 'путь не сохранён'}\n\n"
            "Можно показать папку документов или пересобрать документ из шаблона. "
            "Пересобранный документ может отличаться от оригинала — сверьте его.",
            [
                ("cancel", "Отмена", False),
                ("folder", "Показать папку", False),
                ("rebuild", "Пересобрать из шаблона", True),
            ],
        )
        if choice == "folder":
            try:
                app.show_in_folder(source or documents_dir(app))
            except OSError as exc:
                app.alert(app.window, "Папка документов недоступна", str(exc), "error")
            return False
        return choice == "rebuild"

    # ------------------------------------------------------------------ #
    #  Корзина: перемещение, восстановление, окончательная очистка
    # ------------------------------------------------------------------ #
    def delete_document(self, rec: dict) -> None:
        """Перемещает документ в корзину после подтверждения (п. 2.5, 2.6)."""
        app = self.app
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        party_name = app.store.party_name(rec.get("party_id"))
        source = str((rec.get("extra") or {}).get("source_file") or "")
        message = (
            f"{title}\nКонтрагент: {party_name}\n\n"
            "Переместить документ в корзину?\n\n"
            "Файл на диске останется:\n"
            f"{source or '—'}\n\n"
            "Восстановить документ можно в разделе «Корзина»."
        )
        if not app.confirm(
            app.window, "Удалить документ", message, ok_text="В корзину",
            destructive=True,
        ):
            return
        if not app.store.trash_document(rec.get("id")):
            app.alert(
                app.window, "Не удалось удалить документ",
                app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self._after_registry_change()

    def delete_counterparty(self, rec: dict) -> None:
        """Перемещает контрагента и его документы в корзину (п. 2.5, 2.6)."""
        app = self.app
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = app.store.documents_for(party_id)
        try:
            folder = party_documents_dir(app, rec)
        except OSError:
            folder = "—"
        lines = []
        for doc in docs[:10]:
            kind_label = storage.KIND_LABELS.get(doc.get("kind"), "Документ")
            number = doc.get("number") or "б/н"
            date_text = doc.get("date") or ""
            lines.append(
                f"• {kind_label} № {number}" + (f" от {date_text}" if date_text else "")
            )
        if len(docs) > 10:
            lines.append(f"• …и ещё {len(docs) - 10}")
        listing = "\n".join(lines) if lines else "• документов нет"
        message = (
            f"Удалить контрагента «{name}»?\n\n"
            f"В корзину будут перемещены карточка и документы ({len(docs)} шт.):\n"
            f"{listing}\n\n"
            "Файлы на диске останутся в папке:\n"
            f"{folder}\n\n"
            "Восстановить можно в разделе «Корзина»."
        )
        if not app.confirm(
            app.window, "Удалить контрагента", message, ok_text="В корзину",
            destructive=True,
        ):
            return
        if len(docs) > 20:
            if not app.confirm(
                app.window, "Подтвердите удаление",
                f"У контрагента «{name}» {len(docs)} документов.\n\n"
                "Точно переместить всё в корзину?",
                ok_text="Да, в корзину",
                destructive=True,
            ):
                return
        removed = app.store.trash_counterparty(party_id)
        if app.store.last_error:
            app.alert(
                app.window, "Не удалось удалить контрагента",
                app.store.last_error, "error",
            )
            return
        app.log.info(
            "Контрагент перемещён в корзину: id=%s, документов=%d", party_id, removed,
        )
        if getattr(app, "registry_flow", None) is not None:
            app.registry_flow.forget_party(party_id)
        self._after_registry_change()

    def restore_document(self, rec: dict) -> None:
        """Возвращает документ из корзины (п. 2.6)."""
        app = self.app
        if not app.store.restore_document(rec.get("id")):
            app.alert(
                app.window, "Не удалось восстановить документ",
                app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self._after_registry_change()

    def restore_counterparty(self, rec: dict) -> None:
        """Возвращает карточку контрагента и удалённые вместе с ней документы."""
        app = self.app
        restored = app.store.restore_counterparty(rec.get("id"))
        if app.store.last_error:
            app.alert(
                app.window, "Не удалось восстановить контрагента",
                app.store.last_error, "error",
            )
            return
        app.log.info(
            "Контрагент восстановлен из корзины: id=%s, документов=%d",
            rec.get("id"), restored,
        )
        self._after_registry_change()

    def purge_document(self, rec: dict) -> None:
        """Удаляет запись документа из корзины окончательно."""
        app = self.app
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        if not app.confirm(
            app.window, "Удалить навсегда",
            f"{title}\n\nУдалить запись окончательно?\n\n"
            "Файл на диске останется, но восстановить документ из корзины "
            "будет уже нельзя.",
            ok_text="Удалить навсегда",
            destructive=True,
        ):
            return
        if not app.store.purge_document(rec.get("id")):
            app.alert(
                app.window, "Не удалось удалить запись",
                app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self._after_registry_change()

    def purge_counterparty(self, rec: dict) -> None:
        """Удаляет карточку контрагента и его документы окончательно."""
        app = self.app
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = app.store.documents_for(party_id, include_deleted=True)
        if not app.confirm(
            app.window, "Удалить навсегда",
            f"Удалить контрагента «{name}» окончательно?\n\n"
            f"Вместе с карточкой будут удалены записи документов — {len(docs)} шт.\n"
            "Файлы на диске останутся. Восстановить записи будет нельзя.",
            ok_text="Удалить навсегда",
            destructive=True,
        ):
            return
        app.store.purge_counterparty(party_id)
        if app.store.last_error:
            app.alert(
                app.window, "Не удалось удалить контрагента",
                app.store.last_error, "error",
            )
            return
        self._after_registry_change()

    def empty_trash(self) -> None:
        """Полностью очищает корзину (записи; файлы на диске остаются)."""
        app = self.app
        count = app.store.trash_count()
        if not count:
            return
        if not app.confirm(
            app.window, "Очистить корзину",
            f"В корзине записей: {count}.\n\n"
            "Удалить их окончательно? Файлы документов на диске останутся, "
            "но записи восстановить будет нельзя.",
            ok_text="Очистить корзину",
            destructive=True,
        ):
            return
        removed = app.store.purge_all_trash()
        if not removed and app.store.last_error:
            app.alert(
                app.window, "Не удалось очистить корзину",
                app.store.last_error, "error",
            )
            return
        self._after_registry_change()

    # ------------------------------------------------------------------ #
    #  Обновление интерфейса после операции
    # ------------------------------------------------------------------ #
    def _after_registry_change(self) -> None:
        """Реестр + статус + списки контрагентов (Tk-паритет хвостов операций).

        Tk после тех же операций вызывает refresh_party_combo_values +
        предпросмотры договоров/доп. соглашений + refresh реестра + статус.
        Предпросмотры страниц «Договор»/«Доп. соглашение» появятся вместе
        с ними (M6); сейчас обновляются реестр, статус и комбо
        «Спецификации».
        """
        app = self.app
        if getattr(app, "registry_flow", None) is not None:
            app.registry_flow.refresh()
        if getattr(app, "spec_flow", None) is not None:
            app.spec_flow.refresh_parties()
        app.refresh_status()
