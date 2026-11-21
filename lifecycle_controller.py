"""Жизненный цикл записей реестра — третий контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет сценариями реестра «экспорт /
открытие документа» и «корзина»: подтверждения с описанием последствий,
перемещение записей в корзину, восстановление и окончательная очистка.

Контроллер получает ссылку на приложение и обращается к общим службам через
него (``self.app.store``, ``self.app.root`` …), а вызовы, которые подменяет
smoke-тест, всегда идут через ``self.app.*``:

* ``_choice_dialog`` подменяется как instance-атрибут приложения;
* открытие файла и показ папки идут через фасады приложения
  ``ContractApp._open_file`` / ``ContractApp._show_in_folder`` — их тела
  резолвят модульные ``main.open_file`` / ``main.show_in_folder`` на каждый
  вызов, а smoke-тест подменяет именно модульные имена;
* ``ui.confirm`` / ``ui.alert`` — атрибуты общего модуля ui_macos, подмена
  ``m.ui.confirm`` видна контроллеру напрямую.

Собственного состояния у контроллера нет: каждая операция завершается
синхронно и заканчивается обновлением UI через приложение (реестр, статус,
предпросмотры, комбобоксы контрагентов). Данными владеет DataStore, диалогами
— слой оформления. Поведение один в один повторяет код, вынесенный из
main.py версии 2.3.0, включая честные подтверждения аудита 2.0.1 (п. 2.5,
2.6, 2.9): удаление без подтверждения невозможно, утерянный оригинал
импортированного документа не подменяется молча.
"""

from __future__ import annotations

import os
from tkinter import filedialog
from typing import Any

import services
import storage
import ui_macos as ui


class LifecycleController:
    """Экспорт/открытие документов реестра и корзина (п. 2.5, 2.6, 2.9).

    Без состояния: контроллер только оркестрирует подтверждения и вызовы
    DataStore, после каждой успешной операции запрашивает обновление UI.
    """

    def __init__(self, app: Any) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Реестр: экспорт и открытие документа
    # ------------------------------------------------------------------ #
    def export_document(self, rec: dict):
        """Открывает документ из реестра.

        Импортированные документы открываются как исходный файл из папки
        контрагента; созданные в программе — собираются из шаблона .docx.
        Если оригинал импортированного документа потерян, пользователь явно
        выбирает: показать папку или пересобрать из шаблона (п. 2.9).
        """
        try:
            extra = rec.get("extra") or {}
            source = str(extra.get("source_file") or "")
            if rec.get("kind") == storage.KIND_UNKNOWN:
                # «Неопознанные» не открываем через os.startfile — только папку.
                self.app._show_in_folder(source)
                return
            if extra.get("imported"):
                if source and os.path.exists(source):
                    self.app._open_file(source)
                    return
                if not self._confirm_missing_source(rec, source):
                    return
                # Пользователь разрешил пересобрать документ из шаблона.

            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            party_rec = self.app.store.get_counterparty(rec.get("party_id"))
            party_name = self.app.store.party_name(rec.get("party_id"))
            number = rec.get("number") or "б/н"
            default_filename = (
                f"{kind_label} №{number} {services.safe_filename(party_name)}.docx"
            )

            file_path = filedialog.asksaveasfilename(
                parent=self.app.root,
                defaultextension=".docx",
                filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")],
                initialfile=default_filename,
                initialdir=self.app._doc_save_flow.party_documents_dir(party_rec),
                title="Сохранить документ Word",
            )
            if not file_path:
                return

            doc = self.app._doc_save_flow.build_record_docx(rec)
            services.save_docx_atomic(doc, file_path)
            self.app._open_file(file_path)
            if extra.get("imported"):
                ui.alert(
                    self.app.root, "Документ собран из шаблона",
                    "Оригинал импортированного документа не найден, поэтому "
                    "документ собран из шаблона программы.\n\n"
                    "Сверьте его с оригиналом перед отправкой контрагенту.",
                    "warning",
                )
        except Exception as exc:
            self.app._log.warning(
                "Экспорт документа не удался: id=%s, %s",
                rec.get("id"), type(exc).__name__,
            )
            ui.alert(
                self.app.root, "Не удалось сохранить документ", str(exc), "error",
            )

    def _confirm_missing_source(self, rec: dict, source: str) -> bool:
        """Диалог при утере оригинала импортированного документа (п. 2.9)."""
        self.app._log.warning(
            "Оригинал импортированного документа не найден: id=%s", rec.get("id"),
        )
        choice = self.app._choice_dialog(
            "Оригинал не найден",
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
            self.app._show_in_folder(source or self.app._settings_flow.documents_dir())
            return False
        return choice == "rebuild"

    # ------------------------------------------------------------------ #
    #  Корзина: перемещение, восстановление, окончательная очистка
    # ------------------------------------------------------------------ #
    def delete_document(self, rec: dict):
        """Перемещает документ в корзину после подтверждения (п. 2.5, 2.6)."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        party_name = self.app.store.party_name(rec.get("party_id"))
        source = str((rec.get("extra") or {}).get("source_file") or "")
        message = (
            f"{title}\nКонтрагент: {party_name}\n\n"
            "Переместить документ в корзину?\n\n"
            "Файл на диске останется:\n"
            f"{source or '—'}\n\n"
            "Восстановить документ можно в разделе «Корзина»."
        )
        if not ui.confirm(self.app.root, "Удалить документ", message, ok_text="В корзину",
                          destructive=True):
            return
        if not self.app.store.trash_document(rec.get("id")):
            ui.alert(
                self.app.root, "Не удалось удалить документ",
                self.app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self.app._registry_flow.refresh()
        self.app._update_status()
        self.app._doc_page_flow.preview_contract()
        self.app._doc_page_flow.preview_additional()

    def delete_counterparty(self, rec: dict):
        """Перемещает контрагента и его документы в корзину (п. 2.5, 2.6)."""
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = self.app.store.documents_for(party_id)
        folder = self.app._doc_save_flow.party_documents_dir(rec, interactive=False)
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
        if not ui.confirm(self.app.root, "Удалить контрагента", message, ok_text="В корзину",
                          destructive=True):
            return
        if len(docs) > 20:
            if not ui.confirm(
                self.app.root, "Подтвердите удаление",
                f"У контрагента «{name}» {len(docs)} документов.\n\n"
                "Точно переместить всё в корзину?",
                ok_text="Да, в корзину",
                destructive=True,
            ):
                return
        removed = self.app.store.trash_counterparty(party_id)
        if self.app.store.last_error:
            ui.alert(
                self.app.root, "Не удалось удалить контрагента",
                self.app.store.last_error, "error",
            )
            return
        self.app._log.info(
            "Контрагент перемещён в корзину: id=%s, документов=%d", party_id, removed,
        )
        if self.app._registry_selected_id == party_id:
            self.app._registry_selected_id = None
        # Списки контрагентов в выпадающих меню и предпросмотры устарели.
        self.app._doc_context_flow.refresh_party_combo_values()
        self.app._doc_page_flow.preview_contract()
        self.app._doc_page_flow.preview_additional()
        self.app._registry_flow.refresh()
        self.app._update_status()

    def restore_document(self, rec: dict):
        """Возвращает документ из корзины (п. 2.6)."""
        if not self.app.store.restore_document(rec.get("id")):
            ui.alert(
                self.app.root, "Не удалось восстановить документ",
                self.app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self.app._doc_context_flow.refresh_party_combo_values()
        self.app._doc_page_flow.preview_contract()
        self.app._doc_page_flow.preview_additional()
        self.app._registry_flow.refresh()
        self.app._update_status()

    def restore_counterparty(self, rec: dict):
        """Возвращает карточку контрагента и удалённые вместе с ней документы."""
        restored = self.app.store.restore_counterparty(rec.get("id"))
        if self.app.store.last_error:
            ui.alert(
                self.app.root, "Не удалось восстановить контрагента",
                self.app.store.last_error, "error",
            )
            return
        self.app._log.info(
            "Контрагент восстановлен из корзины: id=%s, документов=%d",
            rec.get("id"), restored,
        )
        self.app._doc_context_flow.refresh_party_combo_values()
        self.app._doc_page_flow.preview_contract()
        self.app._doc_page_flow.preview_additional()
        self.app._registry_flow.refresh()
        self.app._update_status()

    def purge_document(self, rec: dict):
        """Удаляет запись документа из корзины окончательно."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        if not ui.confirm(
            self.app.root, "Удалить навсегда",
            f"{title}\n\nУдалить запись окончательно?\n\n"
            "Файл на диске останется, но восстановить документ из корзины "
            "будет уже нельзя.",
            ok_text="Удалить навсегда",
            destructive=True,
        ):
            return
        if not self.app.store.purge_document(rec.get("id")):
            ui.alert(
                self.app.root, "Не удалось удалить запись",
                self.app.store.last_error or "Запись не найдена.", "error",
            )
            return
        self.app._registry_flow.refresh()
        self.app._update_status()

    def purge_counterparty(self, rec: dict):
        """Удаляет карточку контрагента и его документы окончательно."""
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = self.app.store.documents_for(party_id, include_deleted=True)
        if not ui.confirm(
            self.app.root, "Удалить навсегда",
            f"Удалить контрагента «{name}» окончательно?\n\n"
            f"Вместе с карточкой будут удалены записи документов — {len(docs)} шт.\n"
            "Файлы на диске останутся. Восстановить записи будет нельзя.",
            ok_text="Удалить навсегда",
            destructive=True,
        ):
            return
        self.app.store.purge_counterparty(party_id)
        if self.app.store.last_error:
            ui.alert(
                self.app.root, "Не удалось удалить контрагента",
                self.app.store.last_error, "error",
            )
            return
        self.app._doc_context_flow.refresh_party_combo_values()
        self.app._registry_flow.refresh()
        self.app._update_status()

    def empty_trash(self):
        """Полностью очищает корзину (записи; файлы на диске остаются)."""
        count = self.app.store.trash_count()
        if not count:
            return
        if not ui.confirm(
            self.app.root, "Очистить корзину",
            f"В корзине записей: {count}.\n\n"
            "Удалить их окончательно? Файлы документов на диске останутся, "
            "но записи восстановить будет нельзя.",
            ok_text="Очистить корзину",
            destructive=True,
        ):
            return
        removed = self.app.store.purge_all_trash()
        if not removed and self.app.store.last_error:
            ui.alert(
                self.app.root, "Не удалось очистить корзину",
                self.app.store.last_error, "error",
            )
            return
        self.app._doc_context_flow.refresh_party_combo_values()
        self.app._registry_flow.refresh()
        self.app._update_status()
