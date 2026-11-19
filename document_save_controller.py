"""Спецификация и сохранение документов — пятый контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет страницей «Спецификация» и
сценариями сохранения всех трёх типов документов: сборка страницы и
таблицы позиций, предпросмотр спецификации, подтверждение повтора номера,
транзакция «файл + запись реестра» через сервисный слой и печатные формы
записей реестра.

Контроллер получает ссылку на приложение и обращается к общим службам через
него (``self.app.store``, ``self.app.root`` …), а вызовы, которые подменяет
smoke-тест, всегда идут через ``self.app.*``:

* открытие файла после сохранения — через фасад ``ContractApp._open_file``:
  его тело резолвит модульное ``main.open_file`` на каждый вызов, а
  smoke-тест подменяет именно модульное имя;
* ``ui.confirm`` / ``ui.alert`` — атрибуты общего модуля ui_macos, подмена
  ``m.ui.confirm`` видна контроллеру напрямую.

Состояние страницы остаётся на приложении с прежними именами (``spec_text``,
``spec_table``, ``spec_party_var`` …, ``_spec_preview_after``): его читают и
меняют страницы договора и доп. соглашения, защита несохранённых данных
(п. 4.2, 4.4) и smoke-тест — контроллер владеет кодом, а не данными
страницы. Ширины разметки (``main.ENTRY_WIDTH`` и соседние) и печатные
настройки спецификации (``main.get_spec_settings``) берутся поздним импортом
main: модуль main импортирует этот контроллер, поэтому верхнеуровневый
импорт дал бы цикл, а к моменту любого вызова метода main уже полностью
инициализирован.

Поведение один в один повторяет код, вынесенный из main.py версии 2.5.0,
включая подтверждение повтора номера (аудит 2.0.1, п. 4.1) и транзакцию
сохранения (п. 9.2): при сбое записи реестра файл удаляется, запись не
добавляется — пользователь видит честное сообщение, а не «полусохранённый»
документ.
"""

from __future__ import annotations

import os
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Any

import customtkinter as ctk
import tkinter as tk

import services
import storage
import ui_macos as ui
from document_builder import (
    F, MONTHS_GENITIVE, digits_only, format_quantity, format_amount,
    amount_in_words_ru, prices_valid_until_default, spec_vat_amount_from_items,
)
from services import safe_filename, validate_contract_fields
from widgets import SpecTable


class DocumentSaveController:
    """Страница «Спецификация» и сохранение документов в реестр.

    Контроллер без собственного состояния: все виджеты и переменные
    публикуются на приложении под прежними именами, отложенный предпросмотр
    хранится в ``app._spec_preview_after`` (его сбрасывает smoke-тест).
    """

    def __init__(self, app: Any) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Страница «Спецификация»: построение
    # ------------------------------------------------------------------ #
    def build_spec_page(self):
        """Строит страницу «Спецификация»: таблица позиций + карточки слева."""
        app = self.app
        import main  # поздний импорт: константы разметки main.py (см. докстринг)

        _page, left, right, textbox = app._make_page(
            "spec", "Спецификация",
            "Позиции поставки к договору; справа — таблица позиций и текст документа.",
            with_text=True, text_title="Предпросмотр спецификации (только чтение)",
        )
        app.spec_text = textbox

        # -- Правая область: таблица позиций + предпросмотр ---------------- #
        # Заголовок правой области убираем — в табличной части только строки.
        top_label = right.grid_slaves(row=0, column=0)
        if top_label:
            top_label[0].grid_remove()
        textbox.grid_forget()
        right.grid_rowconfigure(1, weight=0)
        right.grid_rowconfigure(2, weight=1)

        app.spec_table = SpecTable(right, scale=app.scale)
        app.spec_table.grid(row=1, column=0, sticky="ew",
                            padx=int(12 * app.scale), pady=(int(4 * app.scale), 0))

        textbox.grid(row=2, column=0, sticky="nsew",
                     padx=int(12 * app.scale), pady=(int(4 * app.scale), int(12 * app.scale)))
        app.spec_table.on_change = self.on_spec_item_changed

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new",
                  padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Контрагент")

        app.spec_party_var = tk.StringVar(value="")
        app.spec_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(main.ENTRY_WIDTH * app.scale), variable=app.spec_party_var,
        )
        card.wide(app.spec_party_combo, pady=(int(2 * app.scale), int(10 * app.scale)))
        app.spec_party_var.trace_add("write", self.on_spec_party_changed)
        app.spec_party_var.trace_add("write", lambda *_: app._doc_context_flow.on_party_selected("spec"))

        # -- Договор-основание -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new",
                  padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Договор-основание")

        app.spec_contract_var = tk.StringVar(value="")
        app.spec_contract_combo = ui.option_menu(
            card.body, ["— нет сохранённых договоров —"],
            width=int(main.ENTRY_WIDTH * app.scale), variable=app.spec_contract_var,
        )
        card.wide(app.spec_contract_combo, pady=(int(2 * app.scale), int(4 * app.scale)))
        app.spec_contract_info_var = tk.StringVar(value="Договор-основание не выбран.")
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=app.spec_contract_info_var,
                font=ui.font(11), text_color=ui.COLORS["text_secondary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 40) * app.scale),
            ),
            pady=(int(2 * app.scale), int(2 * app.scale)),
        )
        app.spec_contract_var.trace_add("write", self.regenerate_spec)
        app.spec_contract_var.trace_add("write", lambda *_: app._doc_context_flow.update_contract_info("spec"))
        app.spec_contract_var.trace_add("write",
                                        lambda *_: app._on_doc_field_changed("spec", F.NUMBER))

        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new",
                  padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Спецификация")

        app.spec_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(main.SHORT_WIDTH * app.scale),
            textvariable=app.spec_number_var,
        )
        app._page_widgets[("spec", F.NUMBER)] = number
        card.row("Номер спецификации", number, required=True,
                 error_var=app._error_var_for("spec", F.NUMBER))
        app.spec_number_var.trace_add("write",
                                      lambda *_: app._on_doc_field_changed("spec", F.NUMBER))

        app.spec_day_var = tk.StringVar()
        app.spec_month_var = tk.StringVar(value=MONTHS_GENITIVE[datetime.now().month - 1])
        app._add_date_row(card, "spec", app.spec_day_var, app.spec_month_var)
        app.spec_day_var.trace_add("write", self.regenerate_spec)
        app.spec_month_var.trace_add("write", self.regenerate_spec)

        # -- Условия поставки -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=3, column=0, sticky="new",
                  padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Условия поставки")

        app.spec_delivery_term_var = tk.StringVar()
        delivery_term = ui.entry(
            card.body, width=int(main.ENTRY_WIDTH * app.scale),
            textvariable=app.spec_delivery_term_var,
        )
        app._page_widgets[("spec", F.DELIVERY_TERM)] = delivery_term
        card.row(F.DELIVERY_TERM, delivery_term,
                 hint="Например: 30 дней с момента подписания договора.")

        app.spec_payment_terms_var = tk.StringVar()
        payment_terms = ui.entry(
            card.body, width=int(main.ENTRY_WIDTH * app.scale),
            textvariable=app.spec_payment_terms_var,
        )
        app._page_widgets[("spec", F.PAYMENT_TERMS)] = payment_terms
        card.row(F.PAYMENT_TERMS, payment_terms,
                 hint="Например: оплата в течение 10 банковских дней после отгрузки.")

        app.spec_delivery_basis_var = tk.StringVar()
        delivery_basis = ui.entry(
            card.body, width=int(main.ENTRY_WIDTH * app.scale),
            textvariable=app.spec_delivery_basis_var,
        )
        app._page_widgets[("spec", F.DELIVERY_BASIS)] = delivery_basis
        card.row(F.DELIVERY_BASIS, delivery_basis,
                 hint="Например: самовывоз со склада поставщика.")

        app.spec_prices_until_var = tk.StringVar(value=prices_valid_until_default())
        prices_until = ui.entry(
            card.body, width=int(main.ENTRY_WIDTH * app.scale),
            textvariable=app.spec_prices_until_var,
        )
        card.row("Цены действительны до", prices_until,
                 hint="Автоматически — последний день текущего месяца")

        app.spec_delivery_term_var.trace_add("write", self.regenerate_spec)
        app.spec_payment_terms_var.trace_add("write", self.regenerate_spec)
        app.spec_delivery_basis_var.trace_add("write", self.regenerate_spec)

        # -- Позиции (подсказка) ------------------------------------------- #

    # ------------------------------------------------------------------ #
    #  Таблица позиций
    # ------------------------------------------------------------------ #
    @property
    def spec_rows(self) -> list[dict]:
        """Строки позиций (StringVar «name», «qty», «unit», «price»)."""
        return self.app.spec_table.rows

    def add_spec_row(self, name="", qty="", price="", unit=""):
        """Совместимость: добавление строки в таблицу позиций."""
        return self.app.spec_table.add_row(
            {"name": name, "qty": qty, "unit": unit, "price": price}
        )

    def remove_spec_row(self):
        self.app.spec_table.remove_last()

    def spec_items(self) -> list[dict]:
        """Позиции с заполненным наименованием; суммы считаются на лету."""
        return self.app.spec_table.get_items()

    def spec_total(self) -> Decimal:
        return sum(
            (item["qty"] * item["price"] for item in self.spec_items()),
            Decimal("0"),
        )

    def on_spec_item_changed(self, *_args):
        self.regenerate_spec()

    def on_spec_party_changed(self, *_args):
        app = self.app
        rec = app._doc_context_flow.party_from_var(app.spec_party_var)
        contracts = app.store.contracts_for(rec["id"]) if rec else []
        if contracts:
            values = [app._doc_context_flow.contract_display(c) for c in contracts]
        else:
            values = ["— нет сохранённых договоров —"]
        app.spec_contract_combo.set_values(values)
        app.spec_contract_var.set(values[0] if contracts else "")
        if contracts:
            app.spec_contract_info_var.set(
                f"Договор-основание: {values[0]} — подставлен автоматически, проверьте."
            )
        else:
            app.spec_contract_info_var.set(
                "Договор-основание не выбран: сохраните договор поставки."
            )
        self.regenerate_spec()

    # ------------------------------------------------------------------ #
    #  Папки контрагентов и транзакция сохранения (п. 9.2)
    # ------------------------------------------------------------------ #
    def persist_document(self, draft: dict) -> dict:
        """Пишет .docx и регистрирует запись через сервисный слой (п. 9.2).

        При сбое записи реестра файл удаляется, запись не добавляется —
        результат возвращается вызывающему коду для показа сообщения.
        """
        app = self.app
        docs_dir = self.party_documents_dir(
            app.store.get_counterparty(draft.get("party_id"))
        )
        return services.save_document(app.store, draft, docs_dir)

    def party_documents_dir(self, party_rec: dict | None,
                            interactive: bool = True) -> str:
        """Индивидуальная папка контрагента внутри общей папки документов.

        Все документы реестра хранятся по пути <папка документов>/<контрагент>/.
        Имя папки санитизируется, итоговый путь проверяется на выход за
        пределы папки документов.
        """
        app = self.app
        base = app._settings_flow.documents_dir(interactive=interactive)
        if party_rec is None:
            return base
        folder = safe_filename(storage.display_name(party_rec), "Контрагент")
        path = os.path.normpath(os.path.join(base, folder))
        if (
            not app._settings_flow.is_inside_directory(base, path)
            or os.path.normcase(path) == os.path.normcase(os.path.normpath(base))
        ):
            app._log.error(
                "Имя папки контрагента выходит за пределы папки документов",
            )
            ui.alert(
                app.root, "Недопустимое имя папки",
                "Имя контрагента не может быть использовано как папка. "
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        try:
            os.makedirs(path, exist_ok=True)
        except OSError as exc:
            app._log.warning(
                "Не удалось создать папку контрагента: %s", type(exc).__name__,
            )
            ui.alert(
                app.root, "Не удалось создать папку контрагента",
                f"{path}\n\n{exc}\n\n"
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        return path

    # ------------------------------------------------------------------ #
    #  Печатные формы из шаблонов и записей реестра
    # ------------------------------------------------------------------ #
    def render_contract_docx(self, data: dict, is_ip: bool) -> "Document":
        """Договор из шаблона: печатная форма сохраняется без искажений."""
        return services.render_contract(data, is_ip)

    def render_additional_docx(self, data: dict, is_ip: bool, number: str,
                               limit_display: str, term: str) -> "Document":
        """Доп. соглашение из шаблона с подстановкой лимита и отсрочки."""
        return services.render_additional(data, is_ip, number, limit_display, term)

    def render_spec_docx(self, rec: dict) -> "Document":
        """Спецификация из шаблона (ООО/ИП) — сборка в сервисном слое (п. 9.2)."""
        import main  # поздний импорт: печатные настройки main.get_spec_settings

        return services.render_spec(rec, self.app.store, main.get_spec_settings())

    def rec_is_ip(self, rec: dict) -> bool:
        """Признак ИП у сохранённой записи (учитывая старые записи реестра)."""
        return services.rec_is_ip(rec, self.app.store)

    def build_record_docx(self, rec: dict) -> "Document":
        """Собирает .docx записи реестра по её типу (сервисный слой)."""
        return services.build_record_docx(rec, self.app.store)

    def write_docx(self, doc, rec: dict) -> str:
        """Пишет .docx в папку контрагента атомарно (сервисный слой, п. 9.2)."""
        app = self.app
        directory = self.party_documents_dir(
            app.store.get_counterparty(rec.get("party_id"))
        )
        return services.write_docx(doc, rec, app.store, directory)

    def save_and_open_document(self, rec: dict) -> str:
        """Готовит .docx в индивидуальной папке контрагента и открывает его в Word."""
        path = self.write_docx(self.build_record_docx(rec), rec)
        self.app._open_file(path)
        return path

    def persist_document_result_alert(self, result: dict):
        """Показывает сообщение при сбое транзакции сохранения (п. 9.2)."""
        ui.alert(
            self.app.root, "Не удалось сохранить документ",
            f"{result['error']}\n\n"
            "Файл удалён, запись в реестр не добавлена.",
            "error",
        )

    # ------------------------------------------------------------------ #
    #  Предпросмотр спецификации (п. 5.6: отложенная пересборка)
    # ------------------------------------------------------------------ #
    def regenerate_spec(self, *_args):
        """Откладывает пересборку предпросмотра спецификации (п. 5.6).

        Раньше предпросмотр пересобирался на каждый символ в любом из шести
        полей и в каждой ячейке таблицы позиций. Теперь, как и у договора
        и доп. соглашения, пересборка откладывается на 120 мс.
        """
        app = self.app
        pending = getattr(app, "_spec_preview_after", None)
        if pending:
            try:
                app.root.after_cancel(pending)
            except Exception:
                pass
        app._spec_preview_after = app.root.after(120, self.render_spec_preview)

    def flush_spec_preview(self):
        """Синхронно дособирает отложенный предпросмотр (перед чтением текста)."""
        app = self.app
        pending = getattr(app, "_spec_preview_after", None)
        if pending:
            try:
                app.root.after_cancel(pending)
            except Exception:
                pass
            app._spec_preview_after = None
        self.render_spec_preview()

    def render_spec_preview(self):
        """Собирает текст предпросмотра спецификации из состояния страницы."""
        app = self.app
        import main  # поздний импорт: печатные настройки main.get_spec_settings

        app._spec_preview_after = None
        party = app._doc_context_flow.party_from_var(app.spec_party_var)
        if party is None:
            app._set_text(app.spec_text, "Выберите контрагента и договор-основание слева.")
            return
        contract = app._doc_context_flow.contract_from_var(app.spec_contract_var, party["id"])
        data = dict(party.get("fields", {}))

        lines: list[str] = []
        spec = main.get_spec_settings()
        if contract is not None:
            lines.append(
                f"Приложение № {spec['spec_appendix_number']} к Договору поставки "
                f"{app._doc_context_flow.contract_display(contract)}"
            )
        lines.append("")
        number = app.spec_number_var.get().strip() or "б/н"
        lines.append(f"СПЕЦИФИКАЦИЯ № {number}")
        lines.append(
            f"г. {spec['spec_city']}    «{app.spec_day_var.get().strip() or '__'}» "
            f"{app.spec_month_var.get().strip() or '__'} {spec['spec_year']} г."
        )
        lines.append("")
        lines.append(f"Поставщик: {spec['spec_supplier']}")
        lines.append(f"Покупатель: {storage.display_name(party)}, "
                     f"ИНН {data.get(F.INN, '')}, {data.get(F.LEGAL_ADDRESS, '')}")
        lines.append("")
        prices_until = prices_valid_until_default()
        app.spec_prices_until_var.set(prices_until)
        lines.append(f"Цены в Спецификации действительны до {prices_until}.")
        lines.append("Наименование, количество и стоимость:")

        items = self.spec_items()
        if items:
            lines.append("№   Наименование                Кол-во   Ед. изм.   Цена, руб.     Сумма, руб.")
            for index, item in enumerate(items, start=1):
                amount = (item["qty"] * item["price"]).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )
                lines.append(
                    f"{index}.  {item['name']}   —   {format_quantity(item['qty'])} "
                    f"{item['unit']} × {format_amount(item['price'])} = "
                    f"{format_amount(amount)} руб."
                )
            total = sum((i["qty"] * i["price"] for i in items), Decimal("0"))
            lines.append(
                f"Итого: {format_amount(total)} руб. ({amount_in_words_ru(total)})"
            )
            lines.append(
                f"в т.ч. НДС {spec['spec_vat_rate']}%: "
                f"{format_amount(spec_vat_amount_from_items(items, spec['spec_vat_rate']))} руб."
            )
        else:
            lines.append("Позиции не заполнены — заполните таблицу позиций справа.")
        lines.append("")
        conditions = [
            (F.DELIVERY_TERM, app.spec_delivery_term_var.get().strip()),
            (F.PAYMENT_TERMS, app.spec_payment_terms_var.get().strip()),
            (F.DELIVERY_BASIS, app.spec_delivery_basis_var.get().strip()),
        ]
        for condition_label, condition_value in conditions:
            if condition_value:
                lines.append(f"{condition_label}: {condition_value}")
        lines.append("")
        lines.append(
            f"Подписи Сторон: от Поставщика — {spec['spec_supplier_signatory']}, "
            f"от Покупателя — {data.get(F.DIRECTOR, data.get(F.IP_NAME, ''))}"
        )
        app._set_text(app.spec_text, "\n".join(lines))

    # ------------------------------------------------------------------ #
    #  Подтверждение повтора номера (п. 4.1)
    # ------------------------------------------------------------------ #
    def confirm_document_duplicate(
        self, party_id, kind: str, number: str, date_text: str
    ) -> bool:
        """Предупреждает о документе с тем же номером у того же контрагента (п. 4.1)."""
        number = str(number or "").strip()
        if not number:
            return True
        for rec in self.app.store.documents_for(party_id):
            if rec.get("kind") != kind:
                continue
            if str(rec.get("number", "")).strip() != number:
                continue
            existing_date = str(rec.get("date", "")).strip()
            title = storage.KIND_LABELS.get(kind, "Документ")
            details = f"{title} № {number}"
            if existing_date:
                details += f" от {existing_date}"
            return ui.confirm(
                self.app.root, "Повтор номера",
                f"У этого контрагента уже есть {details}.\n\n"
                "Создать ещё один документ с тем же номером?",
                ok_text="Создать ещё", cancel_text="Отмена",
            )
        return True

    # ------------------------------------------------------------------ #
    #  Сохранение документов в реестр
    # ------------------------------------------------------------------ #
    def save_contract(self):
        app = self.app
        try:
            rec = app._doc_context_flow.party_from_var(app.contract_party_var)
            if rec is None:
                ui.alert(app.root, "Контрагент не выбран",
                         "Выберите контрагента из списка или добавьте нового "
                         "через кнопку «Добавить контрагента».", "error")
                return

            data = dict(rec.get("fields", {}))
            data[F.NUMBER] = app.contract_number_var.get().strip()
            data[F.DAY] = app.contract_day_var.get().strip()
            data[F.MONTH] = app.contract_month_var.get().strip()
            is_ip = rec.get("type") == "ip"

            errors = validate_contract_fields(data, is_ip)
            if errors:
                app._clear_page_errors("contract")
                for key, message in errors.items():
                    app._set_page_error("contract", key, message)
                for key in (F.NUMBER, F.DAY):
                    app._focus_page_widget("contract", key)
                    break
                ui.alert(
                    app.root, "Ошибки в данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            date_text = f"{data[F.DAY]} {data[F.MONTH]}".strip()
            if not self.confirm_document_duplicate(
                rec["id"], storage.KIND_CONTRACT, data[F.NUMBER], date_text
            ):
                return
            extra = {"day": data[F.DAY], "month": data[F.MONTH], "is_ip": is_ip}
            draft = {
                "kind": storage.KIND_CONTRACT,
                "party_id": rec["id"],
                "number": data[F.NUMBER],
                "date": date_text,
                "data": data,
                "extra": extra,
                "text": "",
            }
            # Файл и запись реестра — транзакция сервисного слоя (п. 9.2):
            # при сбое записи реестра файл удаляется, запись не добавляется.
            result = self.persist_document(draft)
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            app._doc_context_flow.reset_form_baseline("contract")
            app._registry_flow.refresh()
            app._update_status()
            opened = app._open_file(path)
            app._notify(
                f"Договор № {data[F.NUMBER]} от {date_text} для "
                f"{storage.display_name(rec)} сохранён"
                + (" и открыт." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(app.root, "Ошибка", f"Не удалось сохранить договор:\n{exc}", "error")

    def save_additional(self):
        app = self.app
        try:
            party = app._doc_context_flow.party_from_var(app.addl_party_var)
            if party is None:
                ui.alert(app.root, "Контрагент не выбран",
                         "Выберите контрагента из списка слева.", "error")
                return
            contract = app._doc_context_flow.contract_from_var(app.addl_contract_var, party["id"])
            if contract is None:
                ui.alert(app.root, "Договор не выбран",
                         "Выберите договор-основание. Если договоров нет — создайте "
                         "договор поставки на соответствующей странице.", "error")
                return

            limit_display = app._limit_var.get().strip()
            limit = digits_only(limit_display)
            term = app.term_var.get().strip()
            number = app.addl_number_var.get().strip() or "б/н"

            errors = {}
            if not limit:
                errors[F.LIMIT] = "Заполните поле"
            if not term:
                errors[F.TERM] = "Заполните поле"
            elif not term.isdigit():
                errors[F.TERM] = "Должно быть целое число дней"
            if errors:
                app._clear_page_errors("additional")
                for key, message in errors.items():
                    app._set_page_error("additional", key, message)
                ui.alert(
                    app.root, "Не хватает данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            contract_number = contract.get("number") or ""
            date_text = contract.get("date") or ""
            if number and number != "б/н" and not self.confirm_document_duplicate(
                party["id"], storage.KIND_ADDITIONAL, number, date_text
            ):
                return
            extra = {
                "contract_id": contract["id"],
                "contract_number": contract_number,
                "limit": limit_display,
                "limit_digits": limit,
                "term": term,
                "is_ip": party.get("type") == "ip",
            }
            draft = {
                "kind": storage.KIND_ADDITIONAL,
                "party_id": party["id"],
                "number": number,
                "date": date_text,
                "data": dict(contract.get("data") or {}),
                "extra": extra,
                "text": "",
            }
            result = self.persist_document(draft)
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            app._doc_context_flow.reset_form_baseline("additional")
            app._registry_flow.refresh()
            app._update_status()
            opened = app._open_file(path)
            app._notify(
                f"Доп. соглашение № {number} к договору № {contract_number or 'б/н'} "
                "сохранено"
                + (" и открыто." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(app.root, "Ошибка", f"Не удалось сохранить доп. соглашение:\n{exc}", "error")

    def save_spec(self):
        app = self.app
        try:
            party = app._doc_context_flow.party_from_var(app.spec_party_var)
            if party is None:
                ui.alert(app.root, "Контрагент не выбран",
                         "Выберите контрагента из списка слева.", "error")
                return
            contract = app._doc_context_flow.contract_from_var(app.spec_contract_var, party["id"])
            if contract is None:
                ui.alert(app.root, "Договор не выбран",
                         "Выберите договор-основание. Если договоров нет — создайте "
                         "договор поставки на соответствующей странице.", "error")
                return

            number = app.spec_number_var.get().strip()
            if not number:
                app._set_page_error("spec", F.NUMBER, "Заполните поле")
                app._focus_page_widget("spec", F.NUMBER)
                ui.alert(app.root, "Не хватает данных",
                         "Укажите номер спецификации.", "error")
                return

            items = self.spec_items()
            if not items:
                ui.alert(app.root, "Нет позиций",
                         "Добавьте хотя бы одну позицию с наименованием.", "error")
                return
            for item in items:
                if item["qty"] <= 0:
                    ui.alert(app.root, "Проверьте позиции",
                             f"У позиции «{item['name']}» должно быть количество больше нуля.", "error")
                    return
                if item["price"] <= 0:
                    ui.alert(app.root, "Проверьте позиции",
                             f"У позиции «{item['name']}» должна быть цена больше нуля.", "error")
                    return

            # Предпросмотр собирается с задержкой — перед сохранением
            # дособираем его синхронно, чтобы текст записи был актуальным.
            self.flush_spec_preview()
            text = app.spec_text.get("1.0", "end").strip()
            date_text = f"{app.spec_day_var.get().strip()} {app.spec_month_var.get().strip()}".strip()
            contract_number = contract.get("number") or ""
            if not self.confirm_document_duplicate(
                party["id"], storage.KIND_SPEC, number, date_text
            ):
                return

            # Decimal не сериализуется в JSON — сохраняем кол-во/цены float'ами.
            total = sum((i["qty"] * i["price"] for i in items), Decimal("0"))
            items_json = [
                {"name": i["name"], "qty": float(i["qty"]),
                 "unit": i["unit"], "price": float(i["price"])}
                for i in items
            ]
            extra = {
                "contract_id": contract["id"],
                "contract_number": contract_number,
                "items": items_json,
                "total": float(total),
                "spec_day": app.spec_day_var.get().strip(),
                "spec_month": app.spec_month_var.get().strip(),
                "prices_until": app.spec_prices_until_var.get().strip(),
                "delivery_term": app.spec_delivery_term_var.get().strip(),
                "payment_terms": app.spec_payment_terms_var.get().strip(),
                "delivery_basis": app.spec_delivery_basis_var.get().strip(),
                "is_ip": party.get("type") == "ip",
            }
            draft = {
                "kind": storage.KIND_SPEC,
                "party_id": party["id"],
                "number": number,
                "date": date_text,
                "data": dict(party.get("fields", {})),
                "extra": extra,
                "text": text,
            }
            result = self.persist_document(draft)
            if not result["ok"]:
                self.persist_document_result_alert(result)
                return
            path = result["path"]
            app._doc_context_flow.reset_form_baseline("spec")
            app._registry_flow.refresh()
            app._update_status()
            opened = app._open_file(path)
            app._notify(
                f"Спецификация № {number} к договору № {contract_number or 'б/н'} "
                "сохранена"
                + (" и открыта." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(app.root, "Ошибка", f"Не удалось сохранить спецификацию:\n{exc}", "error")
