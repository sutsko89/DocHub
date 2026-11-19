"""Контекст документов: единый текущий контрагент и защита несохранённых
данных — девятый контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет двумя связанными кластерами:

* **Единый «текущий контрагент» (п. 4.2).** Выпадающие списки страниц
  документов показывают одни и те же имена, выбор, сделанный на одной
  странице, переносится на остальные, а при совпадении названий в список
  добавляется ID карточки (п. 4.1). Устаревший выбор (контрагент удалён
  или переименован) честно очищается. Договор-основание на страницах
  «Доп. соглашение» и «Спецификация» показывается явно — выбранный
  договор не молчит в выпадающем списке (п. 4.2).
* **Защита несохранённых данных (п. 4.4).** Подпись формы страницы —
  «снимок» введённых данных: выбор контрагента/договора намеренно не
  входит (он переносится между страницами и не является введёнными
  данными). При сбое чтения подпись деградирует до None (аудит 2.0.1) —
  страница считается «грязной» и просит подтверждение, вместо того чтобы
  молча решить, что терять нечего. Уход со страницы с изменениями —
  только через подтверждение; Ctrl+S сохраняет текущую страницу.

Контроллер получает ссылку на приложение и обращается к общим службам
через него (``self.app.store``, ``self.app.root``, ``self.app.current_page`` …).
Состояние остаётся на приложении с прежними именами
(``_current_party_id``, ``_current_party_vars``, ``_syncing_party``,
``_form_baseline``): его читают другие контроллеры, защита окна при
закрытии и smoke-тест — контроллер владеет кодом, а не данными.

Маршруты подмен и связность:

* ``ui.confirm`` — атрибут общего модуля ui_macos: подмена ``m.ui.confirm``
  видна контроллеру напрямую;
* сохранение текущей страницы идёт через фасады приложения
  (``app.save_contract`` / ``app.save_additional`` / ``app.save_spec`` /
  ``app._cp_save``), разрешаемые на каждый вызов, — маршруты подмен
  контроллеров сохранения и карточки не обходятся;
* подпись карточки контрагента снимается через фасад ``app._cp_collect``
  (контроллер карточки, шаг 6), строки спецификации — через property
  ``app._spec_rows`` (контроллер сохранения, шаг 5).

Поведение один в один повторяет код, вынесенный из main.py версии 2.9.0.
"""

from __future__ import annotations

from typing import Any

import storage
import ui_macos as ui


class DocumentContextController:
    """Единый «текущий контрагент» и защита несохранённых данных.

    Состояние контекста (текущий контрагент, базовые подписи форм)
    живёт на приложении — контроллер получает его через ``self.app.*``
    на каждый вызов.
    """

    # Порядок страниц в списках «грязных» и в сбросе baseline.
    FORM_PAGES = ("contract", "additional", "spec", "counterparty")

    def __init__(self, app: Any) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Контрагенты: выпадающие списки
    # ------------------------------------------------------------------ #
    def party_duplicate_bases(self) -> set[str]:
        """Названия, которые встречаются у нескольких карточек (п. 4.1)."""
        counts: dict[str, int] = {}
        for rec in self.app.store.active_counterparties():
            base = storage.display_with_inn(rec)
            counts[base] = counts.get(base, 0) + 1
        return {base for base, count in counts.items() if count > 1}

    def party_display(self, rec: dict, duplicates: set[str] | None = None) -> str:
        """Имя в списке; при совпадении названий добавляется ID."""
        base = storage.display_with_inn(rec)
        if duplicates and base in duplicates:
            return f"{base} · ID {rec.get('id')}"
        return base

    def party_display_values(self) -> list[str]:
        app = self.app
        recs = app.store.active_counterparties()
        if not recs:
            return ["— контрагентов нет —"]
        duplicates = self.party_duplicate_bases()
        return [self.party_display(rec, duplicates) for rec in recs]

    def refresh_party_combo_values(self):
        """Обновляет списки и сбрасывает устаревший выбор (п. 4.2).

        Если выбранного контрагента больше нет в реестре (удалён,
        переименован), поле очищается — недействительное имя не должно
        оставаться на экране.
        """
        app = self.app
        values = self.party_display_values()
        valid = set(values)
        placeholders = {"— контрагентов нет —"}
        for combo, attr in (
            (getattr(app, "contract_party_combo", None), "contract_party_var"),
            (getattr(app, "addl_party_combo", None), "addl_party_var"),
            (getattr(app, "spec_party_combo", None), "spec_party_var"),
        ):
            if combo is not None:
                combo.set_values(values)
            var = getattr(app, attr, None)
            if var is None:
                continue
            current = var.get().strip()
            if current and current not in valid and current not in placeholders:
                var.set("")
        if app._current_party_id is not None:
            if app.store.get_counterparty(app._current_party_id) is None:
                app._current_party_id = None
        self.update_current_party_labels()

    def party_from_var(self, var) -> dict | None:
        """Возвращает запись контрагента по выбранному значению списка."""
        app = self.app
        value = var.get().strip()
        if not value:
            return None
        recs = app.store.active_counterparties()
        duplicates = self.party_duplicate_bases()
        for rec in recs:
            if self.party_display(rec, duplicates) == value:
                return rec
        # Совместимость: значение могло быть установлено программно без ID.
        for rec in recs:
            if storage.display_with_inn(rec) == value:
                return rec
        return None

    # -- Единый «текущий контрагент» (п. 4.2) ---------------------------- #
    def current_party_label(self) -> str:
        app = self.app
        if app._current_party_id is None:
            return "Документ для: контрагент не выбран"
        rec = app.store.get_counterparty(app._current_party_id)
        if rec is None:
            return "Документ для: выбранный контрагент не найден — выберите заново"
        return f"Документ для: {self.party_display(rec, self.party_duplicate_bases())}"

    def update_current_party_labels(self) -> None:
        app = self.app
        text = self.current_party_label()
        for var in getattr(app, "_current_party_vars", {}).values():
            try:
                var.set(text)
            except Exception:
                pass

    def set_current_party(self, rec: dict | None, source: str | None = None) -> None:
        """Запоминает контрагента и переносит выбор на все страницы документов."""
        app = self.app
        if not isinstance(rec, dict):
            return
        app._current_party_id = rec.get("id")
        display = self.party_display(rec, self.party_duplicate_bases())
        app._syncing_party = True
        try:
            for key, attr in (
                ("contract", "contract_party_var"),
                ("additional", "addl_party_var"),
                ("spec", "spec_party_var"),
            ):
                if key == source:
                    continue
                var = getattr(app, attr, None)
                if var is None:
                    continue
                try:
                    if var.get() != display:
                        var.set(display)
                except Exception:
                    pass
        finally:
            app._syncing_party = False
        self.update_current_party_labels()

    def on_party_selected(self, key: str) -> None:
        """Пользователь выбрал контрагента на одной из страниц документов."""
        app = self.app
        if getattr(app, "_syncing_party", False):
            return
        attr = {
            "contract": "contract_party_var",
            "additional": "addl_party_var",
            "spec": "spec_party_var",
        }.get(key)
        var = getattr(app, attr, None) if attr else None
        if var is None:
            return
        rec = self.party_from_var(var)
        if rec is None:
            if not var.get().strip():
                app._current_party_id = None
                self.update_current_party_labels()
            return
        self.set_current_party(rec, source=key)

    def update_contract_info(self, page: str) -> None:
        """Показывает выбранный договор-основание явно (п. 4.2)."""
        app = self.app
        if page == "additional":
            var = getattr(app, "addl_contract_var", None)
            info = getattr(app, "addl_contract_info_var", None)
        else:
            var = getattr(app, "spec_contract_var", None)
            info = getattr(app, "spec_contract_info_var", None)
        if var is None or info is None:
            return
        value = var.get().strip()
        if not value or value.startswith("—"):
            info.set("Договор-основание не выбран — укажите его явно.")
        else:
            info.set(f"Договор-основание: {value}")

    def contract_display(self, rec: dict) -> str:
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        return f"№ {number} от {date_text}" if date_text else f"№ {number}"

    def contract_from_var(self, var, party_id) -> dict | None:
        app = self.app
        value = var.get().strip()
        if not value:
            return None
        for rec in app.store.contracts_for(party_id):
            if self.contract_display(rec) == value:
                return rec
        return None

    # ------------------------------------------------------------------ #
    #  Несохранённые данные (п. 4.4)
    # ------------------------------------------------------------------ #
    @staticmethod
    def form_label(key: str) -> str:
        return {
            "contract": "Договор поставки",
            "additional": "Доп. соглашение",
            "spec": "Спецификация",
            "counterparty": "Карточка контрагента",
        }.get(key, key)

    def form_signature(self, key: str) -> tuple | None:
        """Данные формы, потеря которых требует подтверждения.

        Выбор контрагента/договора намеренно не входит в подпись: он
        переносится между страницами (п. 4.2) и не является введёнными данными.

        Аудит 2.0.1: при сбое чтения возвращается None, а не пустой кортеж —
        None не совпадает ни с каким baseline, поэтому защита несохранённых
        данных считает страницу «грязной» и просит подтверждение, вместо
        того чтобы молча решить, что терять нечего.
        """
        app = self.app
        try:
            if key == "contract":
                return (
                    app.contract_number_var.get(),
                    app.contract_day_var.get(),
                    app.contract_month_var.get(),
                )
            if key == "additional":
                return (
                    app.addl_number_var.get(),
                    app._limit_var.get(),
                    app.term_var.get(),
                )
            if key == "spec":
                items = tuple(
                    (row["name"].get(), row["qty"].get(), row["unit"].get(), row["price"].get())
                    for row in app._spec_rows
                )
                return (
                    app.spec_number_var.get(),
                    app.spec_day_var.get(),
                    app.spec_month_var.get(),
                    app.spec_delivery_term_var.get(),
                    app.spec_payment_terms_var.get(),
                    app.spec_delivery_basis_var.get(),
                    app.spec_prices_until_var.get(),
                    items,
                )
            if key == "counterparty":
                mode = "ip" if app._cp_type_var.get() == "ИП" else "ooo"
                data = app._cp_collect(mode)
                if data is None:
                    return None
                return tuple(sorted(data.items()))
        except Exception as exc:
            # Сбой чтения формы — повод предположить несохранённые данные,
            # а не «чистоту» страницы (аудит 2.0.1).
            app._log.warning(
                "Не удалось снять подпись формы «%s»: %s", key, type(exc).__name__,
            )
            return None
        return None

    def reset_form_baseline(self, key: str) -> None:
        app = self.app
        app._form_baseline[key] = self.form_signature(key)

    def reset_all_form_baselines(self) -> None:
        for key in self.FORM_PAGES:
            self.reset_form_baseline(key)

    def is_form_dirty(self, key: str) -> bool:
        app = self.app
        if key not in app._form_baseline:
            return False
        return self.form_signature(key) != app._form_baseline[key]

    def dirty_pages(self) -> list[str]:
        return [
            key for key in self.FORM_PAGES
            if self.is_form_dirty(key)
        ]

    def confirm_leave_page(self) -> bool:
        app = self.app
        key = app.current_page
        if key not in app.pages or not self.is_form_dirty(key):
            return True
        return ui.confirm(
            app.root, "Несохранённые изменения",
            f"На странице «{self.form_label(key)}» есть несохранённые данные.\n\n"
            "Перейти и потерять изменения?",
            ok_text="Перейти", cancel_text="Остаться",
        )

    def save_current_page(self) -> None:
        app = self.app
        commands = {
            "contract": app.save_contract,
            "additional": app.save_additional,
            "spec": app.save_spec,
            "counterparty": app._cp_save,
        }
        command = commands.get(app.current_page)
        if command is not None:
            command()
