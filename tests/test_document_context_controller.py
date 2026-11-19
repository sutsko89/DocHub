"""Юнит-тесты DocumentContextController — девятый контроллер декомпозиции.

Аудит 2.0.1 (P2 «контроллеры»): единый «текущий контрагент» (п. 4.2) и
защита несохранённых данных (п. 4.4) переехали в
document_context_controller.py и проверяются без Tk — контроллер получает
фейковое приложение (duck typing) и настоящее хранилище DataStore, а
smoke-тест продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами (из аудита 2.0.1):
  * совпадающие названия контрагентов получают «· ID …» в списках (п. 4.1);
  * выбор, сделанный на одной странице, переносится на остальные (п. 4.2);
    устаревший выбор (контрагент удалён) честно очищается;
  * подпись формы: выбор контрагента/договора НЕ входит в подпись; при
    сбое чтения подпись деградирует до None — страница «грязная», а не
    «чистая» (аудит P1-3/P2-9);
  * уход со страницы с изменениями — только через подтверждение;
  * Ctrl+S сохраняет текущую страницу через фасады приложения;
  * фасады ContractApp с прежними именами делегируют контроллеру.
"""
from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest

import document_context_controller as dcc
import main
import storage
from document_builder import F
from render_helpers import OOO_FIELDS

DocumentContextController = dcc.DocumentContextController


# --------------------------------------------------------------------------- #
#  Фейки
# --------------------------------------------------------------------------- #
class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeCombo:
    def __init__(self):
        self.values: list | None = None

    def set_values(self, values):
        self.values = list(values)


class FakeApp:
    """Duck-typed приложение: только то, что трогает контроллер контекста."""

    def __init__(self, store):
        self.root = object()
        self.store = store
        self._log = logging.getLogger("test-document-context-controller")
        self.calls: list = []
        self.confirms: list = []
        self.confirm_result = True

        # Контекст (п. 4.2, 4.4) остаётся на приложении.
        self._current_party_id = None
        self._current_party_vars = {"contract": FakeVar("")}
        self._syncing_party = False
        self._form_baseline: dict = {}

        # Страницы: состояние страниц документов (шаги 5–7 декомпозиции).
        self.contract_party_combo = FakeCombo()
        self.addl_party_combo = FakeCombo()
        self.spec_party_combo = FakeCombo()
        self.contract_party_var = FakeVar("")
        self.addl_party_var = FakeVar("")
        self.spec_party_var = FakeVar("")
        self.contract_number_var = FakeVar("")
        self.contract_day_var = FakeVar("")
        self.contract_month_var = FakeVar("")
        self.addl_number_var = FakeVar("")
        self._limit_var = FakeVar("")
        self.term_var = FakeVar("")
        self.addl_contract_var = FakeVar("")
        self.addl_contract_info_var = FakeVar("")
        self.spec_contract_var = FakeVar("")
        self.spec_contract_info_var = FakeVar("")
        self.spec_number_var = FakeVar("")
        self.spec_day_var = FakeVar("")
        self.spec_month_var = FakeVar("")
        self.spec_delivery_term_var = FakeVar("")
        self.spec_payment_terms_var = FakeVar("")
        self.spec_delivery_basis_var = FakeVar("")
        self.spec_prices_until_var = FakeVar("")
        self._spec_rows: list = []

        # Карточка контрагента (шаг 6): подпись снимается через фасад.
        self._cp_type_var = FakeVar("Юр. лицо")

        self.current_page = "contract"
        self.pages = {"contract": object(), "additional": object(),
                      "spec": object(), "counterparty": object()}

    # -- фасады приложения, разрешаемые контроллером на каждый вызов ------- #
    def _cp_collect(self, mode):
        self.calls.append(("cp_collect", mode))
        return {"name": "ООО «Ромашка»", "type": "ooo"}

    def save_contract(self):
        self.calls.append("save_contract")

    def save_additional(self):
        self.calls.append("save_additional")

    def save_spec(self):
        self.calls.append("save_spec")

    def _cp_save(self):
        self.calls.append("cp_save")


@pytest.fixture()
def app(store):
    a = FakeApp(store)
    a._doc_context_flow = DocumentContextController(a)
    return a


@pytest.fixture()
def flow(app):
    return app._doc_context_flow


@pytest.fixture()
def confirm_stub(monkeypatch):
    state = {"results": [], "calls": []}

    def fake_confirm(root, *args, **kwargs):
        state["calls"].append((args, kwargs))
        if state["results"]:
            return state["results"].pop(0)
        return True

    monkeypatch.setattr(dcc.ui, "confirm", fake_confirm)
    return state


def seed_party(store, fields=None):
    data = {k: v for k, v in OOO_FIELDS.items() if k not in (F.NUMBER, F.DAY, F.MONTH)}
    if fields:
        data.update(fields)
    return store.add_counterparty("ooo", data)


# --------------------------------------------------------------------------- #
#  Списки контрагентов (п. 4.1, 4.2)
# --------------------------------------------------------------------------- #
def test_party_duplicate_bases_detects_duplicates(app, flow, store, ooo_fields):
    seed_party(store)
    seed_party(store)  # та же карточка ещё раз — то же отображаемое имя

    duplicates = flow.party_duplicate_bases()
    assert len(duplicates) == 1
    base, = duplicates
    assert "ООО" in base


def test_party_display_appends_id_for_duplicates(app, flow, store):
    rec1 = seed_party(store)
    rec2 = seed_party(store)
    duplicates = flow.party_duplicate_bases()

    assert flow.party_display(rec1, duplicates).endswith(f"ID {rec1['id']}")
    assert flow.party_display(rec2, duplicates).endswith(f"ID {rec2['id']}")
    assert "·" in flow.party_display(rec1, duplicates)


def test_party_display_without_duplicates_plain(app, flow, store):
    rec = seed_party(store)
    assert flow.party_display(rec, set()) == storage.display_with_inn(rec)


def test_party_display_values_empty_registry(app, flow, store):
    assert flow.party_display_values() == ["— контрагентов нет —"]


def test_party_display_values_with_parties(app, flow, store):
    rec = seed_party(store)
    assert flow.party_display_values() == [storage.display_with_inn(rec)]


def test_refresh_party_combos_clears_stale_selection(app, flow, store):
    rec = seed_party(store)
    app.contract_party_var.set("удалённый контрагент")
    app.addl_party_var.set(storage.display_with_inn(rec))
    app._current_party_id = "missing-id"

    flow.refresh_party_combo_values()

    assert app.contract_party_combo.values == [storage.display_with_inn(rec)]
    assert app.contract_party_var.get() == ""          # устаревший выбор очищен
    assert app.addl_party_var.get() == storage.display_with_inn(rec)  # действующий остался
    assert app._current_party_id is None               # ссылка на удалённого сброшена


def test_refresh_party_combos_keeps_placeholder(app, flow, store):
    seed_party(store)
    app.contract_party_var.set("— контрагентов нет —")
    flow.refresh_party_combo_values()
    # Placeholder сохраняется: это не устаревшее имя, а штатная подсказка.
    assert app.contract_party_var.get() == "— контрагентов нет —"


def test_party_from_var_matches_display_with_duplicates(app, flow, store):
    rec = seed_party(store)
    seed_party(store)
    display = flow.party_display(rec, flow.party_duplicate_bases())
    app.contract_party_var.set(display)
    assert flow.party_from_var(app.contract_party_var)["id"] == rec["id"]


def test_party_from_var_legacy_value_without_id(app, flow, store):
    rec = seed_party(store)
    # Значение установлено программно без «· ID …» (совместимость).
    app.contract_party_var.set(storage.display_with_inn(rec))
    assert flow.party_from_var(app.contract_party_var)["id"] == rec["id"]


def test_party_from_var_empty_and_missing(app, flow, store):
    seed_party(store)
    app.contract_party_var.set("")
    assert flow.party_from_var(app.contract_party_var) is None
    app.contract_party_var.set("нет такого")
    assert flow.party_from_var(app.contract_party_var) is None


# --------------------------------------------------------------------------- #
#  Единый «текущий контрагент» (п. 4.2)
# --------------------------------------------------------------------------- #
def test_current_party_label_states(app, flow, store):
    assert flow.current_party_label() == "Документ для: контрагент не выбран"
    app._current_party_id = "missing"
    assert "не найден" in flow.current_party_label()
    rec = seed_party(store)
    app._current_party_id = rec["id"]
    assert flow.current_party_label() == f"Документ для: {storage.display_with_inn(rec)}"


def test_update_current_party_labels_sets_all_vars(app, flow):
    app._current_party_vars = {"contract": FakeVar(""), "additional": FakeVar("")}
    flow.update_current_party_labels()
    assert app._current_party_vars["contract"].value == "Документ для: контрагент не выбран"
    assert app._current_party_vars["additional"].value == "Документ для: контрагент не выбран"


def test_set_current_party_syncs_other_pages(app, flow, store):
    rec = seed_party(store)
    app.addl_party_var.set("старое имя")
    app.spec_party_var.set("")

    flow.set_current_party(rec, source="contract")

    display = storage.display_with_inn(rec)
    assert app._current_party_id == rec["id"]
    assert app.addl_party_var.get() == display          # перенесён на доп. соглашение
    assert app.spec_party_var.get() == display          # и на спецификацию
    assert app._syncing_party is False                  # guard снят после синхронизации


def test_set_current_party_ignores_source_page(app, flow, store):
    rec = seed_party(store)
    display = storage.display_with_inn(rec)
    app.contract_party_var.set("выбор пользователя")
    flow.set_current_party(rec, source="contract")
    # Страница-источник не перетирается: там значение ставит её собственный список.
    assert app.contract_party_var.get() == "выбор пользователя"
    assert app.addl_party_var.get() == display
    assert app.spec_party_var.get() == display


def test_set_current_party_ignores_non_dict(app, flow):
    flow.set_current_party(None)
    assert app._current_party_id is None


def test_on_party_selected_syncs_and_skips_during_sync(app, flow, store):
    rec = seed_party(store)
    app.addl_party_var.set(storage.display_with_inn(rec))

    flow.on_party_selected("additional")

    assert app._current_party_id == rec["id"]

    # Guard: во время программной синхронизации событие игнорируется.
    app._syncing_party = True
    app.contract_party_var.set(storage.display_with_inn(rec))
    flow.on_party_selected("contract")
    assert app._current_party_id == rec["id"]  # не изменился


def test_on_party_selected_empty_clears_current(app, flow, store):
    rec = seed_party(store)
    app._current_party_id = rec["id"]
    app.spec_party_var.set("")
    flow.on_party_selected("spec")
    assert app._current_party_id is None


def test_update_contract_info_additional(app, flow):
    app.addl_contract_var.set("№ 5 от 10 января")
    flow.update_contract_info("additional")
    assert app.addl_contract_info_var.get() == "Договор-основание: № 5 от 10 января"

    app.addl_contract_var.set("— нет сохранённых договоров —")
    flow.update_contract_info("additional")
    assert "не выбран" in app.addl_contract_info_var.get()


def test_update_contract_info_spec(app, flow):
    app.spec_contract_var.set("")
    flow.update_contract_info("spec")
    assert "не выбран" in app.spec_contract_info_var.get()


def test_contract_display_with_and_without_date(app, flow):
    assert flow.contract_display({"number": "5", "date": "10 января"}) == "№ 5 от 10 января"
    assert flow.contract_display({"number": "5", "date": ""}) == "№ 5"
    assert flow.contract_display({"number": "", "date": ""}) == "№ б/н"


def test_contract_from_var(app, flow, store, ooo_fields):
    rec = seed_party(store)
    contract = store.add_document(
        storage.KIND_CONTRACT, rec["id"], ooo_fields[F.NUMBER],
        f"{ooo_fields[F.DAY]} {ooo_fields[F.MONTH]}", dict(ooo_fields),
    )
    app.addl_contract_var.set(f"№ {ooo_fields[F.NUMBER]} от {ooo_fields[F.DAY]} {ooo_fields[F.MONTH]}")
    assert flow.contract_from_var(app.addl_contract_var, rec["id"])["id"] == contract["id"]

    app.addl_contract_var.set("")
    assert flow.contract_from_var(app.addl_contract_var, rec["id"]) is None
    app.addl_contract_var.set("нет такого договора")
    assert flow.contract_from_var(app.addl_contract_var, rec["id"]) is None


# --------------------------------------------------------------------------- #
#  Подписи форм и защита несохранённых данных (п. 4.4)
# --------------------------------------------------------------------------- #
def test_form_label_static_mapping():
    assert DocumentContextController.form_label("contract") == "Договор поставки"
    assert DocumentContextController.form_label("counterparty") == "Карточка контрагента"
    assert DocumentContextController.form_label("unknown") == "unknown"
    # Фасад на ContractApp — тоже staticmethod (вызов без экземпляра).
    assert main.ContractApp._form_label("spec") == "Спецификация"


def test_form_signature_contract(app, flow):
    app.contract_number_var.set("5")
    app.contract_day_var.set("12")
    app.contract_month_var.set("января")
    assert flow.form_signature("contract") == ("5", "12", "января")


def test_form_signature_additional(app, flow):
    app.addl_number_var.set("1")
    app._limit_var.set("1 500 000")
    app.term_var.set("45")
    assert flow.form_signature("additional") == ("1", "1 500 000", "45")


def test_form_signature_spec_includes_items(app, flow):
    app.spec_number_var.set("3")
    app._spec_rows = [{"name": FakeVar("Болт"), "qty": FakeVar("5"),
                       "unit": FakeVar("шт."), "price": FakeVar("200")}]
    signature = flow.form_signature("spec")
    assert signature[0] == "3"
    assert signature[-1] == (("Болт", "5", "шт.", "200"),)


def test_form_signature_counterparty_sorted(app, flow):
    data = flow.form_signature("counterparty")
    assert data == tuple(sorted({"name": "ООО «Ромашка»", "type": "ooo"}.items()))
    assert app.calls == [("cp_collect", "ooo")]


def test_form_signature_counterparty_ip_mode(app, flow):
    app._cp_type_var.set("ИП")
    flow.form_signature("counterparty")
    assert app.calls[-1] == ("cp_collect", "ip")


def test_form_signature_unknown_key(app, flow):
    assert flow.form_signature("registry") is None


def test_form_signature_degrades_to_none_on_failure(app, flow):
    # Сбой чтения виджета (аудит 2.0.1): None, а не пустой кортеж —
    # страница считается «грязной», несохранённые данные защищены.
    class BrokenVar:
        def get(self):
            raise RuntimeError("widget destroyed")

    app.contract_number_var = BrokenVar()
    assert flow.form_signature("contract") is None


def test_baseline_and_dirty_tracking(app, flow):
    app.contract_number_var.set("5")
    flow.reset_form_baseline("contract")
    assert not flow.is_form_dirty("contract")

    app.contract_number_var.set("6")
    assert flow.is_form_dirty("contract")
    assert flow.dirty_pages() == ["contract"]

    flow.reset_all_form_baselines()
    assert flow.dirty_pages() == []


def test_is_form_dirty_unknown_page_clean(app, flow):
    assert flow.is_form_dirty("registry") is False


def test_confirm_leave_page_clean_page_no_dialog(app, flow, confirm_stub):
    app.current_page = "registry"  # нет в pages
    assert flow.confirm_leave_page() is True
    app.current_page = "contract"
    flow.reset_form_baseline("contract")
    assert flow.confirm_leave_page() is True
    assert confirm_stub["calls"] == []  # чистая страница — без диалога


def test_confirm_leave_page_dirty_asks_confirmation(app, flow, confirm_stub):
    app.current_page = "contract"
    app.contract_number_var.set("5")
    flow.reset_form_baseline("contract")
    app.contract_number_var.set("6")

    assert flow.confirm_leave_page() is True  # подтверждение по умолчанию
    (args, kwargs) = confirm_stub["calls"][0]
    assert "Несохранённые изменения" in args[0]
    assert "Договор поставки" in args[1]
    assert kwargs.get("ok_text") == "Перейти"

    confirm_stub["results"].append(False)
    assert flow.confirm_leave_page() is False  # пользователь остался


def test_save_current_page_dispatch(app, flow):
    app.current_page = "contract"
    flow.save_current_page()
    app.current_page = "additional"
    flow.save_current_page()
    app.current_page = "spec"
    flow.save_current_page()
    app.current_page = "counterparty"
    flow.save_current_page()
    assert app.calls == ["save_contract", "save_additional", "save_spec", "cp_save"]


def test_save_current_page_unknown_page_noop(app, flow):
    app.current_page = "registry"
    flow.save_current_page()
    assert app.calls == []


# --------------------------------------------------------------------------- #
#  Фасады ContractApp делегируют контроллеру
# --------------------------------------------------------------------------- #
class RecordingFlow:
    def __init__(self):
        self.calls: list = []

    def _record(self, name, *args, **kwargs):
        self.calls.append((name, args, kwargs))
        return "recorded"

    def party_duplicate_bases(self):
        return self._record("party_duplicate_bases")

    def party_display(self, *args, **kwargs):
        return self._record("party_display", *args, **kwargs)

    def party_display_values(self):
        return self._record("party_display_values")

    def refresh_party_combo_values(self):
        return self._record("refresh_party_combo_values")

    def party_from_var(self, *args):
        return self._record("party_from_var", *args)

    def current_party_label(self):
        return self._record("current_party_label")

    def update_current_party_labels(self):
        return self._record("update_current_party_labels")

    def set_current_party(self, *args, **kwargs):
        return self._record("set_current_party", *args, **kwargs)

    def on_party_selected(self, *args):
        return self._record("on_party_selected", *args)

    def update_contract_info(self, *args):
        return self._record("update_contract_info", *args)

    def contract_display(self, *args):
        return self._record("contract_display", *args)

    def contract_from_var(self, *args):
        return self._record("contract_from_var", *args)

    def form_signature(self, *args):
        return self._record("form_signature", *args)

    def reset_form_baseline(self, *args):
        return self._record("reset_form_baseline", *args)

    def reset_all_form_baselines(self):
        return self._record("reset_all_form_baselines")

    def is_form_dirty(self, *args):
        return self._record("is_form_dirty", *args)

    def dirty_pages(self):
        return self._record("dirty_pages")

    def confirm_leave_page(self):
        return self._record("confirm_leave_page")

    def save_current_page(self):
        return self._record("save_current_page")


def make_facade_app():
    flow = RecordingFlow()
    return SimpleNamespace(_doc_context_flow=flow), flow


def test_facade_party_context_delegates():
    app, flow = make_facade_app()

    main.ContractApp._party_duplicate_bases(app)
    main.ContractApp._party_display(app, {"id": 1}, set())
    main.ContractApp._party_display_values(app)
    main.ContractApp._refresh_party_combo_values(app)
    main.ContractApp._party_from_var(app, "var")
    main.ContractApp._current_party_label(app)
    main.ContractApp._update_current_party_labels(app)
    main.ContractApp._set_current_party(app, {"id": 2}, source="spec")
    main.ContractApp._on_party_selected(app, "contract")
    main.ContractApp._update_contract_info(app, "additional")
    main.ContractApp._contract_display(app, {"number": "5"})
    main.ContractApp._contract_from_var(app, "var", 7)

    assert [call[0] for call in flow.calls] == [
        "party_duplicate_bases", "party_display", "party_display_values",
        "refresh_party_combo_values", "party_from_var", "current_party_label",
        "update_current_party_labels", "set_current_party",
        "on_party_selected", "update_contract_info", "contract_display",
        "contract_from_var",
    ]
    assert flow.calls[7][2] == {"source": "spec"}
    assert flow.calls[8][1] == ("contract",)


def test_facade_form_state_delegates():
    app, flow = make_facade_app()

    main.ContractApp._form_signature(app, "contract")
    main.ContractApp._reset_form_baseline(app, "contract")
    main.ContractApp._reset_all_form_baselines(app)
    main.ContractApp._is_form_dirty(app, "contract")
    main.ContractApp._dirty_pages(app)
    main.ContractApp._confirm_leave_page(app)
    main.ContractApp._save_current_page(app)

    assert [call[0] for call in flow.calls] == [
        "form_signature", "reset_form_baseline", "reset_all_form_baselines",
        "is_form_dirty", "dirty_pages", "confirm_leave_page", "save_current_page",
    ]
    assert flow.calls[0][1] == ("contract",)


def test_facade_form_label_is_static():
    """Фасад _form_label — staticmethod, как исходный метод ContractApp."""
    assert main.ContractApp._form_label("contract") == "Договор поставки"
