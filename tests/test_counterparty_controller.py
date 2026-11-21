"""Юнит-тесты CounterpartyController — шестой контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P2 «контроллеры»): страница «Добавить контрагента» и окно
редактирования карточки переехали в counterparty_controller.py и проверяются
без Tk — контроллер получает фейковое приложение (duck typing) и настоящее
хранилище DataStore, а smoke-тест продолжает покрывать GUI-путь целиком
(сборка страницы, «Заполнить по ИНН», окно редактирования — только Tk).

Контракты, закреплённые тестами (из аудита 2.0.1):
  * сбор данных карточки: значения читаются и очищаются; сбой чтения любого
    поля возвращает None (не пустую строку) и пишется в журнал;
  * у ИП фактический адрес повторяет юридический;
  * сохранение при сбое сборки показывает ошибку и не пишет реестр;
  * ошибки валидации подсвечивают поля и не доходят до хранилища;
  * дубликат ИНН — выбор из трёх исходов (п. 4.1): «Открыть существующего»
    выбирает существующую карточку и уводит на договор, «Отмена» ничего не
    меняет, «Создать всё равно» создаёт дубль;
  * сбой записи реестра не оставляет «фантом» в памяти (п. 9.2);
  * успешное сохранение сбрасывает baseline карточки, обновляет списки и
    сразу переходит к договору с уже выбранным контрагентом;
  * соседние контроллеры вызываются напрямую через ``app._flow.*``
    (фасады ContractApp удалены, v2.12.0), а защита несохранённых данных
    (п. 4.4) видит карточку через тот же конвейер подписи.
"""
from __future__ import annotations

import logging
from types import SimpleNamespace

import pytest

import counterparty_controller as cc
import document_context_controller
import services
import storage
from document_builder import F
from render_helpers import IP_FIELDS, OOO_FIELDS

CounterpartyController = cc.CounterpartyController

# Поля документа (номер/дата) в карточку контрагента не входят.
DOCUMENT_KEYS = (F.NUMBER, F.DAY, F.MONTH)

# CHKSUM (этап 2): карточки проходят проверку контрольных цифр — save_card
# проверяет их для НОВЫХ/ИЗМЕНЁННЫХ реквизитов, поэтому тестовые ИНН/ОГРН
# должны быть корректны. Значения ОOO_FIELDS/IP_FIELDS оставлены без изменений:
# на них построены golden-эталоны печатных форм.
OOO_CARD = {
    k: str(v) for k, v in OOO_FIELDS.items() if k not in DOCUMENT_KEYS
}
OOO_CARD[F.INN] = "2540123459"       # контрольная цифра валидна
OOO_CARD[F.OGRN] = "1035001234564"   # контрольная цифра валидна
IP_CARD = {
    k: str(v) for k, v in IP_FIELDS.items() if k not in DOCUMENT_KEYS
}
IP_CARD[F.INN] = "254012345692"       # 12-значный ИНН, обе контрольные валидны
IP_CARD[F.OGRNIP] = "304253600012343"  # контрольная цифра валидна


def ooo_card() -> dict:
    return dict(OOO_CARD)


def ip_card() -> dict:
    return dict(IP_CARD)


# --------------------------------------------------------------------------- #
#  Фейки: приложение и виджеты
# --------------------------------------------------------------------------- #
class FakeVar:
    """Минимальный аналог tk.StringVar (без Tk: подменяет cc.tk.StringVar)."""

    def __init__(self, value=""):
        self._value = value

    def get(self) -> str:
        return self._value

    def set(self, value) -> None:
        self._value = value


class FakeEntry:
    """Минимальный аналог поля ввода с подсветкой ошибки."""

    def __init__(self, value=""):
        self._value = value
        self.error_marks: list[bool] = []

    def get(self) -> str:
        return self._value

    def set_error(self, on: bool) -> None:
        self.error_marks.append(on)


class BrokenEntry:
    """Поле, чтение которого падает (уничтоженный виджет, сбой Tk)."""

    def get(self):
        raise RuntimeError("widget destroyed")


class FakeFrame:
    """Минимальный аналог кадра «Юр. лицо / ИП»."""

    def __init__(self, name: str):
        self.name = name
        self.raised = 0

    def tkraise(self) -> None:
        self.raised += 1


class FakeApp:
    """Duck-typed приложение: только то, что трогает CounterpartyController."""

    def __init__(self, store: storage.DataStore):
        self.root = object()  # первый аргумент ui.alert — заглушен
        self.store = store
        self._log = logging.getLogger("test-counterparty-controller")
        self.calls: list = []
        self.notifications: list = []
        self.choice_calls: list = []
        self.choice_result: str | None = None
        self.current_party = None

        # Состояние страницы остаётся на приложении (шаг 6 декомпозиции).
        self._cp_entries: dict[str, dict] = {"ooo": {}, "ip": {}}
        self._cp_error_vars: dict[str, dict] = {"ooo": {}, "ip": {}}
        self._cp_type_var = FakeVar("Юр. лицо")
        self._cp_frames = {"ooo": FakeFrame("ooo"), "ip": FakeFrame("ip")}
        # Карточка живёт в контроллере (шаг 6 декомпозиции), как и в
        # ContractApp.__init__.
        self._counterparty_flow = CounterpartyController(self)
        # Соседние контроллеры вызываются напрямую (v2.12.0, фасады
        # ContractApp удалены): записывающие фейки проверяют и порядок
        # вызовов при сохранении карточки.
        self._doc_context_flow = SimpleNamespace(
            reset_form_baseline=lambda page: self.calls.append(("baseline", page)),
            set_current_party=lambda rec, source=None: (
                self.calls.append(("set_party", rec)),
                setattr(self, "current_party", rec))[-1],
            refresh_party_combo_values=lambda: self.calls.append("combo"),
        )
        self._registry_flow = SimpleNamespace(
            refresh=lambda: self.calls.append("registry"))
        self._doc_page_flow = SimpleNamespace(
            preview_contract=lambda *a: None,
            preview_additional=lambda *a: None,
        )

    # -- хуки, оставшиеся на приложении ------------------------------------- #
    def _choice_dialog(self, title, message, options):
        self.choice_calls.append((title, message, options))
        return self.choice_result

    def _select_page(self, key):
        self.calls.append(("page", key))

    def _update_status(self):
        self.calls.append("status")

    def _notify(self, message, path=None):
        self.notifications.append(message)


def seed_entries(app: FakeApp, mode: str, card: dict) -> None:
    """Заполняет состояние страницы фейковыми полями по словарю реквизитов."""
    app._cp_entries[mode] = {key: FakeEntry(value) for key, value in card.items()}


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def ui_stub(monkeypatch):
    """Перехват ui.alert: запись текстов (выбор исходов — через FakeApp)."""
    state = {"alerts": []}

    def fake_alert(root, *args, **kwargs):
        state["alerts"].append((args, kwargs))

    monkeypatch.setattr(cc.ui, "alert", fake_alert)
    return state


@pytest.fixture()
def app(store):
    return FakeApp(store)


@pytest.fixture()
def flow(app):
    return app._counterparty_flow


@pytest.fixture()
def tk_vars(monkeypatch):
    """tk.StringVar -> FakeVar: пути ошибок создают переменные без Tk."""
    monkeypatch.setattr(cc.tk, "StringVar", FakeVar)


# --------------------------------------------------------------------------- #
#  collect: сбор данных карточки
# --------------------------------------------------------------------------- #
def test_collect_reads_values_and_strips(app, flow):
    seed_entries(app, "ooo", {"ooo": ooo_card()}["ooo"])
    app._cp_entries["ooo"][F.INN] = FakeEntry(f" {OOO_FIELDS[F.INN]} ")
    data = flow.collect("ooo")
    assert data[F.INN] == OOO_FIELDS[F.INN]
    assert data[F.SHORT_NAME] == OOO_FIELDS[F.SHORT_NAME]


def test_collect_failure_returns_none_and_logs(app, flow, caplog):
    seed_entries(app, "ooo", ooo_card())
    app._cp_entries["ooo"][F.KPP] = BrokenEntry()
    with caplog.at_level(logging.WARNING, logger="test-counterparty-controller"):
        assert flow.collect("ooo") is None
    assert "Не удалось прочитать поле карточки" in caplog.text
    assert F.KPP in caplog.text


def test_collect_ip_copies_legal_address_to_fact(app, flow):
    seed_entries(app, "ip", ip_card())
    data = flow.collect("ip")
    assert data[F.FACT_ADDRESS] == IP_FIELDS[F.LEGAL_ADDRESS]


# --------------------------------------------------------------------------- #
#  error_var_for / on_type_change
# --------------------------------------------------------------------------- #
def test_error_var_for_creates_and_reuses_var(app, flow, tk_vars):
    get = flow.error_var_for("ooo")
    var = get(F.INN)
    assert get(F.INN) is var, "переменная ошибки переиспользуется"
    assert get(F.KPP) is not var, "у каждого поля своя переменная"


def test_on_type_change_raises_mode_frame(app, flow):
    flow.on_type_change("ИП")
    assert app._cp_frames["ip"].raised == 1
    assert app._cp_frames["ooo"].raised == 0
    flow.on_type_change("Юр. лицо")
    assert app._cp_frames["ooo"].raised == 1


# --------------------------------------------------------------------------- #
#  save_card: сбой сборки и ошибки валидации
# --------------------------------------------------------------------------- #
def test_save_card_collect_failure_alerts_and_saves_nothing(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    app._cp_entries["ooo"][F.INN] = BrokenEntry()
    flow.save_card()
    assert ui_stub["alerts"], "пользователю показано сообщение о сбое"
    assert "Не удалось прочитать поля формы" in ui_stub["alerts"][0][0][0]
    assert app.store.counterparties == [], "до хранилища дело не дошло"


def test_save_card_validation_errors_highlight_and_write_nothing(
        app, flow, ui_stub, tk_vars):
    seed_entries(app, "ooo", ooo_card())
    app._cp_entries["ooo"][F.INN] = FakeEntry("")  # пустой ИНН — ошибка
    flow.save_card()
    assert ui_stub["alerts"], "ошибки валидации показаны"
    assert ui_stub["alerts"][0][0][0] == "Ошибки в данных"
    var = app._cp_error_vars["ooo"][F.INN]
    assert var.get() == "Заполните поле"
    # Подсветка: сначала сброс ошибки у всех полей, потом установка у сбойных.
    assert app._cp_entries["ooo"][F.INN].error_marks == [False, True]
    assert app.store.counterparties == [], "ошибочные данные не пишутся"


def test_save_card_valid_data_shows_no_errors(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    flow.save_card()  # до диалога выбора при дубликате ИНН не доходит
    assert not ui_stub["alerts"]
    assert not app.choice_calls


# --------------------------------------------------------------------------- #
#  save_card: дубликат ИНН — три исхода (п. 4.1)
# --------------------------------------------------------------------------- #
def test_save_card_duplicate_inn_open_existing(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    existing = app.store.add_counterparty(
        "ooo", ooo_card(), save=False)
    app.choice_result = "open"
    flow.save_card()

    assert len(app.choice_calls) == 1
    title, message, options = app.choice_calls[0]
    assert title == "Контрагент с таким ИНН уже есть"
    assert OOO_CARD[F.INN] in message
    assert [opt[0] for opt in options] == ["open", "create", "cancel"]

    assert app.current_party == existing, "выбрана существующая карточка"
    assert app.calls == [
        ("baseline", "counterparty"), ("set_party", existing),
        ("page", "contract"),
    ]
    assert len(app.store.counterparties) == 1, "дубль не создан"


def test_save_card_duplicate_inn_cancel_changes_nothing(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    app.store.add_counterparty("ooo", ooo_card(), save=False)
    app.choice_result = "cancel"
    flow.save_card()

    assert len(app.choice_calls) == 1
    assert app.calls == [], "ни карточка, ни страница не изменились"
    assert len(app.store.counterparties) == 1


def test_save_card_duplicate_inn_create_anyway(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    app.store.add_counterparty("ooo", ooo_card(), save=False)
    app.choice_result = "create"
    flow.save_card()

    assert len(app.store.counterparties) == 2, "дубль создан по явному выбору"
    names = [call[0] if isinstance(call, tuple) else call for call in app.calls]
    assert names == ["baseline", "combo", "status", "registry",
                     "set_party", "page"]


# --------------------------------------------------------------------------- #
#  save_card: запись в реестр
# --------------------------------------------------------------------------- #
def test_save_card_store_failure_removes_phantom(app, flow, ui_stub,
                                                 monkeypatch):
    seed_entries(app, "ooo", ooo_card())

    def broken_add(mode, data):
        rec = {"id": "cp-broken", "type": mode, "fields": dict(data)}
        app.store.counterparties.append(rec)  # «фантом» в памяти
        app.store.last_error = "Не удалось записать файл реестра"
        return rec

    monkeypatch.setattr(app.store, "add_counterparty", broken_add)
    flow.save_card()
    assert app.store.counterparties == [], "«фантом» удалён из памяти"
    assert ui_stub["alerts"], "пользователю показана ошибка"
    assert ui_stub["alerts"][0][0][0] == "Не удалось сохранить контрагента"
    assert app.calls == [], "хуки успеха не выполнялись"


def test_save_card_success_flow(app, flow, ui_stub):
    seed_entries(app, "ooo", ooo_card())
    flow.save_card()
    assert len(app.store.counterparties) == 1
    rec = app.store.counterparties[0]
    assert rec["type"] == "ooo"
    assert rec["fields"][F.INN] == OOO_CARD[F.INN]
    assert app.calls == [
        ("baseline", "counterparty"), "combo", "status", "registry",
        ("set_party", rec), ("page", "contract"),
    ]
    assert app.notifications, "тост показан"
    assert "ООО «Ромашка»" in app.notifications[0]
    assert "добавлен в реестр программы" in app.notifications[0]


def test_save_card_ip_mode_writes_fact_address(app, flow):
    app._cp_type_var.set("ИП")
    seed_entries(app, "ip", ip_card())
    flow.save_card()
    rec = app.store.counterparties[0]
    assert rec["type"] == "ip"
    assert rec["fields"][F.FACT_ADDRESS] == IP_FIELDS[F.LEGAL_ADDRESS]


# --------------------------------------------------------------------------- #
#  Подпись карточки — DocumentContextController (п. 4.4)
# --------------------------------------------------------------------------- #
def test_form_signature_reads_card_through_controller(app, flow):
    seed_entries(app, "ooo", ooo_card())
    signature = document_context_controller.DocumentContextController(
        app).form_signature("counterparty")
    expected = tuple(sorted(flow.collect("ooo").items()))
    assert signature == expected


def test_form_signature_degrades_to_none_on_widget_failure(app, flow):
    seed_entries(app, "ooo", ooo_card())
    app._cp_entries["ooo"][F.INN] = BrokenEntry()
    assert document_context_controller.DocumentContextController(
        app).form_signature("counterparty") is None


# --------------------------------------------------------------------------- #
#  DEF-1: защита карточки при редактировании (D4 + U8-MVP)
# --------------------------------------------------------------------------- #
class FakeWindow:
    """Минимальный аналог Toplevel: только destroy() (без Tk)."""

    def __init__(self):
        self.destroyed = 0

    def destroy(self) -> None:
        self.destroyed += 1


def test_read_entry_values_returns_failed_keys():
    entries = {"inn": FakeEntry(" 123 "), "email": BrokenEntry()}
    data, failed = cc.read_entry_values(entries)
    assert data == {"inn": "123"}
    assert failed == ["email"], "сбойное поле не попадает в data молча"


def test_read_entry_values_logs_failure(caplog):
    with caplog.at_level(logging.WARNING, logger="test-counterparty-controller"):
        cc.read_entry_values({"email": BrokenEntry()}, log=logging.getLogger(
            "test-counterparty-controller"))
    assert "email" in caplog.text and "RuntimeError" in caplog.text


def test_form_signature_marks_broken_field():
    signature = cc.form_signature({"a": FakeEntry("x"), "b": BrokenEntry()})
    assert signature["b"] == cc.READ_ERROR_MARKER
    assert signature["b"] != "", "сбой чтения не выглядит как пустая строка"


def test_split_read_failures_separates_required_and_optional():
    required = services.REQUIRED_PARTY_COMMON + services.REQUIRED_OOO
    blocking, optional = cc.split_read_failures([F.INN, F.EMAIL, F.PHONE], required)
    assert blocking == [F.INN], "обязательное поле блокирует сохранение"
    assert optional == [F.EMAIL, F.PHONE]


def test_merge_saved_values_fills_only_failed_optional():
    saved = {F.EMAIL: "old@test.ru", F.PHONE: "+7 4212 00-00-00"}
    data = cc.merge_saved_values({F.INN: "2540123456"}, [F.EMAIL], saved)
    assert data[F.EMAIL] == "old@test.ru", "прежнее значение подставлено"
    assert F.PHONE not in data, "несбойные поля не трогаются"


def test_close_guard_closes_without_changes():
    window = FakeWindow()
    entries = {"a": FakeEntry("x")}
    guard = cc.EditCloseGuard(window, entries, cc.form_signature(entries),
                              lambda: pytest.fail("без изменений подтверждения нет"))
    guard.request_close()
    assert window.destroyed == 1


def test_close_guard_asks_confirm_with_changes_and_closes():
    window = FakeWindow()
    entries = {"a": FakeEntry("x")}
    baseline = cc.form_signature(entries)
    asked = []
    guard = cc.EditCloseGuard(
        window, entries, baseline, lambda: asked.append(1) or True)
    entries["a"]._value = "y"  # пользователь изменил поле
    guard.request_close()
    assert asked, "при изменениях показано подтверждение"
    assert window.destroyed == 1, "согласие закрывает окно"


def test_close_guard_stays_when_discard_declined():
    window = FakeWindow()
    entries = {"a": FakeEntry("x")}
    baseline = cc.form_signature(entries)
    guard = cc.EditCloseGuard(window, entries, baseline, lambda: False)
    entries["a"]._value = "y"
    guard.request_close()
    assert window.destroyed == 0, "отказ оставляет окно открытым"


def test_close_guard_treats_read_failure_as_change():
    window = FakeWindow()
    entries = {"a": FakeEntry("x"), "b": FakeEntry("y")}
    baseline = cc.form_signature(entries)
    asked = []
    guard = cc.EditCloseGuard(
        window, entries, baseline, lambda: asked.append(1) or False)
    entries["b"] = BrokenEntry()  # виджет сломался посреди правки
    guard.request_close()
    assert asked, "сбой чтения поля не позволяет закрыть окно молча"
    assert window.destroyed == 0
