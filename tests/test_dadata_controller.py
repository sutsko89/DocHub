"""Юнит-тесты DadataController — второй контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P1 «старт покрытия main.py», P2 «контроллеры страниц»): поток
«Заполнить по ИНН» переехал в dadata_controller.py и проверяется без Tk.

Контракты, закреплённые тестами (из аудита 2.0.1):
  * guard повторных кликов: при активном запросе новый воркер не запускается;
  * ИНН валидируется по режиму карточки: 10 цифр для ООО, 12 для ИП
    (именно этот контракт сломал поиск по ИНН у ИП в 2.0.0);
  * без API-ключа — понятная ошибка, запрос не запускается;
  * результат приходит через очередь и опрашивается из главного потока;
  * подстановка реквизитов: ООО — полный набор, ИП — ОГРНИП и ФИО без КПП;
  * фасад ContractApp._cp_busy проецирует состояние контроллера.
"""
from __future__ import annotations

import logging
import queue
import tkinter as tk

import dadata_client
import dadata_controller as dc
import main
from document_builder import F

DadataController = dc.DadataController


# --------------------------------------------------------------------------- #
#  Фейки
# --------------------------------------------------------------------------- #
class FakeRoot:
    def __init__(self):
        self.after_calls: list = []
        self.cancelled: list = []

    def after(self, ms, fn):
        self.after_calls.append((ms, fn))
        return f"after-{len(self.after_calls)}"

    def after_cancel(self, token):
        self.cancelled.append(token)


class FakeVar:
    def __init__(self, value=""):
        self.value = value
        self.history: list = []

    def get(self):
        return self.value

    def set(self, value):
        self.value = value
        self.history.append(value)


class FakeEntry:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def delete(self, *_args):
        self.value = ""

    def insert(self, _index, value):
        self.value = value


class FakeApp:
    """Duck-typed приложение: только то, что трогает DadataController."""

    def __init__(self):
        self.root = FakeRoot()
        self._cp_type_var = FakeVar("Юр. лицо")
        self._cp_status_var = FakeVar("")
        self._cp_entries = {"ooo": {}, "ip": {}}

    def _cp_fetch_worker(self, inn, api_key, mode, result_queue):
        self._cp_entries  # заглушка: фасад для подмены в smoke

    def _cp_result_tick(self):
        pass

    def _cp_busy_tick(self):
        pass


def make_flow() -> tuple[DadataController, FakeApp]:
    app = FakeApp()
    return DadataController(app), app


class RecordedThread:
    def __init__(self, target, args, daemon):
        self.target = target
        self.args = args
        self.daemon = daemon
        self.started = False

    def start(self):
        self.started = True


class RecordingFactory:
    def __init__(self):
        self.created: list = []

    def __call__(self, target=None, args=(), daemon=None, **_kw):
        thread = RecordedThread(target, args, daemon)
        self.created.append(thread)
        return thread


def make_entries(inn: str = "7701234567") -> dict:
    return {
        F.INN: FakeEntry(inn),
        F.KPP: FakeEntry(),
        F.OGRN: FakeEntry(),
        F.OGRNIP: FakeEntry(),
        F.IP_NAME: FakeEntry(),
        F.FULL_NAME: FakeEntry(),
        F.SHORT_NAME: FakeEntry(),
        F.DIRECTOR: FakeEntry(),
        F.LEGAL_ADDRESS: FakeEntry(),
        F.FACT_ADDRESS: FakeEntry(),
        F.ACCOUNT: FakeEntry(),
        F.CORR_ACCOUNT: FakeEntry(),
        F.BANK: FakeEntry(),
        F.BIK: FakeEntry(),
    }


def ooo_entries(inn: str = "7701234567") -> dict:
    return make_entries(inn)


OOO_CARD = {
    "inn": "7701234567", "kpp": "770101001", "ogrn": "1037739010221",
    "full_name": "ООО Ромашка Плюс", "short_name": "ООО Ромашка",
    "director": "Иванов Иван Иванович",
    "legal_address": "г. Москва, ул. Тестовая, д. 1",
    "bank_name": "Тест-Банк", "bik": "044525225",
    "correspondent_account": "30101810100000000225",
}


# --------------------------------------------------------------------------- #
#  Guard и валидация перед запуском
# --------------------------------------------------------------------------- #
def test_start_blocked_while_busy():
    flow, app = make_flow()
    flow.busy = True
    factory = RecordingFactory()

    flow.fill_from_inn(thread_factory=factory)

    assert factory.created == []
    assert flow.busy is True, "guard не сбрасывает состояние активного запроса"
    assert app.root.after_calls == []


def test_inn_length_validated_per_mode(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries("7701234567")  # 10 цифр — верно
    app._cp_entries["ip"] = ooo_entries("312301234567")  # 12 цифр

    # ООО с 9 цифрами — ошибка
    app._cp_entries["ooo"][F.INN] = FakeEntry("770123456")
    flow.fill_from_inn(thread_factory=RecordingFactory())
    assert alerts and "10 цифр" in alerts[-1][2]

    # ИП с 10 цифрами — ошибка «12 цифр» (регрессия дефекта 2.0.0)
    alerts.clear()
    app._cp_type_var.set("ИП")
    app._cp_entries["ip"][F.INN] = FakeEntry("7701234567")
    flow.fill_from_inn(thread_factory=RecordingFactory())
    assert alerts and "12 цифр" in alerts[-1][2]

    # нецифровой ИНН — тоже ошибка
    alerts.clear()
    app._cp_entries["ip"][F.INN] = FakeEntry("абв")
    flow.fill_from_inn(thread_factory=RecordingFactory())
    assert alerts


def test_missing_api_key_reports(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    monkeypatch.setattr(dadata_client, "get_api_key", lambda: "")
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()

    flow.fill_from_inn(thread_factory=RecordingFactory())

    assert alerts and "Нет API-ключа" in alerts[0][1]
    assert flow.busy is False


def test_start_schedules_queue_poll_and_worker(monkeypatch):
    monkeypatch.setattr(dadata_client, "get_api_key", lambda: "test-key")
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()
    factory = RecordingFactory()

    flow.fill_from_inn(thread_factory=factory)

    assert flow.busy is True
    assert flow.result_queue is not None
    assert len(factory.created) == 1
    thread = factory.created[0]
    assert thread.started is True
    assert thread.daemon is True
    assert thread.args == ("7701234567", "test-key", "ooo", flow.result_queue)
    # busy-анимация запущена и опрос очереди перепланирован
    assert app._cp_status_var.history and "Загружаю реквизиты" in \
        app._cp_status_var.history[0]
    assert any(ms == 80 and fn == app._cp_result_tick
               for ms, fn in app.root.after_calls)


# --------------------------------------------------------------------------- #
#  Воркер и опрос результата
# --------------------------------------------------------------------------- #
def test_worker_puts_ok_result(monkeypatch):
    monkeypatch.setattr(dadata_client, "fetch_party",
                        lambda inn, key: dict(OOO_CARD))
    flow, app = make_flow()
    q: queue.Queue = queue.Queue()

    flow.fetch_worker("7701234567", "key", "ooo", q)

    assert q.get_nowait() == ("ok", "ooo", OOO_CARD)


def test_worker_puts_error_result(monkeypatch):
    def broken(inn, key):
        raise RuntimeError("сеть недоступна")

    monkeypatch.setattr(dadata_client, "fetch_party", broken)
    flow, _ = make_flow()
    q: queue.Queue = queue.Queue()

    flow.fetch_worker("7701234567", "key", "ooo", q)

    kind, message = q.get_nowait()
    assert kind == "error" and "сеть недоступна" in message


def test_result_tick_handles_error(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_flow()
    flow.busy = True
    flow.busy_after = "after-1"
    q: queue.Queue = queue.Queue()
    q.put(("error", "таймаут"))
    flow.result_queue = q

    flow.result_tick()

    assert flow.busy is False and flow.result_queue is None
    assert alerts and alerts[0][1] == "Не удалось заполнить реквизиты"


def test_result_tick_reschedules_on_empty_queue():
    flow, app = make_flow()
    flow.result_queue = queue.Queue()

    flow.result_tick()

    assert app.root.after_calls[-1] == (80, app._cp_result_tick)


def test_result_tick_stops_without_queue():
    flow, app = make_flow()
    flow.result_queue = None

    flow.result_tick()

    assert app.root.after_calls == []


# --------------------------------------------------------------------------- #
#  Анимация ожидания
# --------------------------------------------------------------------------- #
def test_busy_tick_animates_and_reschedules():
    flow, app = make_flow()
    flow.busy = True

    flow.busy_tick()
    flow.busy_tick()

    assert app._cp_status_var.history[0].endswith("Загружаю реквизиты из ЕГРЮЛ/ЕГРИП")
    assert app._cp_status_var.history[1].endswith("Загружаю реквизиты из ЕГРЮЛ/ЕГРИП.")
    assert app.root.after_calls[-1] == (300, app._cp_busy_tick)
    assert flow.busy_phase == 2


def test_busy_stop_cancels_timer_and_queue():
    flow, app = make_flow()
    flow.busy = True
    flow.busy_after = "after-9"
    flow.result_queue = queue.Queue()

    flow.busy_stop()

    assert flow.busy is False
    assert flow.result_queue is None
    assert app.root.cancelled == ["after-9"]


def test_busy_tick_ignores_when_not_busy():
    flow, app = make_flow()

    flow.busy_tick()

    assert app._cp_status_var.history == []
    assert app.root.after_calls == []


# --------------------------------------------------------------------------- #
#  Подстановка реквизитов
# --------------------------------------------------------------------------- #
def test_party_loaded_fills_ooo_fields():
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()

    flow.party_loaded("ooo", dict(OOO_CARD))

    entries = app._cp_entries["ooo"]
    assert entries[F.INN].get() == "7701234567"
    assert entries[F.FULL_NAME].get() == "ООО Ромашка Плюс"
    assert entries[F.SHORT_NAME].get() == "ООО Ромашка"
    assert entries[F.DIRECTOR].get() == "Иванов Иван Иванович"
    assert entries[F.KPP].get() == "770101001"
    assert entries[F.OGRN].get() == "1037739010221"
    assert entries[F.BANK].get() == "Тест-Банк"
    assert entries[F.BIK].get() == "044525225"
    assert app._cp_status_var.value == "Реквизиты заполнены из ЕГРЮЛ/ЕГРИП"
    assert flow.busy is False, "анимация остановлена после успеха"


def test_party_loaded_fills_fact_address_only_when_empty():
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()

    flow.party_loaded("ooo", dict(OOO_CARD))
    assert app._cp_entries["ooo"][F.FACT_ADDRESS].get() == OOO_CARD["legal_address"]

    # непустой фактический адрес не затирается
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()
    app._cp_entries["ooo"][F.FACT_ADDRESS] = FakeEntry("г. Москва, ул. Своя, д. 2")
    flow.party_loaded("ooo", dict(OOO_CARD))
    assert app._cp_entries["ooo"][F.FACT_ADDRESS].get() == "г. Москва, ул. Своя, д. 2"


def test_party_loaded_fills_ip_fields_without_kpp():
    flow, app = make_flow()
    app._cp_entries["ip"] = ooo_entries("312301234567")
    ip_card = {
        "inn": "312301234567", "kpp": "", "ogrn": "304500100000340",
        "full_name": "ИП Петров Пётр Петрович", "short_name": "ИП Петров",
        "director": "Петров Пётр Петрович",
        "legal_address": "г. Москва, ул. Тестовая, д. 1",
        "bank_name": "Тест-Банк", "bik": "044525000",
        "correspondent_account": "30101810000000000225",
    }

    flow.party_loaded("ip", ip_card)

    entries = app._cp_entries["ip"]
    assert entries[F.IP_NAME].get() == "Петров Пётр Петрович"
    assert entries[F.OGRNIP].get() == "304500100000340"
    assert entries[F.KPP].get() == "", "КПП в карточку ИП не подставляется"


def test_party_loaded_without_data_alerts(monkeypatch):
    alerts: list = []
    monkeypatch.setattr(dc.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_flow()
    app._cp_entries["ooo"] = ooo_entries()
    empty = {k: "" for k in OOO_CARD}

    flow.party_loaded("ooo", empty)

    assert alerts and alerts[0][1] == "Данных нет"


# --------------------------------------------------------------------------- #
#  Фасад ContractApp: состояние busy через прежние имена
# --------------------------------------------------------------------------- #
def _bare_app() -> "main.ContractApp":
    app = object.__new__(main.ContractApp)
    app._dadata_flow = DadataController(app)
    return app


def test_facade_busy_property_projects_controller_state():
    app = _bare_app()

    app._cp_busy = True
    assert app._dadata_flow.busy is True
    app._cp_busy = False
    assert app._dadata_flow.busy is False


def test_facade_guard_blocks_second_click_while_busy():
    """Зеркало smoke-проверки: повторный клик при активном запросе — no-op."""
    app = _bare_app()
    app._cp_busy = True

    app._cp_fill_from_inn()  # guard срабатывает до обращения к полям

    assert app._cp_busy is True, "guard не сбросил состояние"
    assert app._dadata_flow.result_queue is None
