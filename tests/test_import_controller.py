"""Юнит-тесты ImportController — первый контроллер декомпозиции ContractApp.

Аудит 2.0.1 (P1 «старт покрытия main.py», P2 «контроллеры страниц»): логика
пакетного импорта теперь живёт в import_controller.py и проверяется без
Tk — контроллер получает фейковое приложение (duck typing), а smoke-тест
продолжает покрывать GUI-путь целиком.

Контракты, закреплённые тестами (из аудита 2.0.1):
  * guard повторного входа в начале запуска (drag-and-drop не порождает
    второй импорт и не сбрасывает очередь работающего первого);
  * пакет валиден/отклонён по количеству, объёму, папке и свободному месту;
  * воркер кладёт результаты в СВОЮ очередь и завершается ("done",);
  * регистрация: attach по ИНН, create без дублей, «unknown» для отбоя;
  * пакетная запись реестра одним save() (п. 5.2) с честным отчётом об ошибках.
"""
from __future__ import annotations

import logging
from types import SimpleNamespace
import os
import queue
import threading

import pytest
import tkinter as tk

import import_controller as ic
import settings_controller as scc
import storage
from document_builder import F
from document_importer import MAX_TOTAL_SIZE


# --------------------------------------------------------------------------- #
#  Фейки: приложение, хранилище, виджеты
# --------------------------------------------------------------------------- #
class FakeRoot:
    def __init__(self):
        self.after_calls: list = []

    def after(self, ms, fn):
        self.after_calls.append((ms, fn))


class FakeVar:
    def __init__(self, value=""):
        self.value = value
        self.history: list = []

    def set(self, value):
        self.value = value
        self.history.append(value)


class FakeWin:
    def __init__(self):
        self.destroyed = False

    def grab_release(self):
        pass

    def destroy(self):
        self.destroyed = True


class FakeStore:
    """Минимум Storage для контроллера: поиск, добавление, пакетный save()."""

    def __init__(self):
        self.counterparties: list[dict] = []
        self.documents: list[dict] = []
        self.save_calls: list = []
        self.save_ok = True
        self.last_error = ""
        self._next_id = 1

    def find_by_inn(self, inn):
        for rec in self.counterparties:
            if str((rec.get("fields") or {}).get(F.INN, "")).strip() == inn:
                return rec
        return None

    def find_by_name(self, display):
        for rec in self.counterparties:
            if storage.display_name(rec) == display:
                return rec
        return None

    def add_counterparty(self, mode, fields, save=True):
        rec = {"id": str(self._next_id), "type": mode, "fields": dict(fields)}
        self._next_id += 1
        self.counterparties.append(rec)
        return rec

    def add_document(self, kind, party_id, number, date,
                     data=None, extra=None, text="", save=True):
        self.documents.append({
            "kind": kind, "party_id": party_id,
            "number": number, "date": date,
            "data": data or {}, "extra": extra or {},
            "text": text, "saved": save,
        })
        return self.documents[-1]

    def documents_for(self, party_id):
        return [d for d in self.documents if d["party_id"] == party_id]

    def save(self):
        self.save_calls.append(1)
        return self.save_ok


class FakeApp:
    """Duck-typed приложение: только то, что трогает ImportController."""

    def __init__(self, tmp_path):
        self.root = FakeRoot()
        self.scale = 1.0
        self.store = FakeStore()
        self._log = logging.getLogger("test-import-controller")
        self.calls: list = []
        self.notifications: list[str] = []
        self.ask_party_result: tuple = (None, "skip")
        self.choice_result: str = "skip"
        self.party_from_var_result: dict | None = None
        # Соседние контроллеры вызываются напрямую (v2.12.0, фасады
        # ContractApp удалены): записывающие фейки.
        self._settings_flow = SimpleNamespace(
            documents_dir_for_import=lambda: str(tmp_path),
            is_inside_directory=SettingsController.is_inside_directory,
        )
        self._doc_context_flow = SimpleNamespace(
            party_from_var=lambda var: self.party_from_var_result,
            refresh_party_combo_values=lambda: self.calls.append("combo"),
        )
        self._doc_page_flow = SimpleNamespace(
            preview_contract=lambda *a: self.calls.append("preview_contract"),
            preview_additional=lambda *a: self.calls.append("preview_additional"),
        )
        self._registry_flow = SimpleNamespace(
            refresh=lambda: self.calls.append("registry"))

    def _notify(self, message, path=None):
        self.notifications.append(message)

    def _update_status(self):
        self.calls.append("status")

    def _choice_dialog(self, *args, **kwargs):
        return self.choice_result


def make_flow(tmp_path) -> tuple[ImportController, FakeApp]:
    app = FakeApp(tmp_path)
    flow = ic.ImportController(app)
    # Собственные методы потока вызываются напрямую (v2.12.0): отчёт —
    # записывающая подмена, finish — обёртка с записью результатов.
    flow.show_report = lambda entries, summary: app.calls.append(
        ("report", entries, summary))
    flow.ask_party = lambda result: app.ask_party_result
    real_finish = flow.finish

    def _finish(results):
        # снимок: реальный finish/register_result мутирует записи на месте
        app.calls.append(("finish", [dict(r) for r in results]))
        return real_finish(results)

    flow.finish = _finish
    return flow, app


# для читаемости аннотаций тестов
ImportController = ic.ImportController
SettingsController = scc.SettingsController


def zero_stats() -> dict:
    return {"added": 0, "new_parties": 0, "skipped": 0,
            "duplicates": 0, "cancelled": 0}


def drain(q: "queue.Queue") -> list:
    items = []
    while True:
        try:
            items.append(q.get_nowait())
        except queue.Empty:
            return items


class RecordingFactory:
    """Фабрика потока, которая не запускает воркеры, а записывает вызовы."""

    def __init__(self):
        self.created: list = []

    def __call__(self, target=None, args=(), daemon=None, **_kw):
        self.created.append((target, args, daemon))
        return self


# --------------------------------------------------------------------------- #
#  Guard повторного входа и проверки пакета
# --------------------------------------------------------------------------- #
def test_start_blocked_while_running(tmp_path):
    flow, app = make_flow(tmp_path)
    flow.running = True
    factory = RecordingFactory()

    flow.start(["whatever.docx"], thread_factory=factory)

    assert app.notifications == ["Импорт уже выполняется — дождитесь завершения."]
    assert factory.created == []
    assert flow.queue is None and flow.win is None


def test_start_empty_files_is_noop(tmp_path):
    flow, app = make_flow(tmp_path)
    factory = RecordingFactory()

    flow.start([], thread_factory=factory)

    assert factory.created == []
    assert flow.running is False
    assert app.notifications == []


def test_start_rejects_too_many_files(tmp_path, monkeypatch):
    alerts: list = []
    monkeypatch.setattr(ic.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, app = make_flow(tmp_path)
    files = [str(tmp_path / f"f{i}.docx") for i in range(201)]

    flow.start(files, thread_factory=RecordingFactory())

    assert alerts, "пользователю показано предупреждение о лимите"
    assert "Слишком много файлов" in alerts[0][1]
    assert alerts[0][3] == "warning"
    assert flow.running is False


def test_start_rejects_oversize_package(tmp_path, monkeypatch):
    alerts: list = []
    monkeypatch.setattr(ic.ui, "alert", lambda *a, **kw: alerts.append(a))
    flow, _ = make_flow(tmp_path)
    big = tmp_path / "big.docx"
    with open(big, "wb") as fh:
        fh.truncate(MAX_TOTAL_SIZE + 1)  # разреженный файл: место не занимает

    flow.start([str(big)], thread_factory=RecordingFactory())

    assert alerts and "Слишком большой пакет" in alerts[0][1]


def test_start_reports_unavailable_documents_dir(tmp_path, monkeypatch):
    alerts: list = []
    monkeypatch.setattr(ic.ui, "alert", lambda *a, **kw: alerts.append(a))
    app = FakeApp(tmp_path)

    def broken_dir():
        raise OSError("папка недоступна")

    app._settings_flow = SimpleNamespace(documents_dir_for_import=broken_dir)
    flow = ic.ImportController(app)
    small = tmp_path / "ok.docx"
    small.write_bytes(b"PK")

    flow.start([str(small)], thread_factory=RecordingFactory())

    assert alerts and alerts[0][1] == "Папка документов недоступна"
    assert alerts[0][3] == "error"


def test_start_rejects_when_disk_almost_full(tmp_path, monkeypatch):
    alerts: list = []
    monkeypatch.setattr(ic.ui, "alert", lambda *a, **kw: alerts.append(a))
    monkeypatch.setattr(ic.shutil, "disk_usage",
                        lambda path: type("Usage", (), {"free": 0})())
    flow, _ = make_flow(tmp_path)
    small = tmp_path / "ok.docx"
    small.write_bytes(b"PK")

    flow.start([str(small)], thread_factory=RecordingFactory())

    assert alerts and "Недостаточно места на диске" in alerts[0][1]


def test_pack_problem_accepts_valid_package(tmp_path):
    flow, _ = make_flow(tmp_path)
    small = tmp_path / "ok.docx"
    small.write_bytes(b"PK\x03\x04")

    assert flow._pack_problem([str(small)]) is None


# --------------------------------------------------------------------------- #
#  Воркер и копирование
# --------------------------------------------------------------------------- #
def test_worker_rejects_disallowed_extension(tmp_path):
    flow, _ = make_flow(tmp_path)
    bad = tmp_path / "danger.txt"
    bad.write_bytes(b"x")
    q: queue.Queue = queue.Queue()

    flow.worker([str(bad)], q)

    items = drain(q)
    assert items[-1] == ("done",)
    file_item = next(i[1] for i in items if i[0] == "file")
    assert file_item["ok"] is False
    assert "не поддерживается" in file_item["error"]
    assert file_item["copied"] == ""


def test_worker_analyzes_and_copies(tmp_path, monkeypatch):
    flow, app = make_flow(tmp_path)
    doc = tmp_path / "Договор №5.docx"
    doc.write_bytes(b"PK\x03\x04fake")
    monkeypatch.setattr(
        ic.importer, "analyze_file",
        lambda path: {
            "ok": True, "recognized": True,
            "kind": storage.KIND_CONTRACT, "number": "5", "date": "12 января",
            "party": {"type": "ooo", "display": "ООО Ромашка",
                      "fields": {F.INN: "7701234567"}},
            "text": "текст", "path": str(doc), "error": "", "warning": "",
        },
    )
    q: queue.Queue = queue.Queue()

    flow.worker([str(doc)], q)

    items = drain(q)
    assert items[-1] == ("done",)
    file_item = items[1][1]
    assert file_item["copied"], "разобранный документ скопирован"
    assert "ООО Ромашка" in file_item["copied"]
    assert os.path.exists(file_item["copied"])


def test_worker_honours_cancellation(tmp_path):
    flow, _ = make_flow(tmp_path)
    flow.cancelled = True
    files = [str(tmp_path / f"a{i}.docx") for i in range(3)]
    for name in files:
        with open(name, "wb") as fh:
            fh.write(b"PK")
    q: queue.Queue = queue.Queue()

    flow.worker(files, q)

    assert drain(q) == [("done",)], "после отмены остаток не анализируется"


# --------------------------------------------------------------------------- #
#  Регистрация результатов
# --------------------------------------------------------------------------- #
RECOGNIZED = {
    "ok": True, "recognized": True,
    "kind": storage.KIND_CONTRACT, "number": "5", "date": "12 января",
    "party": {"type": "ooo", "display": "ООО Ромашка",
              "fields": {F.INN: "7701234567"}},
    "text": "", "error": "", "warning": "", "copied": "",
}


def test_register_skips_unrecognized_and_moves_to_unknown(tmp_path):
    flow, app = make_flow(tmp_path)
    flow.move_to_unknown = lambda path: "moved-to-unknown"
    result = {"ok": False, "error": "битый файл", "copied": "some.docx",
              "party": None, "name": "broken.docx", "kind": storage.KIND_UNKNOWN}
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is False
    assert stats["skipped"] == 1
    assert entry["detail"] == "битый файл"
    assert result["copied"] == "moved-to-unknown"


def test_register_attaches_by_inn_without_duplicate_save(tmp_path):
    flow, app = make_flow(tmp_path)
    rec = app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    result = dict(RECOGNIZED, copied="copied.docx")
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is True
    assert stats["added"] == 1
    assert entry["status"] == "added"
    assert app.store.documents[0]["party_id"] == rec["id"]
    assert app.store.documents[0]["saved"] is False, "п. 5.2: без записи на диск"


def test_register_creates_new_party(tmp_path):
    flow, app = make_flow(tmp_path)
    result = dict(RECOGNIZED, copied="copied.docx")
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is True
    assert stats["new_parties"] == 1
    assert entry["detail"] == "создана новая карточка"
    assert app.store.find_by_inn("7701234567") is not None


def test_register_skips_duplicate_document(tmp_path):
    flow, app = make_flow(tmp_path)
    rec = app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    app.store.add_document(storage.KIND_CONTRACT, rec["id"], "5", "12 января")
    result = dict(RECOGNIZED, copied="copied.docx")
    flow.move_to_unknown = lambda path: "moved"
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is False
    assert stats["duplicates"] == 1
    assert entry["status"] == "duplicate"
    assert result["copied"] == "moved"


def test_register_uses_ask_party_facade_for_unrecognized(tmp_path):
    flow, app = make_flow(tmp_path)
    rec = app.store.add_counterparty("ooo", {F.INN: ""})
    app.ask_party_result = (rec, "attach")
    result = dict(RECOGNIZED, recognized=False, copied="copied.docx")
    result["party"] = {"display": "ООО Ромашка", "fields": {}}
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is True
    assert entry["warning"] == "карточка создана без ИНН"


def test_register_skips_when_file_was_not_copied(tmp_path):
    flow, app = make_flow(tmp_path)
    app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    result = dict(RECOGNIZED, copied="")
    entry = ic.ImportController.new_entry(result)
    stats = zero_stats()

    assert flow.register_result(result, entry, stats) is False
    assert stats["skipped"] == 1
    assert entry["detail"] == "файл не скопирован"


# --------------------------------------------------------------------------- #
#  Дубликаты
# --------------------------------------------------------------------------- #
def test_duplicate_detected_by_kind_number_date(tmp_path):
    flow, app = make_flow(tmp_path)
    rec = app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    app.store.add_document(storage.KIND_CONTRACT, rec["id"], " 5 ", "12 января")

    assert flow.document_is_duplicate(
        rec["id"], {"kind": storage.KIND_CONTRACT, "number": "5", "date": "12 января"}
    )
    assert not flow.document_is_duplicate(
        rec["id"], {"kind": storage.KIND_CONTRACT, "number": "6", "date": "12 января"}
    )
    assert not flow.document_is_duplicate(
        rec["id"], {"kind": storage.KIND_SPEC, "number": "5", "date": "12 января"}
    )


def test_duplicate_needs_number_or_date(tmp_path):
    flow, app = make_flow(tmp_path)

    assert not flow.document_is_duplicate(
        "any", {"kind": storage.KIND_CONTRACT, "number": "", "date": ""}
    )


# --------------------------------------------------------------------------- #
#  Диалог привязки без ИНН
# --------------------------------------------------------------------------- #
def test_ask_party_attaches_existing_by_name(tmp_path, monkeypatch):
    flow, app = make_flow(tmp_path)
    rec = app.store.add_counterparty(
        "ooo", {F.INN: "7701234567", "Сокращенное название компании": "ООО Ромашка"}
    )
    monkeypatch.setattr(ic.ui, "confirm", lambda *a, **kw: True)
    result = {"name": "файл.docx", "party": {"display": "ООО Ромашка",
                                             "fields": {}, "missing": "ИНН"}}

    del flow.ask_party  # снять подмену из make_flow — нужен настоящий диалог

    got_rec, action = flow.ask_party(result)

    assert action == "attach" and got_rec is rec


def test_ask_party_choice_dialog_routes_through_app(tmp_path, monkeypatch):
    flow, app = make_flow(tmp_path)
    monkeypatch.setattr(ic.ui, "confirm", lambda *a, **kw: False)
    result = {"name": "файл.docx", "party": {"display": "ООО Ромашка",
                                             "fields": {}, "missing": "ИНН"}}
    captured: list = []
    monkeypatch.setattr(app, "_choice_dialog",
                        lambda *a, **kw: captured.append(a) or "create")

    del flow.ask_party  # снять подмену из make_flow — нужен настоящий диалог

    got_rec, action = flow.ask_party(result)

    assert action == "create" and got_rec is None
    assert captured, "диалог вызван через фасад приложения"


def test_ask_party_offers_current_party(tmp_path):
    flow, app = make_flow(tmp_path)
    current = {"id": "42", "type": "ooo", "fields": {F.INN: "5000000001"}}
    app.contract_party_var = FakeVar("ООО Ромашка (ИНН 5000000001)")
    app.party_from_var_result = current
    app.choice_result = "attach"
    result = {"name": "файл.docx", "party": {"display": "Кто-то новый",
                                             "fields": {}, "missing": "ИНН"}}

    del flow.ask_party  # снять подмену из make_flow — нужен настоящий диалог

    got_rec, action = flow.ask_party(result)

    assert action == "attach" and got_rec is current


# --------------------------------------------------------------------------- #
#  Завершение пакета: один save() и отчёт
# --------------------------------------------------------------------------- #
def test_finish_saves_once_and_reports(tmp_path):
    flow, app = make_flow(tmp_path)
    app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    # Номера различаются: одинаковые номер+дата честно ловились бы как дубликаты.
    results = [dict(RECOGNIZED, copied=f"copied-{i}.docx", number=str(i + 1))
               for i in range(3)]

    flow.finish(results)

    assert app.store.save_calls == [1], "п. 5.2: пакет записан одним save()"
    report = next(c for c in app.calls if isinstance(c, tuple) and c[0] == "report")
    entries, summary = report[1], report[2]
    assert len(entries) == 3
    assert "Добавлено документов: 3" in summary
    assert "combo" in app.calls and "registry" in app.calls


def test_finish_surfaces_save_errors(tmp_path):
    flow, app = make_flow(tmp_path)
    app.store.add_counterparty("ooo", {F.INN: "7701234567"})
    app.store.save_ok = False
    app.store.last_error = "диск переполнен"

    flow.finish([dict(RECOGNIZED, copied="copied.docx")])

    report = next(c for c in app.calls if isinstance(c, tuple) and c[0] == "report")
    assert "не попала на диск" in report[2]
    assert "диск переполнен" in report[2]


def test_finish_without_additions_skips_previews(tmp_path):
    flow, app = make_flow(tmp_path)

    flow.finish([{"ok": False, "error": "не распознан", "copied": "",
                  "party": None, "name": "x.docx", "kind": storage.KIND_UNKNOWN}])

    assert "combo" not in app.calls
    assert "registry" in app.calls and "status" in app.calls


def test_finish_counts_cancelled(tmp_path):
    flow, app = make_flow(tmp_path)
    flow.cancelled_count = 2
    results = [dict(RECOGNIZED, copied="c.docx")]
    app.store.add_counterparty("ooo", {F.INN: "7701234567"})

    flow.finish(results)

    report = next(c for c in app.calls if isinstance(c, tuple) and c[0] == "report")
    assert "Отменено пользователем: 2" in report[2]
    assert flow.cancelled_count == 0, "счётчик сброшен после отчёта"


# --------------------------------------------------------------------------- #
#  Опрос очереди (tick)
# --------------------------------------------------------------------------- #
def test_tick_processes_status_file_done(tmp_path):
    flow, app = make_flow(tmp_path)
    q: queue.Queue = queue.Queue()
    q.put(("status", "Анализ файла 1 из 1:\ndoc.docx"))
    q.put(("file", {"name": "doc.docx"}))
    q.put(("done",))
    flow.queue = q
    flow.results = []
    flow.total = 1
    flow.running = True
    win = FakeWin()
    flow.win = win
    flow.status_var = FakeVar("Подготовка…")
    flow.progress = FakeVar(0)

    flow.tick()

    assert flow.status_var.history[-1].startswith("Анализ файла")
    assert win.destroyed
    assert flow.running is False
    assert flow.queue is None
    assert app.root.after_calls == [], "после done опрос не перепланируется"
    finish_rec = next(c for c in app.calls if isinstance(c, tuple) and c[0] == "finish")
    assert finish_rec[1] == [{"name": "doc.docx"}]


def test_tick_reschedules_on_empty_queue(tmp_path):
    flow, app = make_flow(tmp_path)
    flow.queue = queue.Queue()
    flow.status_var = FakeVar()
    flow.progress = FakeVar(0)
    flow.win = None

    flow.tick()

    assert app.root.after_calls and app.root.after_calls[0][0] == 80
    assert app.root.after_calls[0][1] == flow.tick, "перепланируется собственный tick"


def test_tick_stops_when_window_is_dying(tmp_path):
    flow, app = make_flow(tmp_path)
    q: queue.Queue = queue.Queue()
    q.put(("status", "x"))
    flow.queue = q

    class DyingVar:
        @staticmethod
        def set(_value):
            raise tk.TclError("window destroyed")

    flow.status_var = DyingVar()
    flow.progress = FakeVar(0)

    flow.tick()

    assert app.root.after_calls == [], "TclError прекращает опрос"


# --------------------------------------------------------------------------- #
#  Guard повторного входа и отмена: напрямую через контроллер (v2.12.0)
# --------------------------------------------------------------------------- #
def test_guard_blocks_second_start_via_drag_and_drop():
    """Зеркало smoke-проверки re-entry: второй запуск не сбрасывает состояние."""
    flow = ic.ImportController(SimpleNamespace(
        _notify=lambda *_a, **_kw: None, _settings_flow=SimpleNamespace()))
    flow.running = True
    flow.queue = None
    flow.win = None

    flow.start(["guard.docx"])

    assert flow.queue is None and flow.win is None
    assert flow.running is True, "guard не сбросил состояние"


def test_empty_files_do_not_start_worker():
    flow = ic.ImportController(SimpleNamespace(
        _notify=lambda *_a, **_kw: None, _settings_flow=SimpleNamespace()))

    flow.start([])

    assert flow.running is False
    assert flow.queue is None


def test_cancel_routes_to_state():
    flow = ic.ImportController(SimpleNamespace(_status_var=FakeVar("")))

    flow.cancel()  # импорт не выполняется — no-op
    assert flow.cancelled is False

    flow.running = True
    flow.cancel()
    assert flow.cancelled is True
