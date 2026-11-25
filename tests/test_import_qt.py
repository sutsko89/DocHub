r"""Тесты M6-b — пакетный импорт на Qt (перенос test_import_controller).

Контракты Tk-импорта сохраняются дословно:

* guard повторного входа в начале запуска (аудит 2.0.1, P1); очередь у
  каждого запуска своя (воркеру через args);
* лимиты пакета до запуска: 200 файлов, суммарный объём, доступность
  папки документов, свободное место (п. 3.4);
* воркер: allowlist .docx/.xlsx, разбор, копирование разобранных файлов
  в папку контрагента, отмена остатка пакета (п. 5.3);
* U9: серия файлов без ИНН и серия конфликтов названий — ОДИН диалог
  сводки на серию; критерий плана: импорт 20 файлов без ИНН — 1–2
  диалога вместо 20+; «создать по названиям» переиспользует карточки
  с тем же названием и группирует файлы одного названия в одну карточку;
* регистрация: авто-карточка только при найденном ИНН; дубликаты
  (тип + номер + дата) повторно не добавляются; отклонённые файлы — в
  папку unknown (п. 3.5); пакетная запись ОДНИМ save() (п. 5.2);
* отчёт с «Открыть реестр» (U10); точки входа — кнопка реестра и
  перетаскивание файлов в окно (п. 4.5).

Диалоги подменяются на уровне qt_dialogs (маршруты контроллеров —
риски M6 плана); воркер запускается синхронной фабрикой потока.
"""

from __future__ import annotations

import os
import queue
import sys

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import import_qt
import main_qt
import qt_dialogs
import shell_open
import storage
from document_builder import F
from document_importer import MAX_TOTAL_SIZE


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
class SyncThread:
    """Фабрика потока, исполняющая воркер синхронно (как в Tk-тестах)."""

    def __init__(self, target=None, args=(), daemon=None, **_kw):
        self._target = target
        self._args = args

    def start(self) -> None:
        self._target(*self._args)


RECOGNIZED = {
    "ok": True, "recognized": True,
    "kind": storage.KIND_CONTRACT, "number": "5", "date": "12 января",
    "party": {"type": "ooo", "display": "ООО Ромашка",
              "fields": {F.INN: "7701234567"}},
    "text": "", "error": "", "warning": "", "copied": "",
}

NO_INN = {
    "ok": True, "recognized": False,
    "kind": storage.KIND_CONTRACT, "number": "", "date": "",
    "party": {"type": "ooo", "display": "ООО Старый файл", "fields": {}},
    "text": "", "error": "", "warning": "", "copied": "",
}


@pytest.fixture()
def stub_dialogs(monkeypatch):
    """Диалоги: подтверждения True, алерты/выборы собираются."""
    captured = {
        "confirms": [], "alerts": [], "choices": [],
        "choice_result": None, "confirm_result": True,
    }

    def fake_confirm(*args, **kwargs):
        captured["confirms"].append(args)
        return captured["confirm_result"]

    def fake_choice(*args, **kwargs):
        captured["choices"].append(args)
        return captured["choice_result"]

    captured["confirm"] = fake_confirm
    captured["choice"] = fake_choice
    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(qt_dialogs, "alert",
                        lambda *a, **k: captured["alerts"].append(a))
    monkeypatch.setattr(qt_dialogs, "choice", fake_choice)
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs, monkeypatch, tmp_path):
    """Каркас с живыми страницами; открытие файлов и папка документов — tmp."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    monkeypatch.setattr(shell_open, "documents_known_folder", lambda: str(tmp_path))
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl.startup_flow.stop()
    ctrl.stop_poll_timer("startup_tick")
    ctrl._stop_timers()
    ctrl.stop_poll_timer("import_tick")
    for flow in (ctrl.spec_flow, ctrl.contract_flow, ctrl.additional_flow):
        if getattr(flow, "preview_timer", None) is not None:
            flow.preview_timer.stop()


@pytest.fixture(autouse=True)
def patch_report(monkeypatch):
    """Отчёт импорта — записывающая подмена на уровне контроллера."""
    reports: list = []

    def fake_report(self, entries, summary):
        reports.append((entries, summary))

    monkeypatch.setattr(import_qt.ImportFlowQt, "show_report", fake_report)
    return reports


def start_sync(controller, files):
    """Запускает импорт синхронно и выкачивает очередь до 'done'."""
    flow = controller.import_flow
    flow.start(files, thread_factory=SyncThread)
    # Воркер уже отработал (синхронная фабрика) — опрос очереди вручную.
    flow.tick()


def make_files(tmp_path, count, prefix="Договор №"):
    """Реальные файлы .docx (воркер копирует содержимое на диск)."""
    paths = []
    for i in range(count):
        path = tmp_path / f"{prefix}{i + 1}.docx"
        path.write_bytes(b"PK\x03\x04fake")
        paths.append(str(path))
    return paths


# --------------------------------------------------------------------------- #
#  Guard повторного входа и проверки пакета (п. 3.4, аудит 2.0.1)
# --------------------------------------------------------------------------- #
def test_start_blocked_while_running(controller, stub_dialogs):
    flow = controller.import_flow
    flow.running = True
    flow.start(["whatever.docx"], thread_factory=SyncThread)
    assert flow.results == [] and flow.queue is None
    assert "Импорт уже выполняется" in controller.toast_label.text()


def test_start_empty_files_is_noop(controller):
    flow = controller.import_flow
    flow.start([], thread_factory=SyncThread)
    assert flow.running is False
    assert flow.win is None


def test_start_rejects_too_many_files(controller, tmp_path, stub_dialogs):
    files = make_files(tmp_path, 201)
    controller.import_flow.start(files, thread_factory=SyncThread)
    assert controller.import_flow.running is False
    assert any("Слишком много файлов" == str(a[1])
               for a in stub_dialogs["alerts"])


def test_start_rejects_oversize_package(controller, tmp_path, stub_dialogs):
    big = tmp_path / "big.docx"
    with open(big, "wb") as fh:
        fh.truncate(MAX_TOTAL_SIZE + 1)  # разреженный файл
    controller.import_flow.start([str(big)], thread_factory=SyncThread)
    assert any("Слишком большой пакет" == str(a[1])
               for a in stub_dialogs["alerts"])


def test_start_reports_unavailable_documents_dir(controller, tmp_path,
                                                 stub_dialogs, monkeypatch):
    def broken():
        raise OSError("папка недоступна")

    monkeypatch.setattr(controller.import_flow, "documents_dir_for_import", broken)
    small = tmp_path / "ok.docx"
    small.write_bytes(b"PK\x03\x04")
    controller.import_flow.start([str(small)], thread_factory=SyncThread)
    assert controller.import_flow.running is False
    assert any("Папка документов недоступна" == str(a[1])
               for a in stub_dialogs["alerts"])


def test_start_rejects_when_disk_almost_full(controller, tmp_path,
                                             stub_dialogs, monkeypatch):
    monkeypatch.setattr(
        import_qt.shutil, "disk_usage",
        lambda path: type("Usage", (), {"free": 0})(),
    )
    small = tmp_path / "ok.docx"
    small.write_bytes(b"PK\x03\x04")
    controller.import_flow.start([str(small)], thread_factory=SyncThread)
    assert any("Недостаточно места на диске" == str(a[1])
               for a in stub_dialogs["alerts"])


# --------------------------------------------------------------------------- #
#  Воркер: allowlist, разбор, копирование, отмена (п. 5.3)
# --------------------------------------------------------------------------- #
def test_worker_rejects_disallowed_extension(controller, tmp_path):
    bad = tmp_path / "danger.txt"
    bad.write_bytes(b"x")
    q: queue.Queue = queue.Queue()
    controller.import_flow.worker([str(bad)], q)
    items = []
    while True:
        try:
            items.append(q.get_nowait())
        except queue.Empty:
            break
    assert items[-1] == ("done",)
    file_item = next(i[1] for i in items if i[0] == "file")
    assert file_item["ok"] is False
    assert "не поддерживается" in file_item["error"]
    assert file_item["copied"] == ""


def test_worker_analyzes_and_copies(controller, tmp_path, monkeypatch):
    doc = tmp_path / "Договор №5.docx"
    doc.write_bytes(b"PK\x03\x04fake")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(RECOGNIZED, path=path),
    )
    q: queue.Queue = queue.Queue()
    controller.import_flow.worker([str(doc)], q)
    items = []
    while True:
        try:
            items.append(q.get_nowait())
        except queue.Empty:
            break
    assert items[-1] == ("done",)
    file_item = items[1][1]
    assert file_item["copied"], "разобранный документ скопирован"
    assert "ООО Ромашка" in file_item["copied"]
    assert os.path.exists(file_item["copied"])


def test_worker_honours_cancellation(controller, tmp_path, monkeypatch):
    flow = controller.import_flow
    flow.cancelled = True
    files = make_files(tmp_path, 3)
    monkeypatch.setattr(import_qt.importer, "analyze_file", lambda path: dict(RECOGNIZED))
    q: queue.Queue = queue.Queue()
    flow.worker(files, q)
    assert q.get_nowait() == ("done",), "остаток пакета не анализируется (п. 5.3)"


# --------------------------------------------------------------------------- #
#  U9: пакетная сводка — 20 файлов без ИНН → один диалог (критерий M6)
# --------------------------------------------------------------------------- #
def test_u9_20_files_without_inn_single_dialog(controller, tmp_path,
                                               monkeypatch, stub_dialogs):
    """Импорт 20 файлов без ИНН — 1 диалог сводки вместо 20+ (критерий U9)."""
    files = make_files(tmp_path, 20, prefix="акт ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(NO_INN, path=path, name=os.path.basename(path)),
    )
    stub_dialogs["choice_result"] = "skip_all"
    start_sync(controller, files)
    # Ровно один пакетный диалог (плюс отчёт, который подменён).
    assert len(stub_dialogs["choices"]) == 1
    message = str(stub_dialogs["choices"][0][2])
    assert "20 файлов без ИНН" in message


def test_u9_create_by_name_reuses_and_groups(controller, tmp_path,
                                             monkeypatch, stub_dialogs):
    """«Создать по названиям»: карточка переиспользуется; файлы одного
    названия попадают в одну карточку — дубли не плодятся (карточка U9)."""
    files = make_files(tmp_path, 6, prefix="акт ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(NO_INN, path=path, name=os.path.basename(path)),
    )
    stub_dialogs["choice_result"] = "create_by_name"
    start_sync(controller, files)

    stats_text = controller.import_flow  # отчёт подменён — смотрим реестр
    parties = controller.store.active_counterparties()
    assert len(parties) == 1, "одно название — одна карточка"
    docs = [d for d in controller.store.documents
            if d["party_id"] == parties[0]["id"]]
    assert len(docs) == 6, "все файлы привязаны к одной карточке"
    assert parties[0]["fields"].get(F.INN, "") == "", "карточка без ИНН"


def test_u9_create_by_name_uses_existing_card(controller, tmp_path,
                                              monkeypatch, stub_dialogs):
    """Существующая карточка с тем же названием переиспользуется (U9)."""
    existing = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Старый файл", F.INN: "1111111111"},
        save=False,
    )
    files = make_files(tmp_path, 3, prefix="акт ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(
            NO_INN, path=path, name=os.path.basename(path),
            party={"type": "ooo", "display": "ООО Старый файл", "fields": {}},
        ),
    )
    stub_dialogs["choice_result"] = "create_by_name"
    start_sync(controller, files)

    parties = controller.store.active_counterparties()
    assert len(parties) == 1, "новая карточка не создаётся"
    assert parties[0]["id"] == existing["id"]
    docs = controller.store.documents_for(existing["id"])
    assert len(docs) == 3


def test_u9_attach_current_option(controller, tmp_path, monkeypatch, stub_dialogs):
    """«Привязать все к выбранному» — привязка к текущему контрагенту."""
    current = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Текущий", F.INN: "2222222222"},
        save=False,
    )
    controller.set_current_party(current)
    files = make_files(tmp_path, 4, prefix="акт ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(NO_INN, path=path, name=os.path.basename(path)),
    )
    stub_dialogs["choice_result"] = "attach_current"
    start_sync(controller, files)

    docs = controller.store.documents_for(current["id"])
    assert len(docs) == 4
    assert len(controller.store.active_counterparties()) == 1


def test_u9_conflicts_single_dialog(controller, tmp_path, monkeypatch,
                                    stub_dialogs):
    """Конфликт названий (другой ИНН) — один диалог сводки на серию."""
    controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Ромашка", F.INN: "9999999999"}, save=False,
    )
    files = make_files(tmp_path, 3, prefix="пакет ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(
            RECOGNIZED, path=path, name=os.path.basename(path),
            number=str(1 + files.index(path)),
        ),
    )
    stub_dialogs["choice_result"] = "attach_all"
    start_sync(controller, files)
    assert len(stub_dialogs["choices"]) == 1
    assert "Совпадение названий" == str(stub_dialogs["choices"][0][1])
    docs = [d for d in controller.store.documents
            if d["kind"] == storage.KIND_CONTRACT]
    assert len(docs) == 3, "все привязаны к найденной карточке"


# --------------------------------------------------------------------------- #
#  Регистрация: ИНН, дубликаты, unknown, пакетная запись (п. 3.3, 5.2)
# --------------------------------------------------------------------------- #
def test_register_attaches_by_inn_without_dialog(controller, tmp_path,
                                                 monkeypatch, stub_dialogs):
    """Файл с распознанным ИНН привязывается к существующей карточке."""
    party = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Ромашка", F.INN: "7701234567"},
        save=False,
    )
    files = make_files(tmp_path, 1)
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(RECOGNIZED, path=path, name=os.path.basename(path)),
    )
    start_sync(controller, files)
    docs = controller.store.documents_for(party["id"])
    assert len(docs) == 1
    assert docs[0]["extra"]["imported"] is True
    assert docs[0]["extra"]["source_file"], "исходный файл записан в extra"
    assert stub_dialogs["choices"] == [], "диалогов не требуется"
    assert stub_dialogs["confirms"] == []


def test_duplicate_document_not_added(controller, tmp_path, monkeypatch,
                                      stub_dialogs):
    """Повторный импорт того же документа — дубликат, копия в unknown."""
    party = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Ромашка", F.INN: "7701234567"},
        save=False,
    )
    controller.store.add_document(
        storage.KIND_CONTRACT, party["id"], "5", "12 января", data={}, save=False,
    )
    files = make_files(tmp_path, 1)
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(RECOGNIZED, path=path, name=os.path.basename(path)),
    )
    start_sync(controller, files)
    docs = controller.store.documents_for(party["id"])
    assert len(docs) == 1, "дубликат повторно не добавляется"
    # Копия перенесена в unknown; исходник пользователя не тронут (п. 3.5).
    unknown = os.path.join(controller.import_flow.documents_dir_for_import(),
                           "unknown")
    assert os.path.isdir(unknown)
    assert os.listdir(unknown), "копия дубликата лежит в unknown"
    assert os.path.exists(files[0]), "исходный файл пользователя остаётся на месте"


def test_unrecognized_file_moves_to_unknown(controller, tmp_path, monkeypatch,
                                            stub_dialogs):
    """Нераспознанный файл не попадает в реестр (п. 3.5).

    Файл не копировался (разбор не удался до копирования) — в unknown
    переносить нечего, исходник пользователя остаётся на месте.
    """
    files = make_files(tmp_path, 1, prefix="битый ")
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: {
            "ok": False, "recognized": False, "kind": storage.KIND_UNKNOWN,
            "number": "", "date": "", "party": None, "text": "",
            "path": path, "name": os.path.basename(path),
            "error": "Файл повреждён", "warning": "", "copied": "",
        },
    )
    start_sync(controller, files)
    assert controller.store.documents == []
    assert controller.store.active_counterparties() == []
    assert os.path.exists(files[0]), "исходник не тронут"


def test_batch_writes_store_once(controller, tmp_path, monkeypatch,
                                 stub_dialogs):
    """Пакетная запись: изменения пакета пишутся ОДНИМ save() (п. 5.2)."""
    files = make_files(tmp_path, 5)
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(
            RECOGNIZED, path=path, name=os.path.basename(path),
            number=str(int(os.path.splitext(os.path.basename(path))[0].split("№")[1])),
        ),
    )
    save_calls: list = []
    real_save = controller.store.save
    monkeypatch.setattr(
        controller.store, "save",
        lambda: (save_calls.append(1), real_save())[1],
    )
    start_sync(controller, files)
    assert save_calls == [1], "ровно одна запись реестра на пакет"
    assert len(controller.store.documents) == 5


def test_cancellation_counts_cancelled(controller, tmp_path, monkeypatch,
                                       stub_dialogs, patch_report):
    """Отмена: необработанная часть пакета честно посчитана (п. 5.3)."""
    files = make_files(tmp_path, 5)
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(RECOGNIZED, path=path, name=os.path.basename(path)),
    )

    class NonStartFactory:
        """Фабрика, не запускающая воркер: он стартует вручную после отмены."""

        def __init__(self):
            self.target = None
            self.args = ()

        def __call__(self, target=None, args=(), daemon=None, **_kw):
            self.target, self.args = target, args
            return self

        def start(self) -> None:
            pass  # воркер запускается вручную после отмены

    flow = controller.import_flow
    factory = NonStartFactory()
    flow.start(files, thread_factory=factory)
    flow.cancel()  # отмена до старта воркера — весь пакет пропускается
    factory.target(*factory.args)  # воркер видит флаг отмены
    flow.tick()  # 'done' → finish: cancelled_count = 5
    assert patch_report, "отчёт показан"
    _entries, summary = patch_report[0]
    assert "Отменено пользователем: 5" in summary
    assert controller.store.documents == []


def test_report_lists_per_file_status(controller, tmp_path, monkeypatch,
                                      patch_report, stub_dialogs):
    """Отчёт: по каждому файлу тип, результат и причина (п. 3.5)."""
    files = make_files(tmp_path, 2)
    monkeypatch.setattr(
        import_qt.importer, "analyze_file",
        lambda path: dict(
            RECOGNIZED, path=path, name=os.path.basename(path),
            number=str(1 + files.index(path)),
        ),
    )
    start_sync(controller, files)
    assert patch_report, "отчёт показан"
    entries, summary = patch_report[0]
    assert len(entries) == 2
    assert all(e["status"] == "added" for e in entries)
    assert "Добавлено документов: 2" in summary
    assert "новых контрагентов: 1" in summary


def test_open_registry_from_report_u10(controller, tmp_path, monkeypatch,
                                       patch_report, stub_dialogs):
    """U10: «Открыть реестр» из отчёта открывает обновлённый реестр."""
    flow = controller.import_flow
    controller.select_page("spec")
    flow.open_registry()
    assert controller.current_page == "registry"


def test_report_dialog_builds_lines(qapp, patch_report):
    """Диалог отчёта: строки по файлам с типом и статусом (оболочка U10)."""
    entries = [
        {"name": "Договор №5.docx", "kind": storage.KIND_CONTRACT,
         "status": "added", "detail": "создана новая карточка", "warning": ""},
        {"name": "битый.docx", "kind": storage.KIND_UNKNOWN,
         "status": "skip", "detail": "Файл повреждён", "warning": ""},
    ]
    dialog = import_qt.ImportReportDialog(
        None, entries, "Загрузка завершена.\nДобавлено документов: 1",
        storage.KIND_LABELS, lambda: None,
    )
    text = dialog.details.toPlainText()
    assert "Договор №5.docx — Договор поставки, добавлен" in text
    assert "битый.docx" in text and "Файл повреждён" in text


# --------------------------------------------------------------------------- #
#  Точки входа: кнопка реестра и перетаскивание файлов (п. 4.5)
# --------------------------------------------------------------------------- #
def test_import_button_opens_file_dialog(controller, monkeypatch):
    """Кнопка «Добавить документы…» открывает диалог выбора файлов."""
    calls: list = []
    monkeypatch.setattr(
        import_qt.QFileDialog, "getOpenFileNames",
        lambda *a, **kw: (["/tmp/a.docx", "/tmp/b.xlsx"], "фильтр")
        if not calls.append(1) else ([], ""),
    )
    started: list = []
    monkeypatch.setattr(controller.import_flow, "start",
                        lambda files, thread_factory=None: started.append(files))
    controller.import_documents_dialog()
    assert len(calls) == 1
    assert started == [["/tmp/a.docx", "/tmp/b.xlsx"]]


def test_import_button_cancelled_no_start(controller, monkeypatch):
    monkeypatch.setattr(
        import_qt.QFileDialog, "getOpenFileNames", lambda *a, **kw: ([], ""),
    )
    started: list = []
    monkeypatch.setattr(controller.import_flow, "start",
                        lambda files, thread_factory=None: started.append(files))
    controller.import_documents_dialog()
    assert started == []


def test_drop_files_start_import(controller, monkeypatch):
    """Перетащенные файлы уходят в пакетный импорт (п. 4.5)."""
    started: list = []
    monkeypatch.setattr(controller.import_flow, "start",
                        lambda files, thread_factory=None: started.append(files))
    controller.import_drop_files(["/tmp/a.docx"])
    assert started == [["/tmp/a.docx"]]


def test_drop_files_blocked_while_running(controller, monkeypatch):
    controller.import_flow.running = True
    controller.import_drop_files(["/tmp/a.docx"])
    assert "Импорт уже выполняется" in controller.toast_label.text()


def test_registry_header_has_import_button(controller):
    """Кнопка импорта в шапке реестра (точка входа Tk-паритета)."""
    from PySide6.QtWidgets import QPushButton

    buttons = [
        b.text() for b in controller.pages["registry"].findChildren(QPushButton)
    ]
    assert "Добавить документы…" in buttons
