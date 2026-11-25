r"""Тесты M6-c — стартовые задачи на Qt (START-1), TRASH-WARN и сверка.

Qt-эквиваленты test_startup_controller / test_trash_warn / сверочной части
test_registry_controller (Tk-контракты сохраняются дословно):

* воркер START-1: read-only работа в фоне (очистки .docx.tmp, подсчёт
  просроченной корзины, сверка по bulk-карте source_file — LAZY-1, сбор
  сроков по снимку списка); результаты применяются ТОЛЬКО в главном
  потоке (tick из QTimer); флаг остановки проверяется между шагами;
* TRASH-WARN: сводка «будет удалено N записей» в журнале ДО удаления,
  итог — в статусе и немодальном уведомлении; пустой план — тишина
  (критерий: корзина не удаляет записи молча);
* сверка (DEF-3): список сирот попадает в баннер реестра с кнопкой
  «Разобрать»; диалог «Восстановить в реестр» создаёт записи (тип и
  контрагент — по имени файла/папки, D1 — честные счётчики при сбое);
  «В корзину» переносит файлы в служебную _trash;
* сроки доп. соглашений (2.18.0): диалог раз в сутки, отметка пишется
  только при показе; «Открыть реестр» открывает реестр;
* отзывчивость интерфейса (критерий M6): во время сверки таймеры UI
  продолжают срабатывать — рабочий поток не блокирует главный.
"""

from __future__ import annotations

import os
import sys
import threading
import time
from datetime import datetime, timedelta

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import agreements as agreements_mod
import main_qt
import qt_dialogs
import reconcile
import shell_open
import startup_qt
import storage
from document_builder import F
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QLabel, QPushButton


# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_dialogs(monkeypatch):
    captured = {
        "confirms": [], "alerts": [], "choices": [], "choice_result": None,
    }

    def fake_choice(*args, **kwargs):
        captured["choices"].append(args)
        return captured["choice_result"]

    captured["choice"] = fake_choice
    monkeypatch.setattr(qt_dialogs, "confirm",
                        lambda *a, **k: captured["confirms"].append(a) or True)
    monkeypatch.setattr(qt_dialogs, "alert",
                        lambda *a, **k: captured["alerts"].append(a))
    monkeypatch.setattr(qt_dialogs, "choice", fake_choice)
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs, monkeypatch, tmp_path):
    """Каркас без автозапуска стартового воркера (запуск управляется тестом)."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    monkeypatch.setattr(shell_open, "documents_known_folder", lambda: str(tmp_path))
    # BKUP-1: автокопия в тестах не нужна (storage-тесты покрывают ротацию).
    monkeypatch.setattr(main_qt.AppController, "_create_startup_backup",
                        lambda self: None)
    # Стартовый воркер запускается тестами явно: автозапуск каркаса глушится,
    # но оригинальный run сохраняется для run_and_drain.
    real_run = startup_qt.StartupFlowQt.run
    monkeypatch.setattr(startup_qt.StartupFlowQt, "run",
                        lambda self, include_notice=True: False)
    ctrl = main_qt.AppController(qapp)
    ctrl.startup_flow._real_run = real_run
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl.startup_flow.stop()
    ctrl._stop_timers()
    for flow in (ctrl.spec_flow, ctrl.contract_flow, ctrl.additional_flow):
        if getattr(flow, "preview_timer", None) is not None:
            flow.preview_timer.stop()


def run_and_drain(controller, qapp, include_notice=True, hold=None,
                  entered=None, stop_while_held=False):
    """Запускает воркер (реальный поток) и применяет результаты в UI.

    Возвращает список применённых результатов (kind'ы) в порядке
    поступления — по записи в _apply. ``hold``/``entered`` — управление
    воркером, удерживаемым на шаге (проверка флага остановки): после
    entered.set() ставится stop(), затем hold отпускает воркер.
    """
    flow = controller.startup_flow
    applied: list = []
    real_apply = flow._apply

    def recording_apply(kind, payload):
        applied.append(kind)
        return real_apply(kind, payload)

    flow._apply = recording_apply
    flow.running = False
    flow._queue = None
    assert flow._real_run(flow, include_notice=include_notice) is True
    if entered is not None:
        assert entered.wait(timeout=5), "воркер не дошёл до удерживаемого шага"
        if stop_while_held:
            flow.stop()
        if hold is not None:
            hold.set()
    for _ in range(2000):
        qapp.processEvents()
        if flow._queue is not None:
            flow.tick()  # дренаж; 'done' останавливает таймер опроса
        if not flow.running:
            return applied
    raise AssertionError("стартовый воркер не завершился за отведённое время")


# --------------------------------------------------------------------------- #
#  Воркер START-1: порядок задач и применения в главном потоке
# --------------------------------------------------------------------------- #
def test_startup_worker_applies_in_main_thread(controller, qapp, stub_dialogs):
    """Воркер: temp → trash_plan → orphans → notice → done; применение — UI."""
    applied = run_and_drain(controller, qapp)
    assert applied == ["temp", "trash_plan", "orphans", "notice"]
    assert controller.startup_flow.running is False


def test_startup_stop_flag_interrupts_worker(controller, qapp, stub_dialogs,
                                             monkeypatch):
    """Флаг остановки: остаток задач не выполняется (START-1).

    Воркер удерживается на первом шаге; stop() во время работы — после
    отпускания применён только завершённый шаг, остальные пропущены.
    """
    flow = controller.startup_flow
    entered = threading.Event()
    release = threading.Event()

    def slow_temp():
        entered.set()
        release.wait(timeout=5)
        return 0

    monkeypatch.setattr(flow, "_cleanup_temp_documents", slow_temp)
    applied = run_and_drain(controller, qapp, hold=release, entered=entered,
                            stop_while_held=True)
    assert applied == ["temp"], \
        "после остановки применён только завершённый шаг"


def test_startup_notice_skipped_when_gate_closed(controller, qapp, stub_dialogs):
    """include_notice=False — сбор сроков не выполняется, остальные задачи идут."""
    applied = run_and_drain(controller, qapp, include_notice=False)
    assert applied == ["temp", "trash_plan", "orphans", "notice"], \
        "notice применяется с пустым payload (тишина)"
    assert controller.startup_flow.running is False


# --------------------------------------------------------------------------- #
#  TRASH-WARN: сводка до удаления, итог после (критерий этапа 2)
# --------------------------------------------------------------------------- #
def _seed_expired_trash(controller) -> dict:
    party = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Корзина", F.INN: "5555555555"}, save=False,
    )
    doc = controller.store.add_document(
        storage.KIND_CONTRACT, party["id"], "9", "1 мая", data={}, save=False,
    )
    controller.store.trash_document(doc["id"])
    # deleted_at читается из памяти записи — старим напрямую (Tk-паритет).
    doc["deleted_at"] = (
        datetime.now() - timedelta(days=storage.TRASH_RETENTION_DAYS + 5)
    ).isoformat(timespec="seconds")
    assert controller.store.count_expired_trash() >= 1
    return doc


def test_trash_warn_cleanup_with_summary_and_notice(controller, stub_dialogs,
                                                    monkeypatch):
    """План из воркера → сводка в журнале ДО удаления → итог в статусе."""
    _seed_expired_trash(controller)
    warnings: list = []
    monkeypatch.setattr(
        controller.log, "warning",
        lambda msg, *a: warnings.append(msg % a if a else msg),
    )
    controller.cleanup_expired_trash(planned=1)
    assert any(
        "Автоочистка корзины: будет удалено 1 записей" in w for w in warnings
    ), "сводка ДО удаления записана в журнал (TRASH-WARN)"
    assert controller.store.count_expired_trash() == 0, "просроченные записи удалены"
    assert "Автоочистка корзины: удалено 1 записей" in controller.toast_label.text()


def test_trash_warn_empty_plan_is_silent(controller):
    """Пустой план — никаких действий и сообщений (Tk-контракт)."""
    controller.cleanup_expired_trash(planned=0)
    assert controller.toast_label.text() == ""


def test_trash_warn_counts_plan_without_param(controller):
    """Без параметра план считается на месте (ручные вызовы и тесты)."""
    _seed_expired_trash(controller)
    controller.cleanup_expired_trash()
    assert controller.store.count_expired_trash() == 0


def test_startup_applies_trash_plan_from_worker(controller, qapp, stub_dialogs):
    """trash_plan из воркера: мутация реестра — в главном потоке (TRASH-WARN)."""
    _seed_expired_trash(controller)
    run_and_drain(controller, qapp)
    assert controller.store.count_expired_trash() == 0
    assert "Автоочистка корзины" in controller.toast_label.text()


# --------------------------------------------------------------------------- #
#  Сверка «реестр ↔ папка документов» (DEF-3) и баннер реестра
# --------------------------------------------------------------------------- #
def _make_orphan(controller, tmp_path, name="Договор №777 от мая.docx"):
    """Файл вне реестра в папке контрагента (имя папки → контрагент)."""
    party = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Сверка", F.INN: "4444444444"}, save=False,
    )
    controller.store.save()
    import services

    folder = services.safe_filename(storage.display_name(party), "Контрагент")
    docs_dir = os.path.join(str(tmp_path), "EW.DocHub", folder)
    os.makedirs(docs_dir, exist_ok=True)
    path = os.path.join(docs_dir, name)
    with open(path, "wb") as fh:
        fh.write(b"PK\x03\x04fake")
    return party, path


def test_scan_consistency_from_worker_shows_banner(controller, tmp_path,
                                                   stub_dialogs):
    """Результат сверки из воркера: баннер реестра с кнопкой «Разобрать»."""
    _party, path = _make_orphan(controller, tmp_path)
    controller.registry_flow.scan_consistency(orphans=[path])
    texts = [
        label.text() for label in controller.pages["registry"].findChildren(QLabel)
    ]
    assert any("файлов вне реестра: 1" in t for t in texts), \
        "баннер сирот показан"
    buttons = [
        b.text() for b in controller.pages["registry"].findChildren(QPushButton)
    ]
    assert "Разобрать" in buttons


def test_reconcile_restore_creates_records(controller, qapp, tmp_path,
                                           stub_dialogs):
    """«Восстановить в реестр»: тип и контрагент — по имени файла/папки."""
    party, path = _make_orphan(controller, tmp_path)
    controller.registry_flow.scan_consistency(orphans=[path])
    stub_dialogs["choice_result"] = "restore"
    controller.registry_flow.open_reconcile_dialog()
    # deleteLater баннера применяется очередью событий (отложенное удаление).
    from PySide6.QtCore import QEvent

    qapp.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    qapp.processEvents()
    docs = controller.store.documents_for(party["id"])
    assert len(docs) == 1, "запись создана"
    assert docs[0]["kind"] == storage.KIND_CONTRACT, "тип по имени файла"
    assert docs[0]["number"] == "777", "номер из имени файла"
    assert docs[0]["extra"]["restored"] is True
    assert stub_dialogs["alerts"], "итог восстановления показан"
    message = str(stub_dialogs["alerts"][-1][2])
    assert "Добавлено записей в реестр: 1" in message
    # Баннер исчез: сирот больше нет.
    texts = [
        label.text() for label in controller.pages["registry"].findChildren(QLabel)
    ]
    assert not any("файлов вне реестра" in t for t in texts)


def test_reconcile_restore_failure_honest_message(controller, tmp_path,
                                                  stub_dialogs, monkeypatch):
    """D1: сбой финальной записи — ошибка названа, счётчики честные.

    add_document(save=False) успешен в памяти (1 запись добавлена),
    store.save() не удался — пользователь видит «Восстановлено файлов: 1»
    И текст ошибки записи (Tk-паритет reconcile.restore_orphans).
    """
    _party, path = _make_orphan(controller, tmp_path)
    controller.registry_flow.scan_consistency(orphans=[path])

    def failing_save():
        controller.store.last_error = "Не удалось записать реестр: отказ ОС"
        return False

    monkeypatch.setattr(controller.store, "save", failing_save)
    stub_dialogs["choice_result"] = "restore"
    controller.registry_flow.open_reconcile_dialog()
    alerts = [a for a in stub_dialogs["alerts"]
              if str(a[1]) == "Восстановление не удалось"]
    assert alerts, "честное сообщение о сбое (D1)"
    message = str(alerts[-1][2])
    assert "Не удалось записать реестр" in message, "ошибка записи названа"
    assert "Восстановлено файлов: 1" in message
    assert "не восстановлено: 0" in message


def test_reconcile_move_to_trash(controller, tmp_path, stub_dialogs):
    """«В корзину»: файлы переносятся в служебную папку _trash."""
    party, path = _make_orphan(controller, tmp_path)
    controller.registry_flow.scan_consistency(orphans=[path])
    stub_dialogs["choice_result"] = "trash"
    controller.registry_flow.open_reconcile_dialog()
    trash_dir = os.path.join(str(tmp_path), "EW.DocHub",
                             reconcile.TRASH_DIR_NAME)
    assert os.path.isdir(trash_dir)
    moved = [
        os.path.join(root, name)
        for root, _dirs, names in os.walk(trash_dir) for name in names
    ]
    assert moved, "файл перемещён в _trash"
    assert controller.store.documents_for(party["id"]) == []


def test_reconcile_no_orphans_info_alert(controller, stub_dialogs):
    """Сирот нет — информационное сообщение (Tk-контракт)."""
    controller.registry_flow.open_reconcile_dialog()
    assert any(
        "Все файлы документов учтены в реестре" in str(a[2])
        for a in stub_dialogs["alerts"]
    )


def test_scan_consistency_manual_scan_reads_orphans(controller, tmp_path,
                                                    stub_dialogs):
    """Ручная сверка (без параметра) сканирует папку сама (LAZY-1 bulk)."""
    _party, path = _make_orphan(controller, tmp_path)
    controller.registry_flow.scan_consistency()
    assert controller.registry_flow._orphans == [path]


def test_startup_applies_orphans_from_worker(controller, qapp, tmp_path,
                                             stub_dialogs):
    """orphans из воркера: баннер появляется после применения в UI."""
    _party, path = _make_orphan(controller, tmp_path)
    run_and_drain(controller, qapp)
    assert controller.registry_flow._orphans == [path]
    texts = [
        label.text() for label in controller.pages["registry"].findChildren(QLabel)
    ]
    assert any("файлов вне реестра" in t for t in texts)


# --------------------------------------------------------------------------- #
#  Сроки доп. соглашений (2.18.0): суточный гейт и диалог
# --------------------------------------------------------------------------- #
def _seed_expired_agreement(controller):
    party = controller.store.add_counterparty(
        "ooo", {F.SHORT_NAME: "ООО Срок", F.INN: "3333333333"}, save=False,
    )
    contract = controller.store.add_document(
        storage.KIND_CONTRACT, party["id"], "10", "1 января",
        data={F.SHORT_NAME: "ООО Срок"}, save=False,
    )
    controller.store.add_document(
        storage.KIND_ADDITIONAL, party["id"], "2", "1 февраля",
        data={}, save=False,
        extra={
            "contract_id": contract["id"],
            "valid_until": "01.10.2026",
            "valid_until_iso": "2026-10-01",
        },
    )
    controller.store.save()


def test_expiry_notice_dialog_and_registry_transition(controller, qapp,
                                                      stub_dialogs):
    """Уведомление о сроках: диалог, отметка в настройках, переход в реестр."""
    _seed_expired_agreement(controller)
    notice = agreements_mod.notice_text(controller.store)
    assert notice, "сроки собираются из реестра"
    controller.select_page("spec")
    stub_dialogs["choice_result"] = "registry"
    controller.apply_expiry_notice(notice)
    assert stub_dialogs["choices"], "диалог сроков показан"
    assert controller.current_page == "registry", \
        "«Открыть реестр» открывает реестр"
    # Отметка «показано» записана в настройки (суточный гейт).
    saved = controller.settings.get(startup_qt.EXPIRY_NOTICE_SETTINGS_KEY)
    assert isinstance(saved, float)
    assert not controller._expiry_notice_due(), "гейт закрыт на сутки"


def test_startup_notice_reaches_dialog(controller, qapp, stub_dialogs):
    """notice из воркера применяется: диалог сроков от каркаса."""
    _seed_expired_agreement(controller)
    stub_dialogs["choice_result"] = "later"
    run_and_drain(controller, qapp, include_notice=True)
    assert any(
        "Сроки действия доп. соглашений" == str(c[1])
        for c in stub_dialogs["choices"]
    )
    assert startup_qt.EXPIRY_NOTICE_SETTINGS_KEY in controller.settings


def test_expiry_notice_silent_when_nothing_to_warn(controller):
    """Тишина, когда тревожить не о чем; отметка не пишется (Tk-контракт)."""
    controller.apply_expiry_notice(None)
    assert startup_qt.EXPIRY_NOTICE_SETTINGS_KEY not in controller.settings


def test_expiry_notice_gate_after_recent_show(controller):
    """Недавняя отметка — уведомление не показывается (include_notice=False)."""
    controller.settings[startup_qt.EXPIRY_NOTICE_SETTINGS_KEY] = time.time()
    assert controller._expiry_notice_due() is False


# --------------------------------------------------------------------------- #
#  LAZY-1: экспорт после ленивой загрузки возвращает полный текст
# --------------------------------------------------------------------------- #
def test_export_after_lazy_load_returns_full_text(
    qapp, appdata, stub_dialogs, monkeypatch, tmp_path,
):
    """Критерий M6/LAZY-1: экспорт из реестра по ленивой записи полный.

    База пишется до старта каркаса; каркас загружает её с ленивыми телами
    («text»/«data»/«extra» не в памяти) — экспорт договора собирает
    печатную форму из дочитанного тела (store.document_body).
    """
    import lifecycle_qt

    from PySide6.QtWidgets import QFileDialog

    seed = storage.DataStore()
    seed.add_counterparty(
        "ooo",
        {F.SHORT_NAME: "ООО Ленивая сверка", F.INN: "1234567890"},
        save=False,
    )
    seed.add_document(
        storage.KIND_CONTRACT, seed.counterparties[0]["id"], "55", "1 июня",
        data={F.SHORT_NAME: "ООО ЛЕНИВЫЙ-ТЕКСТ-ЭКСПОРТА",
              F.INN: "1234567890"},
        extra={"day": "1", "month": "июня", "is_ip": False},
        text="", save=True,
    )
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    monkeypatch.setattr(shell_open, "documents_known_folder", lambda: str(tmp_path))
    monkeypatch.setattr(main_qt.AppController, "_create_startup_backup",
                        lambda self: None)
    monkeypatch.setattr(startup_qt.StartupFlowQt, "run",
                        lambda self, include_notice=True: False)
    ctrl = main_qt.AppController(qapp)
    try:
        qapp.processEvents()
        rec = ctrl.store.documents[0]
        assert "data" not in rec, "тело ленивое после загрузки каркаса"
        target = str(tmp_path / "export.docx")
        monkeypatch.setattr(
            lifecycle_qt.QFileDialog, "getSaveFileName",
            lambda *a, **kw: (target, "фильтр"),
        )
        ctrl.lifecycle_flow.export_document(rec)
        assert os.path.exists(target), "файл экспорта записан"
        # Читаем файл напрямую через python-docx.
        import docx as docx_module

        document = docx_module.Document(target)
        text = "\n".join(p.text for p in document.paragraphs)
        assert "ООО ЛЕНИВЫЙ-ТЕКСТ-ЭКСПОРТА" in text, \
            "ленивое тело дочитано — печатная форма полная"
    finally:
        ctrl.startup_flow.stop()
        ctrl._stop_timers()
        for flow in (ctrl.spec_flow, ctrl.contract_flow, ctrl.additional_flow):
            if getattr(flow, "preview_timer", None) is not None:
                flow.preview_timer.stop()


# --------------------------------------------------------------------------- #
#  Отзывчивость интерфейса: пинг UI во время сверки (критерий M6)
# --------------------------------------------------------------------------- #
def test_ui_ping_during_reconcile_scan(controller, qapp, tmp_path,
                                       stub_dialogs, monkeypatch):
    """Сверка в рабочем потоке не блокирует главный: таймеры пингуют.

    Критерий плана: «операции реестра не прерывают отклик интерфейса».
    Воркер удерживается на сверке ~350 мс; QTimer на 20 мс продолжает
    тикать, пока свёрка идёт в фоне.
    """
    entered = threading.Event()
    release = threading.Event()

    def slow_scan():
        entered.set()
        release.wait(timeout=5)
        return []

    flow = controller.startup_flow
    monkeypatch.setattr(flow, "_find_orphans", slow_scan)
    flow.running = False
    flow._queue = None
    assert flow._real_run(flow, include_notice=False) is True
    assert entered.wait(timeout=5), "воркер дошёл до сверки"

    pings: list = []
    ping_timer = QTimer(controller.window)
    ping_timer.setInterval(20)
    ping_timer.timeout.connect(lambda: pings.append(len(pings)))
    ping_timer.start()

    deadline = time.monotonic() + 0.35
    while time.monotonic() < deadline:
        qapp.processEvents()
    pings_during = len(pings)
    release.set()
    while flow.running:
        qapp.processEvents()
        if flow._queue is not None:
            flow.tick()
    ping_timer.stop()
    qapp.processEvents()

    assert pings_during >= 5, (
        f"UI не блокируется: таймер сработал {pings_during} раз во время сверки"
    )
    assert flow.running is False
