r"""Тесты M7 — обновления на Qt (перенос controller-части test_updater).

Контракты Tk-контроллера updater_controller.UpdateController сохраняются
дословно (2.15.0 / 2.18.0 / 2.19.2):

* три рубильника автопроверки: env-глушилка (EWDOCHUB_NO_UPDATE_CHECK —
  приоритетнее всего), переключатель пользователя, суточный троттлинг;
* guard повторных проверок и блокировка кнопки «Проверить обновления»;
* диалог результата («Обновить сейчас» / «Открыть страницу» / «Позже»),
  честные «версия последняя»/ошибка при ручной проверке, молчание авто;
* рабочий поток скачивания отдаёт progress/phase/done|cancelled|error в
  очередь; SHA-256 по SHA256SUMS.txt; битый файл удаляется;
* перед установкой — деструктивное подтверждение (U7: приложение
  закроется) и снимок реестра (BKUP-1); установка — тихий Inno Setup,
  приложение закрывается штатно.

Специфика Qt-этапа M7 (И.2 плана): ассет релиза выбирается ПО ФЛЕЙВОРУ —
чужой флейвор не скачивается; статус версии — с суффиксом («2.28.2-qt»);
опрос очередей — именованные poll-таймеры каркаса (QTimer), терминальное
событие останавливает таймер.

Сеть не используется: сетевые функции updater подменяются целиком; потоки
исполняются синхронно через подменную фабрику (детерминированный прогон).
"""

from __future__ import annotations

import json
import os
import queue
import sys
import tempfile
import threading
import time
import urllib.error
from hashlib import sha256

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import app_common
import main_qt
import qt_dialogs
import shell_open
import updater
import updater_qt

# --------------------------------------------------------------------------- #
#  Фикстуры
# --------------------------------------------------------------------------- #
class _SyncThread:
    """Фабрика потока: воркер исполняется в start() синхронно."""

    def __call__(self, target=None, args=(), daemon=None, **_kw):
        self._target = target
        self._args = args
        return self

    def start(self) -> None:
        self._target(*self._args)


@pytest.fixture()
def stub_dialogs(monkeypatch):
    captured = {"alerts": [], "confirms": [], "choices": [],
                "confirm_result": True, "choice_result": None}

    def fake_alert(*args, **kwargs):
        captured["alerts"].append(args)

    def fake_confirm(*args, **kwargs):
        # main_qt.confirm передаёт всё позиционно (parent, title, message,
        # ok_text, cancel_text, destructive) — захватываем полный кортеж.
        captured["confirms"].append(args)
        return captured["confirm_result"]

    def fake_choice(*args, **kwargs):
        captured["choices"].append(args)
        return captured["choice_result"]

    monkeypatch.setattr(qt_dialogs, "alert", fake_alert)
    monkeypatch.setattr(qt_dialogs, "confirm", fake_confirm)
    monkeypatch.setattr(qt_dialogs, "choice", fake_choice)
    return captured


@pytest.fixture()
def controller(qapp, appdata, stub_dialogs, monkeypatch, tmp_path):
    """Каркас со страницей настроек; фоновые работы и папки изолированы."""
    monkeypatch.setattr(shell_open, "open_file_system", lambda path: True)
    monkeypatch.setattr(shell_open, "show_in_file_manager", lambda path: True)
    monkeypatch.setattr(main_qt.AppController, "_create_startup_backup",
                        lambda self: None)
    ctrl = main_qt.AppController(qapp)
    ctrl.window.show()
    qapp.processEvents()
    yield ctrl
    ctrl.startup_flow.stop()
    ctrl._stop_timers()


@pytest.fixture()
def flow(controller):
    """UpdateControllerQt с синхронной фабрикой потока и своим флейвором.

    Заменяет каркасный поток (реальный threading.Thread) на
    детерминированный и перецепляет кнопку «Проверить обновления» —
    она была подключена к исходному потоку при построении страницы.
    """
    new_flow = updater_qt.UpdateControllerQt(
        controller, flavor="qt", thread_factory=_SyncThread())
    controller.update_flow = new_flow
    button = getattr(controller, "_update_check_button", None)
    if button is not None:
        try:
            button.clicked.disconnect()
        except (RuntimeError, TypeError):
            pass
        button.clicked.connect(new_flow.check_now)
    return new_flow


def _release(**asset_overrides):
    """Релиз с ассетами обоих флейворов (разобранный parse_release_payload)."""
    assets = [
        {"name": "EW.DocHub-2.29.0-setup.exe",
         "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-setup.exe"},
        {"name": "EW.DocHub-2.29.0-qt-setup.exe",
         "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe"},
        {"name": "SHA256SUMS.txt",
         "browser_download_url": "https://github.com/x/SHA256SUMS.txt"},
    ]
    assets.extend(asset_overrides.get("extra_assets", []))
    for skip in asset_overrides.get("skip", []):
        assets = [a for a in assets if skip.lower() not in a["name"].lower()]
    return updater.parse_release_payload({
        "tag_name": "v2.29.0",
        "html_url": "https://github.com/x/tag/v2.29.0",
        "assets": assets,
    })


def _saved_settings() -> dict:
    """settings.json изолированного профиля (запись через app.save_setting)."""
    with open(app_common.settings_path(), encoding="utf-8") as fh:
        return json.load(fh)


def _drain(q: "queue.Queue") -> list:
    items = []
    while True:
        try:
            items.append(q.get_nowait())
        except queue.Empty:
            return items


# --------------------------------------------------------------------------- #
#  Флейвор сборки (И.2)
# --------------------------------------------------------------------------- #
def test_flavor_resolved_from_scaffold(controller):
    """Флейвор без явного аргумента — UI_FLAVOR каркаса («qt»)."""
    assert controller.update_flow.flavor == main_qt.UI_FLAVOR == "qt"


def test_flavor_explicit_override(controller):
    explicit = updater_qt.UpdateControllerQt(controller, flavor="tk")
    assert explicit.flavor == "tk"


def test_status_text_contains_flavor_suffix(flow, monkeypatch):
    """Статус версии — с суффиксом флейвора («2.28.2-qt»; И.2)."""
    assert app_common.APP_VERSION in flow.status_text()
    assert flow.status_text().endswith(f"-{flow.flavor}")
    monkeypatch.setattr(app_common, "APP_VERSION", "9.9.9", raising=False)
    assert flow.status_text() == "Установлена версия 9.9.9-qt"


# --------------------------------------------------------------------------- #
#  Автопроверка: три рубильника (Tk-контракт maybe_auto_check)
# --------------------------------------------------------------------------- #
def test_maybe_auto_check_respects_env(controller, flow, monkeypatch):
    """Env-глушилка (глушится в conftest/smoke) блокирует даже «пору»."""
    monkeypatch.setenv(updater.DISABLE_ENV, "1")
    flow.maybe_auto_check()
    assert flow.busy is False
    assert flow.result_queue is None
    assert updater_qt.CHECK_TIMER not in controller._timers


def test_maybe_auto_check_respects_user_setting(
        controller, flow, monkeypatch):
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    monkeypatch.setattr(updater, "check_for_update",
                        lambda current, **kw: (_ for _ in ()).throw(
                            AssertionError("сеть не должна вызываться")))
    controller.settings[updater.SETTINGS_KEY_AUTO] = False
    flow.maybe_auto_check()
    assert flow.busy is False and flow.result_queue is None


def test_maybe_auto_check_respects_throttle(controller, flow, monkeypatch):
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    monkeypatch.setattr(updater, "check_for_update",
                        lambda current, **kw: (_ for _ in ()).throw(
                            AssertionError("сеть не должна вызываться")))
    controller.settings[updater.SETTINGS_KEY_TS] = time.time()
    flow.maybe_auto_check()
    assert flow.busy is False and flow.result_queue is None


def test_maybe_auto_check_starts_when_due_and_finishes_silently(
        controller, flow, stub_dialogs, monkeypatch):
    """Пора проверять: воркер стартует; «версия последняя» при авто — тихо.

    Факт проверки фиксируется в settings.json (суточный троттлинг),
    guard снимается, таймер опроса остановлен.
    """
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    monkeypatch.setattr(updater, "check_for_update", lambda current, **kw:
                        {"status": "latest", "message": "последняя",
                         "release": {"version_str": "2.29.0", "tag": "v2.29.0"}})
    flow.maybe_auto_check()
    assert flow.busy is True  # воркер отработал синхронно, очередь полна
    flow.result_tick()
    assert flow.busy is False
    assert flow.result_queue is None
    assert updater_qt.CHECK_TIMER not in controller._timers
    assert stub_dialogs["alerts"] == [], "автопроверка молчит"
    assert updater.SETTINGS_KEY_TS in _saved_settings()


# --------------------------------------------------------------------------- #
#  Ручная проверка: guard, честные ответы, блокировка кнопки
# --------------------------------------------------------------------------- #
def test_check_now_guard_while_busy(flow):
    flow.busy = True
    flow.check_now()
    assert flow.result_queue is None, "guard не даёт второй проверки"


def test_check_now_manual_latest_alerts_success(
        controller, flow, stub_dialogs, monkeypatch):
    monkeypatch.setattr(updater, "check_for_update", lambda current, **kw:
                        {"status": "latest", "message": "последняя",
                         "release": None})
    flow.check_now()
    flow.result_tick()
    assert flow.busy is False
    alerts = stub_dialogs["alerts"]
    assert alerts and alerts[0][1] == "Обновлений нет"
    assert "последнюю версию" in alerts[0][2]


def test_check_now_manual_error_alerts_with_page_hint(
        controller, flow, stub_dialogs, monkeypatch):
    def broken(current, **kw):
        raise urllib.error.URLError("нет сети")

    monkeypatch.setattr(updater, "check_for_update", broken)
    flow.check_now()
    flow.result_tick()
    alerts = stub_dialogs["alerts"]
    assert alerts and alerts[0][1] == "Не удалось проверить обновления"
    assert updater.PAGE_URL in alerts[0][2]


def test_set_busy_ui_blocks_and_restores_button(controller, flow):
    button = controller._update_check_button
    flow._set_busy_ui(True)
    assert not button.isEnabled()
    assert button.text() == "Проверяем…"
    flow._set_busy_ui(False)
    assert button.isEnabled()
    assert button.text() == "Проверить обновления"


def test_set_busy_ui_without_button_is_safe(controller):
    bare = updater_qt.UpdateControllerQt(controller, flavor="qt")
    delattr(controller, "_update_check_button")
    bare._set_busy_ui(True)  # нет кнопки — не исключение
    bare._set_busy_ui(False)


def test_refresh_status_text_updates_label(controller, flow, monkeypatch):
    monkeypatch.setattr(app_common, "APP_VERSION", "3.0.0", raising=False)
    flow._refresh_status_text()
    assert controller._update_status_label.text() \
        == "Установлена версия 3.0.0-qt"


def test_result_tick_without_queue_stops_timer(controller, flow):
    controller.add_poll_timer(
        updater_qt.CHECK_TIMER, 80, flow.result_tick)
    flow.result_tick()  # очередь None — таймер останавливается, без исключений
    assert updater_qt.CHECK_TIMER not in controller._timers


# --------------------------------------------------------------------------- #
#  Диалог результата: обновить / страница / позже
# --------------------------------------------------------------------------- #
def test_update_dialog_offers_install_for_own_flavor(
        controller, flow, stub_dialogs):
    """Релиз с Qt-ассетом: «Обновить сейчас» — главная кнопка диалога."""
    release = _release()
    stub_dialogs["choice_result"] = "later"
    flow._show_update_dialog(release)
    assert stub_dialogs["choices"], "диалог обновления показан"
    _parent, title, message, options = stub_dialogs["choices"][0]
    assert title == "Доступна новая версия EW.DocHub"
    keys = [opt[0] for opt in options]
    assert keys == ["later", "page", "install"]
    primary = [opt for opt in options if opt[2]]
    assert primary and primary[0][0] == "install"
    assert str(app_common.APP_VERSION) in message


def test_update_dialog_without_own_flavor_offers_page_first(
        controller, flow, stub_dialogs):
    """Qt-ассета в релизе нет: страница релиза — главная кнопка.

    Чужой Tk-установщик НЕ предлагается и НЕ скачивается (И.2).
    """
    release = _release(skip="2.29.0-qt")
    stub_dialogs["choice_result"] = "later"
    flow._show_update_dialog(release)
    _parent, title, message, options = stub_dialogs["choices"][0]
    keys = [opt[0] for opt in options]
    assert keys == ["later", "page"]
    primary = [opt for opt in options if opt[2]]
    assert primary and primary[0][0] == "page"
    assert "2.29.0" in message


def test_dialog_page_choice_opens_browser(controller, flow, monkeypatch):
    release = _release()
    opened: list = []
    monkeypatch.setattr(updater_qt.webbrowser, "open",
                        lambda url: opened.append(url))
    flow.open_release_page(release)
    assert opened == [release["url"]]


def test_dialog_later_keeps_page(controller, flow, stub_dialogs):
    release = _release()
    stub_dialogs["choice_result"] = "later"
    flow._show_update_dialog(release)
    assert flow.download_queue is None
    assert stub_dialogs["confirms"] == []


# --------------------------------------------------------------------------- #
#  Скачивание: чужой флейвор не скачивается (критерий M7)
# --------------------------------------------------------------------------- #
def test_start_download_foreign_flavor_not_downloaded(
        controller, flow, monkeypatch):
    """Релиз только с Tk-ассетом: Qt-сборка идёт на страницу, не качает."""
    downloaded: list = []
    opened: list = []
    monkeypatch.setattr(updater, "download_file",
                        lambda *a, **kw: downloaded.append(a))
    monkeypatch.setattr(updater_qt.webbrowser, "open",
                        lambda url: opened.append(url))
    release = _release(skip="2.29.0-qt")
    flow.start_download(release)
    assert downloaded == [], "чужой флейвор не скачивается"
    assert opened == [release["url"]]
    assert flow.download_queue is None


def test_start_download_own_flavor_uses_flavor_asset(
        controller, flow, monkeypatch, stub_dialogs):
    """Свой ассет: скачивается имя СВОЕГО флейвора, окно прогресса открыто."""
    release = _release()
    seen_urls: list = []

    def fake_download(url, dest_path, progress_cb=None, cancel_event=None,
                      timeout=None):
        seen_urls.append(url)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        payload = b"setup-payload"
        if callable(progress_cb):
            progress_cb(0, len(payload))
            progress_cb(len(payload), len(payload))
        with open(dest_path, "wb") as fh:
            fh.write(payload)
        return dest_path

    monkeypatch.setattr(updater, "download_file", fake_download)
    monkeypatch.setattr(updater, "fetch_text", lambda url:
                        f"{sha256(b'setup-payload').hexdigest()}  "
                        "EW.DocHub-2.29.0-qt-setup.exe")
    flow.start_download(release)
    assert seen_urls == [release["installers"]["qt"]["url"]]
    assert flow.download_window is not None
    # Воркер отработал синхронно — очередь полна; осушаем тиком.
    flow.download_tick()
    assert stub_dialogs["confirms"], "done дошёл до подтверждения установки"
    assert flow.download_queue is None
    assert flow.download_window is None
    assert updater_qt.DOWNLOAD_TIMER not in controller._timers


def test_start_download_guard_while_active(controller, flow):
    """Повторный start_download при активном скачивании — no-op."""
    release = _release()
    flow.download_queue = queue.Queue()
    flow.start_download(release)  # не должен перезаписать очередь/окно
    assert flow.download_window is None


# --------------------------------------------------------------------------- #
#  Рабочий поток скачивания (Tk-контракт очереди)
# --------------------------------------------------------------------------- #
def _download_stub(payload: bytes):
    """Стаб updater.download_file: каталог назначения + отчёт о прогрессе —
    контракты реальной функции (она сама создаёт каталог и отчитывается)."""

    def fake_download(url, path, progress_cb=None, cancel_event=None,
                      timeout=None):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if callable(progress_cb):
            progress_cb(0, len(payload))
            progress_cb(len(payload), len(payload))
        with open(path, "wb") as fh:
            fh.write(payload)
        return str(path)

    return fake_download


def test_download_worker_reports_progress_phase_done(tmp_path, monkeypatch):
    monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
    payload = b"setup-payload"
    name = "EW.DocHub-2.29.0-qt-setup.exe"

    monkeypatch.setattr(updater, "download_file", _download_stub(payload))
    monkeypatch.setattr(updater, "fetch_text", lambda url:
                        f"{sha256(payload).hexdigest()}  {name}")
    result_queue: queue.Queue = queue.Queue()
    flow_qt = updater_qt.UpdateControllerQt(None, flavor="qt")
    flow_qt._download_worker(
        {"sums_url": "https://github.com/x/SHA256SUMS.txt"},
        "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe", name,
        threading.Event(), result_queue)
    kinds = [item[0] for item in _drain(result_queue)]
    assert "progress" in kinds and "phase" in kinds
    assert kinds[-1] == "done"
    assert os.path.exists(tmp_path / "EW.DocHub-update" / name)


def test_download_worker_bad_checksum_removes_file(tmp_path, monkeypatch):
    monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
    payload = b"setup-payload"
    name = "EW.DocHub-2.29.0-qt-setup.exe"

    monkeypatch.setattr(updater, "download_file", _download_stub(payload))
    monkeypatch.setattr(updater, "fetch_text", lambda url:
                        f"{sha256(b'other payload').hexdigest()}  {name}")
    result_queue: queue.Queue = queue.Queue()
    flow_qt = updater_qt.UpdateControllerQt(None, flavor="qt")
    flow_qt._download_worker(
        {"sums_url": "https://github.com/x/SHA256SUMS.txt"},
        "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe", name,
        threading.Event(), result_queue)
    items = _drain(result_queue)
    assert items[-1][0] == "error"
    assert "SHA-256" in items[-1][1]
    assert not os.path.exists(tmp_path / "EW.DocHub-update" / name), \
        "битый файл удалён"


def test_download_worker_missing_sums_is_error(tmp_path, monkeypatch):
    monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
    monkeypatch.setattr(updater, "download_file",
                        lambda url, path, **kw: str(path))
    result_queue: queue.Queue = queue.Queue()
    flow_qt = updater_qt.UpdateControllerQt(None, flavor="qt")
    flow_qt._download_worker(
        {"sums_url": ""}, "https://github.com/x/setup.exe",
        "EW.DocHub-2.29.0-qt-setup.exe", threading.Event(), result_queue)
    items = _drain(result_queue)
    assert items[-1][0] == "error"
    assert "контрольную сумму" in items[-1][1]


def test_download_worker_cancelled_is_reported(tmp_path, monkeypatch):
    cancelled = threading.Event()

    def cancelled_download(url, path, progress_cb=None, cancel_event=None,
                           timeout=None):
        cancel_event.set()
        raise updater.DownloadCancelled()

    monkeypatch.setattr(updater, "download_file", cancelled_download)
    result_queue: queue.Queue = queue.Queue()
    flow_qt = updater_qt.UpdateControllerQt(None, flavor="qt")
    flow_qt._download_worker(
        {"sums_url": "https://github.com/x/SHA256SUMS.txt"},
        "https://github.com/x/setup.exe", "EW.DocHub-2.29.0-qt-setup.exe",
        cancelled, result_queue)
    assert _drain(result_queue)[-1][0] == "cancelled"


def test_download_worker_foreign_host_rejected(tmp_path):
    """SEC-1: чужой хост — отказ без сетевого запроса (честная ошибка)."""
    result_queue: queue.Queue = queue.Queue()
    flow_qt = updater_qt.UpdateControllerQt(None, flavor="qt")
    flow_qt._download_worker(
        {"sums_url": "https://github.com/x/SHA256SUMS.txt"},
        "https://example.com/evil.exe", "evil.exe",
        threading.Event(), result_queue)
    items = _drain(result_queue)
    assert items[-1][0] == "error"
    assert "GitHub" in items[-1][1]


# --------------------------------------------------------------------------- #
#  Окно прогресса и отмена
# --------------------------------------------------------------------------- #
def test_progress_window_close_is_cancel(controller, flow, qapp):
    release = _release()
    flow.download_queue = queue.Queue()
    flow.cancel_event = threading.Event()
    flow._build_progress_window(release)
    window = flow.download_window
    assert window.isVisible()
    window.close()
    qapp.processEvents()
    assert flow.cancel_event.is_set(), "закрытие окна = отмена"
    assert not window.isVisible()


def test_cancel_download_sets_event_and_stops_timer(controller, flow):
    flow.download_queue = queue.Queue()
    flow.cancel_event = threading.Event()
    controller.add_poll_timer(
        updater_qt.DOWNLOAD_TIMER, 80, flow.download_tick)
    flow.cancel_download()
    assert flow.cancel_event.is_set()
    assert flow.download_queue is None
    assert updater_qt.DOWNLOAD_TIMER not in controller._timers


def test_progress_window_shows_bytes_and_phase(controller, flow):
    flow._build_progress_window(_release())
    window = flow.download_window
    window.set_progress(1024 * 1024, 20 * 1024 * 1024)
    assert "МБ" in window.status_label.text()
    window.set_progress(512 * 1024, 700 * 1024)
    assert "КБ" in window.status_label.text()
    window.set_status("Проверка целостности файла…")
    assert window.status_label.text() == "Проверка целостности файла…"
    window.set_progress(100, 0)  # без Content-Length — неопределённый режим
    assert window.progress.maximum() == 0
    flow._close_progress_window()
    assert flow.download_window is None


# --------------------------------------------------------------------------- #
#  Установка: подтверждение, снимок реестра, запуск, закрытие приложения
# --------------------------------------------------------------------------- #
def test_install_confirmation_is_destructive(controller, flow, stub_dialogs):
    """U7: программа закроется — случайный Enter недопустим (Tk-контракт)."""
    stub_dialogs["confirm_result"] = False
    flow._confirm_install("/tmp/setup.exe")
    assert stub_dialogs["confirms"], "подтверждение показано"
    args = stub_dialogs["confirms"][0]
    # main_qt.confirm(parent, title, message, ok_text, cancel_text, destructive)
    assert args[2].startswith("Установщик проверен")
    assert args[3] == "Обновить"
    assert args[5] is True, "деструктивное подтверждение (U7)"


def test_run_installer_launches_and_closes_app(
        controller, flow, monkeypatch, stub_dialogs):
    """Подтверждение → снимок реестра → тихий запуск → отложенное закрытие."""
    launched: list = []
    callbacks: list = []
    monkeypatch.setattr(updater, "run_installer_command",
                        lambda path: launched.append(path))
    monkeypatch.setattr(controller, "call_later",
                        lambda delay, cb: callbacks.append((delay, cb)))
    snapshotted: list = []
    import storage as storage_mod
    monkeypatch.setattr(storage_mod, "create_update_snapshot",
                        lambda store, version: snapshotted.append(version))
    flow.run_installer("/tmp/EW.DocHub-2.29.0-qt-setup.exe")
    assert launched == ["/tmp/EW.DocHub-2.29.0-qt-setup.exe"]
    assert snapshotted == [app_common.APP_VERSION], "снимок реестра создан"
    assert callbacks and callbacks[0][0] == 500
    assert callable(callbacks[0][1]), "закрытие приложения отложено"


def test_run_installer_failure_alerts_and_stays_open(
        controller, flow, monkeypatch, stub_dialogs):
    """Отказ запуска установщика — честный алерт с путём, закрытия нет."""
    launched: list = []
    callbacks: list = []

    def broken(path):
        launched.append(path)
        raise OSError("нет прав")

    monkeypatch.setattr(updater, "run_installer_command", broken)
    monkeypatch.setattr(controller, "call_later",
                        lambda delay, cb: callbacks.append((delay, cb)))
    flow.run_installer("/tmp/setup.exe")
    assert launched == ["/tmp/setup.exe"]
    assert callbacks == [], "при отказе приложение не закрывается"
    assert stub_dialogs["alerts"], "алерт с путём показан"
    assert "/tmp/setup.exe" in stub_dialogs["alerts"][0][2]


# --------------------------------------------------------------------------- #
#  Страница «Настройки»: карточка «Обновления» (Г #17)
# --------------------------------------------------------------------------- #
def test_settings_page_update_card(controller):
    """Кнопка, статус версии с суффиксом и чекбокс автопроверки на месте."""
    assert controller._update_check_button.text() == "Проверить обновления"
    assert controller._update_check_button.isEnabled()
    label_text = controller._update_status_label.text()
    assert app_common.APP_VERSION in label_text
    assert label_text.endswith("-qt")
    assert controller.settings_flow._update_auto_checkbox.isChecked() is True


def test_settings_page_auto_check_persisted(controller):
    """Переключатель автопроверки пишется в settings.json немедленно."""
    checkbox = controller.settings_flow._update_auto_checkbox
    checkbox.setChecked(False)
    assert _saved_settings()[updater.SETTINGS_KEY_AUTO] is False
    checkbox.setChecked(True)
    assert _saved_settings()[updater.SETTINGS_KEY_AUTO] is True


def test_settings_button_check_now_full_route(
        controller, flow, stub_dialogs, monkeypatch):
    """Кнопка «Проверить обновления»: guard → проверка → честный ответ."""
    monkeypatch.setattr(updater, "check_for_update", lambda current, **kw:
                        {"status": "latest", "message": "последняя",
                         "release": None})
    controller._update_check_button.click()
    flow.result_tick()
    alerts = stub_dialogs["alerts"]
    assert alerts and alerts[0][1] == "Обновлений нет"
    assert controller._update_check_button.isEnabled()
