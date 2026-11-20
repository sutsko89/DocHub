"""Юнит-тесты проверки обновлений — updater.py и UpdateController (2.15.0).

Сеть полностью изолирована: HTTP-функции подменяются, реальных запросов
к GitHub тесты не делают (env-глушилка EWDOCHUB_NO_UPDATE_CHECK ставится
в conftest). Проверяются контракты:

  * версии: разбор «v2.15.0», «2.15», мусора; строгое сравнение с
    выравниванием длины («2.15» == «2.15.0»);
  * троттлинг автопроверки: пустые настройки — пора, свежая отметка — нет;
  * env-глушилка приоритетнее всего: check_enabled() → False;
  * разбор ответа GitHub: обычный релиз, draft → None, без тега → None,
    поиск ``*-setup.exe`` в ассетах;
  * check_for_update: «доступна новая», «версия последняя», сеть → понятный
    текст ошибки (HTTP 403, URLError, таймаут), без исключений наружу;
  * UpdateController: guard busy, три рубильника maybe_auto_check,
    обработка результата _finish (диалог/тишина/ошибка) с подменой
    ui.alert / ui.confirm / webbrowser — без Tk, duck-typed приложение.
"""
from __future__ import annotations

import logging
import queue
import threading
import urllib.error
from types import SimpleNamespace

import pytest

import updater
import updater_controller as ucc


# --------------------------------------------------------------------------- #
#  Версии
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("2.15.0", (2, 15, 0)),
        ("v2.15.0", (2, 15, 0)),
        ("2.15", (2, 15)),
        ("  v3.0.1  ", (3, 0, 1)),
        ("2.15.0-beta", (2, 15, 0)),
        ("", None),
        ("unknown", None),
        ("v..", None),
    ],
)
def test_parse_version(text, expected):
    assert updater.parse_version(text) == expected


@pytest.mark.parametrize(
    ("candidate", "current", "expected"),
    [
        ("2.16.0", "2.15.0", True),
        ("2.15.1", "2.15.0", True),
        ("3.0.0", "2.15.0", True),
        ("2.15.0", "2.15.0", False),
        ("2.14.9", "2.15.0", False),
        ("2.15", "2.15.0", False),      # выравнивание: 2.15 == 2.15.0
        ("v2.16", "2.15.0", True),
        ("garbage", "2.15.0", False),   # неразбираемое — не обновление
        ("2.16.0", "garbage", False),
    ],
)
def test_is_newer(candidate, current, expected):
    assert updater.is_newer(candidate, current) == expected


# --------------------------------------------------------------------------- #
#  Троттлинг и глушилка
# --------------------------------------------------------------------------- #
def test_auto_check_due_empty_settings():
    assert updater.auto_check_due({}) is True


def test_auto_check_due_fresh_timestamp():
    settings = {updater.SETTINGS_KEY_TS: 1_000_000.0}
    assert updater.auto_check_due(settings, now=1_000_000.0 + 60) is False


def test_auto_check_due_stale_timestamp():
    settings = {updater.SETTINGS_KEY_TS: 1_000_000.0}
    assert updater.auto_check_due(
        settings, now=1_000_000.0 + updater.AUTO_CHECK_INTERVAL + 1) is True


def test_auto_check_due_broken_timestamp_value():
    # Битое значение (строка) не роняет проверку — трактуется как «не было».
    assert updater.auto_check_due({updater.SETTINGS_KEY_TS: "брак"}) is True


def test_mark_checked_writes_timestamp():
    settings: dict = {}
    updater.mark_checked(settings)
    assert isinstance(settings[updater.SETTINGS_KEY_TS], float)


def test_auto_check_enabled_default_on():
    assert updater.auto_check_enabled({}) is True
    assert updater.auto_check_enabled({updater.SETTINGS_KEY_AUTO: False}) is False


def test_check_enabled_respects_env(monkeypatch):
    monkeypatch.setenv(updater.DISABLE_ENV, "1")
    assert updater.check_enabled() is False
    monkeypatch.delenv(updater.DISABLE_ENV)
    assert updater.check_enabled() is True


# --------------------------------------------------------------------------- #
#  Разбор ответа GitHub
# --------------------------------------------------------------------------- #
def _release_payload(**overrides):
    payload = {
        "tag_name": "v2.16.0",
        "name": "EW.DocHub 2.16.0",
        "html_url": "https://github.com/sutsko89/DocHub/releases/tag/v2.16.0",
        "body": "Что нового",
        "assets": [
            {"name": "EW.DocHub-2.16.0-setup.exe",
             "browser_download_url": "https://github.com/sutsko89/DocHub/releases/download/v2.16.0/EW.DocHub-2.16.0-setup.exe"},
            {"name": "SHA256SUMS.txt",
             "browser_download_url": "https://example.com/SHA256SUMS.txt"},
        ],
    }
    payload.update(overrides)
    return payload


def test_parse_release_payload_full():
    release = updater.parse_release_payload(_release_payload())
    assert release["tag"] == "v2.16.0"
    assert release["version"] == (2, 16, 0)
    assert release["version_str"] == "2.16.0"
    assert release["name"] == "EW.DocHub 2.16.0"
    assert release["url"].endswith("/tag/v2.16.0")
    assert release["installer_url"].endswith("-setup.exe")
    assert release["notes"] == "Что нового"


def test_parse_release_payload_draft_is_skipped():
    assert updater.parse_release_payload(_release_payload(draft=True)) is None


def test_parse_release_payload_no_tag_is_none():
    assert updater.parse_release_payload({"tag_name": ""}) is None
    assert updater.parse_release_payload(None) is None
    assert updater.parse_release_payload("строка") is None


def test_parse_release_payload_without_assets():
    release = updater.parse_release_payload(
        _release_payload(assets=[], html_url=""))
    assert release["installer_url"] == ""
    assert release["url"] == updater.PAGE_URL  # fallback на страницу релизов


# --------------------------------------------------------------------------- #
#  check_for_update: сеть подменяется целиком
# --------------------------------------------------------------------------- #
def test_check_for_update_found(monkeypatch):
    monkeypatch.setattr(
        updater, "fetch_latest_release",
        lambda timeout=updater.TIMEOUT: updater.parse_release_payload(
            _release_payload()))
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "update"
    assert "2.16.0" in result["message"]
    assert result["release"]["tag"] == "v2.16.0"


def test_check_for_update_latest(monkeypatch):
    monkeypatch.setattr(
        updater, "fetch_latest_release",
        lambda timeout=updater.TIMEOUT: updater.parse_release_payload(
            _release_payload(tag_name="v2.15.0")))
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "latest"


def test_check_for_update_http_error_is_human(monkeypatch):
    def boom(timeout=updater.TIMEOUT):
        raise urllib.error.HTTPError(
            updater.RELEASES_API_URL, 403, "rate limit", {}, None)
    monkeypatch.setattr(updater, "fetch_latest_release", boom)
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "error"
    assert "частот" in result["message"] or "ограничил" in result["message"]


def test_check_for_update_network_error_is_human(monkeypatch):
    def boom(timeout=updater.TIMEOUT):
        raise urllib.error.URLError("connection refused")
    monkeypatch.setattr(updater, "fetch_latest_release", boom)
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "error"
    assert "соединени" in result["message"].lower()


def test_check_for_update_timeout_is_human(monkeypatch):
    def boom(timeout=updater.TIMEOUT):
        raise TimeoutError()
    monkeypatch.setattr(updater, "fetch_latest_release", boom)
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "error"
    assert "не отвечает" in result["message"]


def test_check_for_update_unparseable_response(monkeypatch):
    monkeypatch.setattr(updater, "fetch_latest_release", lambda timeout=0: None)
    result = updater.check_for_update("2.15.0")
    assert result["status"] == "error"


def test_human_error_404_mentions_missing_releases():
    message = updater.human_error(
        urllib.error.HTTPError("x", 404, "not found", {}, None))
    assert "не опубликованы" in message


# --------------------------------------------------------------------------- #
#  UpdateController: duck-typed приложение без Tk
# --------------------------------------------------------------------------- #
class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeButton:
    def __init__(self):
        self.configs: list[dict] = []

    def configure(self, **kwargs):
        self.configs.append(kwargs)


class FakeApp:
    def __init__(self):
        self.root = SimpleNamespace(after=lambda *_: None)
        self._settings: dict = {}
        self._update_status_var = FakeVar("")
        self._update_check_button = FakeButton()
        self.calls: list = []
        self.persist_ok = True
        self.choice_result = None  # 2.18.0: ответ диалога обновления

    def _persist_settings(self):
        self.calls.append("persist")
        return self.persist_ok

    def _choice_dialog(self, title, message, options):
        self.calls.append(("choice", title, message, tuple(o[0] for o in options)))
        return self.choice_result

    def close(self):
        self.calls.append("close")


@pytest.fixture()
def app():
    a = FakeApp()
    a._update_flow = ucc.UpdateController(a)
    return a


@pytest.fixture()
def ui_stubs(monkeypatch):
    state = {"alerts": [], "confirms": [], "confirm_result": True,
             "opened": []}

    def fake_alert(root, title, message, kind="info", ok_text="ОК"):
        state["alerts"].append((title, message, kind))

    def fake_confirm(root, title, message, **kwargs):
        state["confirms"].append((title, message, kwargs))
        return state["confirm_result"]

    def fake_open(url):
        state["opened"].append(url)
        return True

    monkeypatch.setattr(ucc.ui, "alert", fake_alert)
    monkeypatch.setattr(ucc.ui, "confirm", fake_confirm)
    monkeypatch.setattr(ucc.webbrowser, "open", fake_open)
    return state


def test_maybe_auto_check_respects_env(monkeypatch, app):
    monkeypatch.setenv(updater.DISABLE_ENV, "1")
    app._update_flow.maybe_auto_check()
    assert app._update_flow.busy is False
    assert app._update_flow.result_queue is None


def test_maybe_auto_check_respects_user_setting(monkeypatch, app):
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    app._settings[updater.SETTINGS_KEY_AUTO] = False
    app._update_flow.maybe_auto_check()
    assert app._update_flow.busy is False


def test_maybe_auto_check_respects_throttle(monkeypatch, app):
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    app._settings[updater.SETTINGS_KEY_TS] = __import__("time").time()
    app._update_flow.maybe_auto_check()
    assert app._update_flow.busy is False


def test_maybe_auto_check_starts_when_due(monkeypatch, app):
    # Полный цикл со «синхронным потоком» и очередью, опрос вручную.
    monkeypatch.delenv(updater.DISABLE_ENV, raising=False)
    started: list = []

    class SyncThread:
        def __init__(self, target, args=(), daemon=False):
            self._target, self._args = target, args

        def start(self):
            started.append(True)
            self._target(*self._args)

    monkeypatch.setattr(ucc.threading, "Thread", SyncThread)
    monkeypatch.setattr(
        updater, "check_for_update",
        lambda current, timeout=updater.TIMEOUT: {
            "status": "latest", "message": "ok", "release": None})
    app._update_flow.maybe_auto_check()
    assert started == [True]
    # Результат в очереди; _finish ещё не выполнялся (Tk не опрашивал).
    assert app._update_flow.busy is True
    auto, result = app._update_flow.result_queue.get_nowait()
    assert auto is True and result["status"] == "latest"


def test_check_now_guard_while_busy(app):
    app._update_flow.busy = True
    app._update_flow.check_now()
    # Guard: новая очередь не создаётся, поток не запускается.
    assert app._update_flow.result_queue is None


def test_finish_update_dialog_offers_install_and_download(app, ui_stubs, monkeypatch):
    release = {"tag": "v2.16.0", "version_str": "2.16.0",
               "url": "https://github.com/sutsko89/DocHub/releases/tag/v2.16.0",
               "installer_url": "https://github.com/x/EW.DocHub-2.16.0-setup.exe"}
    started = []
    monkeypatch.setattr(app._update_flow, "start_download",
                        lambda rel: started.append(rel))
    app.choice_result = "install"
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert app._update_flow.busy is False
    assert started == [release], "«Обновить сейчас» запускает скачивание"
    assert ui_stubs["opened"] == []
    # Времени отметка + настройки сохранены + статус обновлён.
    assert updater.SETTINGS_KEY_TS in app._settings
    assert "persist" in app.calls
    choices = [call for call in app.calls if call[0] == "choice"]
    assert choices, "диалог предлагается пользователю"
    assert "install" in choices[0][3], "в диалоге есть «Обновить сейчас»"


def test_finish_update_dialog_page_choice_opens_browser(app, ui_stubs):
    release = {"tag": "v2.16.0", "version_str": "2.16.0",
               "url": "https://github.com/sutsko89/DocHub/releases/tag/v2.16.0",
               "installer_url": "https://github.com/x/EW.DocHub-2.16.0-setup.exe"}
    app.choice_result = "page"
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert ui_stubs["opened"] == [release["url"]]


def test_finish_update_dialog_later_keeps_page(app, ui_stubs):
    release = {"tag": "v2.16.0", "version_str": "2.16.0", "url": "https://x",
               "installer_url": "https://github.com/x/EW.DocHub-2.16.0-setup.exe"}
    app.choice_result = "later"
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert ui_stubs["opened"] == []


def test_finish_update_without_installer_offers_page_directly(app, ui_stubs):
    release = {"tag": "v2.16.0", "version_str": "2.16.0", "url": "https://x"}
    app.choice_result = None  # окно закрыли крестом
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert ui_stubs["opened"] == [], "закрытие диалога ничего не открывает"
    options = app.calls[0][3]
    assert "install" not in options, "без установщика «Обновить сейчас» не предлагается"


def test_finish_auto_stays_silent_when_latest(app, ui_stubs):
    app._update_flow._finish(True, {"status": "latest", "message": "ok",
                                    "release": None})
    assert ui_stubs["alerts"] == []
    assert ui_stubs["confirms"] == []
    assert updater.SETTINGS_KEY_TS in app._settings


def test_finish_auto_stays_silent_on_error(app, ui_stubs):
    app._update_flow._finish(True, {"status": "error", "message": "нет сети",
                                    "release": None})
    assert ui_stubs["alerts"] == []
    assert updater.SETTINGS_KEY_TS in app._settings


def test_finish_manual_latest_shows_success(app, ui_stubs):
    app._update_flow._finish(False, {"status": "latest", "message": "ok",
                                     "release": None})
    assert len(ui_stubs["alerts"]) == 1
    title, message, kind = ui_stubs["alerts"][0]
    assert "Обновлений нет" in title and kind == "success"


def test_finish_manual_error_shows_message(app, ui_stubs):
    app._update_flow._finish(False, {"status": "error", "message": "нет сети",
                                     "release": None})
    title, message, kind = ui_stubs["alerts"][0]
    assert "Не удалось проверить" in title and kind == "error"
    assert "нет сети" in message


def test_finish_blocks_button_during_check(app):
    flow = app._update_flow
    flow._set_busy_ui(True)
    assert app._update_check_button.configs[-1]["state"] == "disabled"
    flow._set_busy_ui(False)
    assert app._update_check_button.configs[-1]["state"] == "normal"
    assert app._update_check_button.configs[-1]["text"] == "Проверить обновления"


def test_set_busy_ui_without_button_is_safe():
    # Карточка «Обновления» ещё не построена — ничего не должно падать.
    bare = SimpleNamespace()
    flow = ucc.UpdateController(bare)
    flow._set_busy_ui(True)  # не бросает


def test_result_tick_silent_when_window_closes(app):
    # Окно закрывается: after() бросает RuntimeError — цикл просто завершается.
    def broken_after(*_args):
        raise RuntimeError("main loop is closed")

    app.root = SimpleNamespace(after=broken_after)
    app._update_flow.result_queue = queue.Queue()  # пустая → queue.Empty
    app._update_flow.result_tick()  # не бросает, очередь живёт до следующего тика


def test_status_text_contains_version(app, monkeypatch):
    import main

    monkeypatch.setattr(main, "APP_VERSION", "9.9.9", raising=False)
    assert "9.9.9" in app._update_flow.status_text()


# --------------------------------------------------------------------------- #
#  Скачивание и целостность (2.18.0)
# --------------------------------------------------------------------------- #
def test_parse_release_payload_extracts_sums_and_installer_name():
    payload = {
        "tag_name": "v2.18.0", "name": "EW.DocHub 2.18.0", "draft": False,
        "html_url": "https://github.com/sutsko89/DocHub/releases/tag/v2.18.0",
        "assets": [
            {"name": "EW.DocHub-2.18.0-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.18.0-setup.exe"},
            {"name": "SHA256SUMS.txt",
             "browser_download_url": "https://github.com/x/SHA256SUMS.txt"},
        ],
    }
    release = updater.parse_release_payload(payload)
    assert release["installer_url"].endswith("-setup.exe")
    assert release["sums_url"].endswith("SHA256SUMS.txt")
    assert release["installer_name"] == "EW.DocHub-2.18.0-setup.exe"


def test_download_file_streams_and_reports_progress(tmp_path):
    import urllib.request

    source = tmp_path / "payload.bin"
    source.write_bytes(b"x" * 5000)
    dest = tmp_path / "out.bin"
    url = "file://" + urllib.request.pathname2url(str(source))
    progress: list = []

    result = updater.download_file(url, str(dest), progress_cb=lambda d, t: progress.append((d, t)))

    assert result == str(dest)
    assert dest.read_bytes() == source.read_bytes()
    assert progress[0] == (0, 5000), "первый отчёт — ноль из известного размера"
    assert progress[-1] == (5000, 5000)


def test_download_file_cancelled_leaves_no_output(tmp_path):
    import threading
    import urllib.request

    source = tmp_path / "payload.bin"
    source.write_bytes(b"y" * 100000)
    dest = tmp_path / "out.bin"
    url = "file://" + urllib.request.pathname2url(str(source))
    cancel = threading.Event()

    def cancel_at(done, total):
        cancel.set()  # отменяемся на первом же отчёте

    try:
        updater.download_file(url, str(dest), progress_cb=cancel_at, cancel_event=cancel)
        cancelled = False
    except updater.DownloadCancelled:
        cancelled = True
    assert cancelled
    assert not dest.exists(), "после отмены целевого файла нет"
    assert not (dest.with_suffix(dest.suffix + ".part")).exists()


def test_match_sums_line_formats():
    sums = "abc123…  no\n"
    digest = "a" * 64
    text = f"{digest}  EW.DocHub-2.18.0-setup.exe\n{ 'b'*64 } *EW.DocHub-old-setup.exe\n"
    assert updater.match_sums_line(text, "EW.DocHub-2.18.0-setup.exe") == digest
    assert updater.match_sums_line(text, "ew.dochub-2.18.0-setup.exe") == digest
    assert updater.match_sums_line(text, "missing.exe") == ""
    assert updater.match_sums_line(text, "") == ""
    assert updater.match_sums_line("badhash  x.exe", "x.exe") == "", "нечестный хеш отброшен"


def test_verify_sha256(tmp_path):
    import hashlib

    data = b"EW.DocHub setup payload"
    path = tmp_path / "setup.exe"
    path.write_bytes(data)
    assert updater.verify_sha256(str(path), hashlib.sha256(data).hexdigest())
    assert not updater.verify_sha256(str(path), "0" * 64)
    assert not updater.verify_sha256(str(path), "")


def test_download_worker_done_and_sha_check(app, tmp_path, monkeypatch):
    import hashlib
    import queue as queue_mod

    release = {
        "installer_name": "EW.DocHub-2.18.0-setup.exe",
        "sums_url": "https://github.com/x/SHA256SUMS.txt",
    }
    payload = b"installer-bytes"
    path = tmp_path / "EW.DocHub-2.18.0-setup.exe"
    path.write_bytes(payload)
    digest = hashlib.sha256(payload).hexdigest()

    monkeypatch.setattr(updater, "download_file", lambda url, dest, **kw: path)
    monkeypatch.setattr(updater, "fetch_text", lambda url: f"{digest}  {release['installer_name']}")
    verified = []
    monkeypatch.setattr(updater, "verify_sha256",
                        lambda p, expected: (verified.append(p), True)[1])
    result_queue = queue_mod.Queue()
    app._update_flow._download_worker(
        release, "https://x/setup.exe", threading.Event(), result_queue)

    items = []
    while not result_queue.empty():
        items.append(result_queue.get_nowait())
    kinds = [item[0] for item in items]
    assert "phase" in kinds and "done" in kinds
    assert items[-1][0] == "done"
    assert verified, "SHA-256 проверяется перед запуском установщика"


def test_download_worker_sha_mismatch_removes_file(app, tmp_path, monkeypatch):
    import queue as queue_mod

    release = {
        "installer_name": "EW.DocHub-2.18.0-setup.exe",
        "sums_url": "https://github.com/x/SHA256SUMS.txt",
    }
    path = tmp_path / "EW.DocHub-2.18.0-setup.exe"
    path.write_bytes(b"tampered")
    monkeypatch.setattr(updater, "download_file", lambda url, dest, **kw: path)
    monkeypatch.setattr(updater, "fetch_text", lambda url: f"{'a' * 64}  {release['installer_name']}")
    removed = []
    monkeypatch.setattr(updater, "verify_sha256", lambda p, expected: False)
    monkeypatch.setattr(ucc.os, "remove", lambda p: removed.append(p))
    result_queue = queue_mod.Queue()
    app._update_flow._download_worker(
        release, "https://x/setup.exe", threading.Event(), result_queue)

    items = []
    while not result_queue.empty():
        items.append(result_queue.get_nowait())
    assert items[-1][0] == "error"
    assert "целостности" in items[-1][1]
    assert removed, "файл с неверной суммой удаляется"


def test_download_worker_without_sums_reports_error(app, tmp_path, monkeypatch):
    import queue as queue_mod

    release = {"installer_name": "setup.exe", "sums_url": ""}
    path = tmp_path / "setup.exe"
    path.write_bytes(b"x")
    monkeypatch.setattr(updater, "download_file", lambda url, dest, **kw: path)
    result_queue = queue_mod.Queue()
    app._update_flow._download_worker(
        release, "https://x/setup.exe", threading.Event(), result_queue)
    items = []
    while not result_queue.empty():
        items.append(result_queue.get_nowait())
    assert items[-1][0] == "error"
    assert "SHA256SUMS" in items[-1][1]


def test_download_worker_cancelled(app, tmp_path, monkeypatch):
    import queue as queue_mod

    release = {"installer_name": "setup.exe", "sums_url": ""}

    def cancelled_download(url, dest, **kwargs):
        raise updater.DownloadCancelled()

    monkeypatch.setattr(updater, "download_file", cancelled_download)
    result_queue = queue_mod.Queue()
    app._update_flow._download_worker(
        release, "https://x/setup.exe", threading.Event(), result_queue)
    assert result_queue.get_nowait()[0] == "cancelled"


def test_run_installer_command_requires_windows(app, tmp_path, monkeypatch):
    path = tmp_path / "setup.exe"
    path.write_bytes(b"MZ")
    monkeypatch.setattr(ucc.sys, "platform", "linux")
    try:
        ucc.run_installer_command(str(path))
        raised = None
    except OSError as exc:
        raised = exc
    assert raised is not None, "вне Windows установщик не запускается"


def test_run_installer_command_missing_file(tmp_path):
    try:
        ucc.run_installer_command(str(tmp_path / "nope.exe"))
        raised = None
    except FileNotFoundError:
        raised = True
    assert raised, "отсутствующий установщик обнаруживается"


def test_run_installer_and_app_close(app, tmp_path, monkeypatch):
    launched = []
    closed = []
    path = tmp_path / "setup.exe"
    path.write_bytes(b"MZ")
    monkeypatch.setattr(ucc, "run_installer_command",
                        lambda p: launched.append(p))
    app.root = SimpleNamespace(after=lambda ms, fn: closed.append((ms, fn)),
                               close=lambda: closed.append("closed"))
    app._update_flow.run_installer(str(path))
    assert launched == [str(path)]
    assert closed, "после запуска установщика приложение закрывается"


def test_confirm_install_declined_keeps_running(app, ui_stubs, monkeypatch):
    ui_stubs["confirm_result"] = False
    launched = []
    monkeypatch.setattr(ucc, "run_installer_command", lambda p: launched.append(p))
    app._update_flow._confirm_install("/tmp/setup.exe")
    assert launched == [], "отказ не запускает установку"


def test_start_download_without_installer_opens_page(app, ui_stubs, monkeypatch):
    release = {"url": "https://github.com/sutsko89/DocHub/releases"}
    app._update_flow.start_download(release)
    assert ui_stubs["opened"] == [release["url"]], "нет установщика — страница релиза"


def test_download_tick_done_shows_confirm(app, ui_stubs, monkeypatch):
    import queue as queue_mod

    launched = []
    monkeypatch.setattr(ucc, "run_installer_command", lambda p: launched.append(p))
    flow = app._update_flow
    flow.download_queue = queue_mod.Queue()
    flow.download_queue.put(("done", "/tmp/EW.DocHub-2.18.0-setup.exe"))
    flow.download_tick()
    assert launched == ["/tmp/EW.DocHub-2.18.0-setup.exe"]
    assert len(ui_stubs["confirms"]) == 1, "перед установкой — подтверждение"


def test_download_tick_error_alerts(app, ui_stubs):
    import queue as queue_mod

    flow = app._update_flow
    flow.download_queue = queue_mod.Queue()
    flow.download_queue.put(("error", "Нет соединения"))
    flow.download_tick()
    assert len(ui_stubs["alerts"]) == 1
    assert "Обновление не удалось" in ui_stubs["alerts"][0][0]
    assert "Нет соединения" in ui_stubs["alerts"][0][1]
