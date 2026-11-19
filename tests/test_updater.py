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

    def _persist_settings(self):
        self.calls.append("persist")
        return self.persist_ok


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


def test_finish_update_dialog_and_open(app, ui_stubs):
    release = {"tag": "v2.16.0", "version_str": "2.16.0",
               "url": "https://github.com/sutsko89/DocHub/releases/tag/v2.16.0"}
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert app._update_flow.busy is False
    assert len(ui_stubs["confirms"]) == 1
    # Времени отметка + настройки сохранены + статус обновлён.
    assert updater.SETTINGS_KEY_TS in app._settings
    assert "persist" in app.calls
    assert "2.16.0" in ui_stubs["confirms"][0][1]
    assert ui_stubs["opened"] == [release["url"]]


def test_finish_update_declined_keeps_page(app, ui_stubs):
    ui_stubs["confirm_result"] = False
    release = {"tag": "v2.16.0", "version_str": "2.16.0", "url": "https://x"}
    app._update_flow._finish(False, {"status": "update", "message": "",
                                     "release": release})
    assert ui_stubs["opened"] == []


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
