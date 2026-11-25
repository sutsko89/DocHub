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
import sys
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
    import app_common

    monkeypatch.setattr(app_common, "APP_VERSION", "9.9.9", raising=False)
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


class _FakeResponse:
    """Duck-typed ответ urlopen: отдаёт payload порциями с заголовками."""

    def __init__(self, payload: bytes = b"", headers=None, huge=False):
        self._payload = payload
        self._pos = 0
        self.headers = headers or {}
        self._huge = huge

    def read(self, size=-1):
        if self._huge:
            return b"z" * (size if size and size > 0 else 4096)
        if self._pos >= len(self._payload):
            return b""
        chunk = self._payload[self._pos:self._pos + size] if size and size > 0 \
            else self._payload[self._pos:]
        self._pos += len(chunk)
        return chunk

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


SETUP_URL = (
    "https://github.com/sutsko89/DocHub/releases/download/"
    "v2.22.0/EW.DocHub-2.22.0-setup.exe"
)


def test_download_file_streams_and_reports_progress(tmp_path, monkeypatch):
    import urllib.request

    payload = b"x" * 5000
    opened: list[str] = []

    def fake_urlopen(request, timeout=None):
        opened.append(request.full_url)
        return _FakeResponse(payload, headers={"Content-Length": "5000"})

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    dest = tmp_path / "out.bin"
    progress: list = []

    result = updater.download_file(
        SETUP_URL, str(dest), progress_cb=lambda d, t: progress.append((d, t)),
    )

    assert result == str(dest)
    assert dest.read_bytes() == payload
    assert opened == [SETUP_URL], "запрос ушёл на allowlist-хост"
    assert progress[0] == (0, 5000), "первый отчёт — ноль из известного размера"
    assert progress[-1] == (5000, 5000)


def test_download_file_cancelled_leaves_no_output(tmp_path, monkeypatch):
    import threading
    import urllib.request

    monkeypatch.setattr(
        urllib.request, "urlopen",
        lambda request, timeout=None: _FakeResponse(b"y" * 100000),
    )
    dest = tmp_path / "out.bin"
    cancel = threading.Event()

    def cancel_at(done, total):
        cancel.set()  # отменяемся на первом же отчёте

    try:
        updater.download_file(
            SETUP_URL, str(dest), progress_cb=cancel_at, cancel_event=cancel,
        )
        cancelled = False
    except updater.DownloadCancelled:
        cancelled = True
    assert cancelled
    assert not dest.exists(), "после отмены целевого файла нет"
    assert not (dest.with_suffix(dest.suffix + ".part")).exists()


# --------------------------------------------------------------------------- #
#  SEC-1 (D11): allowlist хостов и лимит размера
# --------------------------------------------------------------------------- #
def test_is_allowed_download_url_matrix():
    assert updater.is_allowed_download_url(SETUP_URL), "github.com разрешён"
    assert updater.is_allowed_download_url(
        "https://objects.githubusercontent.com/asset/exe"
    ), "хост раздачи ассетов разрешён"
    assert updater.is_allowed_download_url(
        "https://API.GitHub.com/repos/x/releases/latest"
    ), "регистр хоста не важен"
    assert not updater.is_allowed_download_url(
        "https://evil.example.com/setup.exe"
    ), "чужой хост отклонён"
    assert not updater.is_allowed_download_url(
        "http://github.com/setup.exe"
    ), "не-https отклонён"
    assert not updater.is_allowed_download_url(
        "file:///C:/Windows/System32/cmd.exe"
    ), "локальная схема отклонена"
    assert not updater.is_allowed_download_url("")
    assert not updater.is_allowed_download_url("не URL")


def test_download_foreign_host_rejected_without_request(monkeypatch, tmp_path):
    import urllib.request

    def fake_urlopen(request, timeout=None):
        raise AssertionError("сетевой запрос не должен выполняться")

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    with pytest.raises(updater.DownloadNotAllowed):
        updater.download_file(
            "https://evil.example.com/EW.DocHub-setup.exe",
            str(tmp_path / "x.exe"),
        )
    with pytest.raises(updater.DownloadNotAllowed):
        updater.fetch_text("https://evil.example.com/sha256sums.txt")


def test_human_error_explains_allowlist_rejection():
    exc = updater.DownloadNotAllowed("Скачивание разрешено только с хостов GitHub")
    assert "GitHub" in updater.human_error(exc), "пользователю видна причина"


def test_download_size_limit_aborts_by_content_length(monkeypatch, tmp_path):
    import urllib.request

    monkeypatch.setattr(
        urllib.request, "urlopen",
        lambda request, timeout=None: _FakeResponse(
            b"", headers={"Content-Length": str(updater.MAX_DOWNLOAD_BYTES + 1)},
        ),
    )
    dest = tmp_path / "big.exe"
    with pytest.raises(updater.DownloadNotAllowed):
        updater.download_file(SETUP_URL, str(dest))
    assert not dest.exists() and not (tmp_path / "big.exe.part").exists()


def test_download_size_limit_aborts_midstream(monkeypatch, tmp_path):
    """Сервер без честного Content-Length: лимит ловится по счёту порций."""
    import urllib.request

    monkeypatch.setattr(updater, "CHUNK_SIZE", 64)
    monkeypatch.setattr(updater, "MAX_DOWNLOAD_BYTES", 128)
    monkeypatch.setattr(
        urllib.request, "urlopen",
        lambda request, timeout=None: _FakeResponse(huge=True),
    )
    dest = tmp_path / "endless.exe"
    with pytest.raises(updater.DownloadNotAllowed):
        updater.download_file(SETUP_URL, str(dest))
    assert not dest.exists(), "усечённый файл не попадает в целевое имя"


def test_fetch_text_limit_rejects_oversized_sums(monkeypatch):
    import urllib.request

    monkeypatch.setattr(updater, "MAX_TEXT_BYTES", 100)
    monkeypatch.setattr(
        urllib.request, "urlopen",
        lambda request, timeout=None: _FakeResponse(b"a" * 201),
    )
    with pytest.raises(updater.DownloadNotAllowed):
        updater.fetch_text(
            "https://github.com/sutsko89/DocHub/releases/download/sha256sums.txt"
        )


def test_fetch_text_passes_within_limit(monkeypatch):
    import urllib.request

    sums = b"{" * 0 + ("abc" * 20).encode()
    monkeypatch.setattr(
        urllib.request, "urlopen",
        lambda request, timeout=None: _FakeResponse(sums),
    )
    assert updater.fetch_text(
        "https://github.com/sutsko89/DocHub/releases/download/sha256sums.txt"
    ) == sums.decode()


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
    # M7: проверка платформы переехала вместе с функцией в updater.py —
    # цель подмены следует за реализацией (контракт «вне Windows отказ»).
    monkeypatch.setattr(updater.sys, "platform", "linux")
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


# --------------------------------------------------------------------------- #
#  download_tick: цикл опроса (регрессия 2.19.2 — зависание «0,0 МБ из 19,5 МБ»)
# --------------------------------------------------------------------------- #
class FakeProgress:
    """Полоса прогресса-двойник: помнит установленные значения."""

    def __init__(self):
        self.values: list[float] = []

    def set(self, value):
        self.values.append(value)


def test_download_tick_progress_keeps_polling(app, ui_stubs):
    """Событие progress НЕ терминальное: after перевешивается, цикл живёт.

    Регрессия 2.18.0: после первого пакета прогресса опрос умирал — окно
    застывало на «0,0 МБ из 19,5 МБ», а финальное done никто не забирал.
    """
    flow = app._update_flow
    scheduled: list = []
    app.root = SimpleNamespace(after=lambda *a: scheduled.append(a))
    flow.download_queue = queue.Queue()
    flow.download_queue.put(("progress", 0, 20447232))
    flow.download_tick()
    assert scheduled, "после события прогресса after перевешивается"
    assert scheduled[-1][0] == 80
    assert scheduled[-1][1] == flow.download_tick
    assert flow.download_queue is not None, "прогресс не завершает опрос"


def test_download_tick_progress_updates_ui(app, ui_stubs):
    flow = app._update_flow
    flow._download_progress = FakeProgress()
    flow._download_status_var = FakeVar("")
    flow.download_queue = queue.Queue()
    flow.download_queue.put(("progress", 5 * 1024 * 1024, 19.5 * 1024 * 1024))
    flow.download_tick()
    assert flow._download_progress.values, "полоса прогресса обновляется"
    assert "5.0 МБ из" in flow._download_status_var.value


def test_download_tick_phase_keeps_window_and_polling(app, ui_stubs):
    """Фаза проверки суммы не закрывает окно и не останавливает опрос."""
    flow = app._update_flow
    scheduled: list = []
    app.root = SimpleNamespace(after=lambda *a: scheduled.append(a))
    flow._download_status_var = FakeVar("")
    destroyed: list = []
    flow.download_window = SimpleNamespace(
        winfo_exists=lambda: True, destroy=lambda: destroyed.append("w"))
    flow.download_queue = queue.Queue()
    flow.download_queue.put(("phase", "Проверка целостности файла…"))
    flow.download_tick()
    assert flow._download_status_var.value == "Проверка целостности файла…"
    assert destroyed == [], "фаза проверки не закрывает окно скачивания"
    assert flow.download_queue is not None, "фаза проверки не завершает опрос"
    assert scheduled, "после фазы опрос продолжается"


def test_download_tick_drains_burst_then_confirms(app, ui_stubs, monkeypatch):
    """Пачка progress+phase+done за один тик: осушается целиком до done."""
    launched = []
    monkeypatch.setattr(ucc, "run_installer_command", lambda p: launched.append(p))
    flow = app._update_flow
    flow._download_progress = FakeProgress()
    flow._download_status_var = FakeVar("")
    scheduled: list = []
    app.root = SimpleNamespace(after=lambda *a: scheduled.append(a))
    flow.download_window = SimpleNamespace(
        winfo_exists=lambda: True, destroy=lambda: None)
    flow.download_queue = queue.Queue()
    for item in (("progress", 0, 20447232),
                 ("progress", 1024 * 1024, 20447232),
                 ("phase", "Проверка целостности файла…"),
                 ("done", "/tmp/EW.DocHub-2.19.2-setup.exe")):
        flow.download_queue.put(item)
    flow.download_tick()
    assert len(flow._download_progress.values) == 2, "оба события прогресса обработаны"
    assert flow._download_status_var.value == "Проверка целостности файла…"
    assert launched == ["/tmp/EW.DocHub-2.19.2-setup.exe"], "done доходит до установки"
    assert len(ui_stubs["confirms"]) == 1, "перед установкой — подтверждение"
    assert flow.download_queue is None, "терминальное событие останавливает опрос"
    # run_installer запланировал закрытие приложения через after(500, close);
    # опрос очереди (интервал 80) после done не возобновляется.
    assert all(a[0] != 80 for a in scheduled), "после done опрос не планируется"


def test_download_tick_cancelled_stops_polling(app, ui_stubs):
    """Отмена — терминальное событие: опрос остановлен, без диалогов."""
    flow = app._update_flow
    scheduled: list = []
    app.root = SimpleNamespace(after=lambda *a: scheduled.append(a))
    flow.download_queue = queue.Queue()
    flow.download_queue.put(("cancelled", None))
    flow.download_tick()
    assert flow.download_queue is None
    assert ui_stubs["alerts"] == [] and ui_stubs["confirms"] == []
    assert scheduled == [], "терминальное событие не планирует новый опрос"


# --------------------------------------------------------------------------- #
#  Выбор ассета по флейвору (И.2 плана миграции, этап M7)
# --------------------------------------------------------------------------- #
def test_parse_installer_name_matrix():
    """Имя ассета → флейвор: Qt-суффикс длиннее Tk, проверяется первым."""
    assert updater.parse_installer_name("EW.DocHub-2.29.0-setup.exe") == "tk"
    assert updater.parse_installer_name("EW.DocHub-2.29.0-qt-setup.exe") == "qt"
    assert updater.parse_installer_name(
        "ew.docHub-2.29.0-QT-SETUP.exe".lower()) == "qt"
    assert updater.parse_installer_name("SHA256SUMS.txt") is None
    assert updater.parse_installer_name("EW.DocHub-2.29.0.zip") is None
    assert updater.parse_installer_name("") is None
    assert updater.parse_installer_name(None) is None


def test_parse_release_payload_collects_both_flavors():
    """Оба ассета в одном релизе: карта installers полная, Tk-контракты целы."""
    release = updater.parse_release_payload({
        "tag_name": "v2.29.0",
        "html_url": "https://github.com/x/tag/v2.29.0",
        "assets": [
            {"name": "EW.DocHub-2.29.0-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-setup.exe"},
            {"name": "EW.DocHub-2.29.0-qt-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe"},
            {"name": "SHA256SUMS.txt",
             "browser_download_url": "https://github.com/x/SHA256SUMS.txt"},
        ],
    })
    assert set(release["installers"]) == {"tk", "qt"}
    assert release["installers"]["qt"]["name"] == "EW.DocHub-2.29.0-qt-setup.exe"
    # Контракт 2.18.0: installer_url/installer_name — Tk-ассет (Tk-контроллер).
    assert release["installer_url"].endswith("EW.DocHub-2.29.0-setup.exe")
    assert release["installer_name"] == "EW.DocHub-2.29.0-setup.exe"
    assert release["sums_url"].endswith("SHA256SUMS.txt")


def test_parse_release_payload_qt_only_release_tk_gets_empty():
    """Релиз только с Qt-ассетом: Tk-поле пусто — Tk-сборка идёт на страницу."""
    release = updater.parse_release_payload({
        "tag_name": "v2.29.0",
        "assets": [
            {"name": "EW.DocHub-2.29.0-qt-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe"},
        ],
    })
    assert release["installer_url"] == "" and release["installer_name"] == ""
    assert set(release["installers"]) == {"qt"}


def test_installer_asset_for_flavor_returns_own_asset():
    release = updater.parse_release_payload({
        "tag_name": "v2.29.0",
        "assets": [
            {"name": "EW.DocHub-2.29.0-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-setup.exe"},
            {"name": "EW.DocHub-2.29.0-qt-setup.exe",
             "browser_download_url": "https://github.com/x/EW.DocHub-2.29.0-qt-setup.exe"},
        ],
    })
    assert updater.installer_asset_for_flavor(release, "qt")["url"] \
        .endswith("EW.DocHub-2.29.0-qt-setup.exe")
    assert updater.installer_asset_for_flavor(release, "tk")["url"] \
        .endswith("EW.DocHub-2.29.0-setup.exe")


def test_installer_asset_for_flavor_never_returns_foreign():
    """Критерий M7: чужой флейвор не скачивается.

    Qt-сборка при отсутствии Qt-ассета получает пустой словарь
    (диалог «установщик ещё не выложен»), а НЕ Tk-установщик; и наоборот.
    """
    qt_only = updater.parse_release_payload({
        "tag_name": "v2.29.0",
        "assets": [{"name": "EW.DocHub-2.29.0-qt-setup.exe",
                    "browser_download_url": "https://github.com/x/qt-setup.exe"}],
    })
    assert updater.installer_asset_for_flavor(qt_only, "qt") != {}
    assert updater.installer_asset_for_flavor(qt_only, "tk") == {}
    tk_only = updater.parse_release_payload({
        "tag_name": "v2.28.2",
        "assets": [{"name": "EW.DocHub-2.28.2-setup.exe",
                    "browser_download_url": "https://github.com/x/tk-setup.exe"}],
    })
    assert updater.installer_asset_for_flavor(tk_only, "tk") != {}
    assert updater.installer_asset_for_flavor(tk_only, "qt") == {}


def test_installer_asset_for_flavor_bad_input_is_empty():
    """Сломанный/неполный релиз — пустой результат, не исключение."""
    assert updater.installer_asset_for_flavor({}, "qt") == {}
    assert updater.installer_asset_for_flavor(None, "qt") == {}
    assert updater.installer_asset_for_flavor({"installers": "мусор"}, "qt") == {}
    assert updater.installer_asset_for_flavor({"installers": {"qt": {}}}, "qt") == {}
    assert updater.installer_asset_for_flavor(
        {"installers": {"qt": {"url": "", "name": "x"}}}, "qt") == {}


def test_match_sums_line_finds_flavor_asset_by_name():
    """SHA256SUMS.txt перечисляет оба ассета — выбор по точному имени файла.

    Формат файла не меняется (запрет 4): match_sums_line ищет строку по
    имени, флейвор-выбор делается на уровне ассета релиза.
    """
    sums = (
        f"{'a' * 64}  EW.DocHub-2.29.0-setup.exe\n"
        f"{'b' * 64}  EW.DocHub-2.29.0-qt-setup.exe\n"
    )
    assert updater.match_sums_line(
        sums, "EW.DocHub-2.29.0-qt-setup.exe") == "b" * 64
    assert updater.match_sums_line(
        sums, "ew.docHub-2.29.0-qt-setup.exe") == "b" * 64
    assert updater.match_sums_line(
        sums, "EW.DocHub-2.29.0-setup.exe") == "a" * 64
    assert updater.match_sums_line(
        sums, "EW.DocHub-2.30.0-qt-setup.exe") == ""


def test_run_installer_command_moved_to_core(tmp_path):
    """M7: реализация в GUI-free ядре; Tk-контроллер делегирует имя."""
    import updater_controller as ucc

    assert ucc.run_installer_command is not None
    # Отсутствующий файл — та же ошибка, что и до переезда (Tk-контракт).
    with pytest.raises(FileNotFoundError):
        updater.run_installer_command(str(tmp_path / "nope.exe"))
    with pytest.raises(FileNotFoundError):
        ucc.run_installer_command(str(tmp_path / "nope.exe"))
    # Не-Windows — честный отказ (обновление — Windows-сценарий).
    if not sys.platform.startswith("win"):
        with pytest.raises(OSError):
            updater.run_installer_command(str(tmp_path / "setup.exe"))
