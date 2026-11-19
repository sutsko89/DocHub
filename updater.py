"""Проверка обновлений EW.DocHub через GitHub Releases API.

Первый публичный релиз (2.15.0): приложение само проверяет наличие новых
версий и предлагает пользователю открыть страницу релиза. Модуль
намеренно без Tk — только версии, HTTP и разбор ответа GitHub:

* ``parse_version`` / ``is_newer`` — сравнение версий «v2.15.0» ↔ «2.15.0»;
* ``fetch_latest_release`` — GET последнего релиза GitHub (stdlib urllib,
  никаких новых зависимостей — работает и в сборке PyInstaller);
* ``check_for_update`` — единая точка проверки с понятным пользователю
  результатом («доступна новая», «версия последняя», «ошибка сети»);
* ``auto_check_due`` / ``mark_checked`` — суточный троттлинг
  автопроверки через settings.json (ключ ``last_update_check``).

Оркестрация с очередью и опросом из главного потока Tk — в
updater_controller.py (паттерн dadata_controller: рабочий поток не трогает
Tk). Модуль не импортирует main — APP_VERSION передаётся аргументом,
поэтому верхнеуровневый импорт updater.py безопасен из любого модуля.

Отключение проверки (приватность, offline-среды, тесты): переменная
окружения ``EWDOCHUB_NO_UPDATE_CHECK=1`` — приоритетнее всех настроек;
в tests/conftest.py и smoke_test.py она выставляется, поэтому тесты не
обращаются в сеть. Пользовательский переключатель — настройка
``auto_update_check`` в settings.json (по умолчанию включено).
"""

from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.request

# Публичный репозиторий проекта: последний релиз (не pre-release, не draft).
RELEASES_API_URL = "https://api.github.com/repos/sutsko89/DocHub/releases/latest"
PAGE_URL = "https://github.com/sutsko89/DocHub/releases"

TIMEOUT = 8  # секунд — автопроверка не должна заметно тормозить старт

# Автопроверка не чаще раза в сутки (суточный троттлинг, см. auto_check_due).
AUTO_CHECK_INTERVAL = 24 * 60 * 60

# Env-глушилка: тесты/smoke/офлайн-среды не должны ходить в сеть.
DISABLE_ENV = "EWDOCHUB_NO_UPDATE_CHECK"

# Ключи settings.json: время последней проверки и переключатель пользователя.
SETTINGS_KEY_TS = "last_update_check"
SETTINGS_KEY_AUTO = "auto_update_check"

_USER_AGENT = "EW.DocHub-Updater"


# --------------------------------------------------------------------------- #
#  Включённость проверки
# --------------------------------------------------------------------------- #
def check_enabled() -> bool:
    """Общий рубильник проверки: False — никаких сетевых вызовов.

    Env-глушилка приоритетнее пользовательской настройки: она существует
    для тестов и офлайн-сред, где сетевой вызов недопустим в принципе.
    """
    return not os.environ.get(DISABLE_ENV)


def auto_check_enabled(settings: dict) -> bool:
    """Пользовательский переключатель авто-проверки (по умолчанию включён)."""
    return bool(settings.get(SETTINGS_KEY_AUTO, True))


def auto_check_due(settings: dict, now: float | None = None) -> bool:
    """Пора ли снова автопроверять: не чаще раза в AUTO_CHECK_INTERVAL."""
    ts = settings.get(SETTINGS_KEY_TS)
    ts = float(ts) if isinstance(ts, (int, float)) else 0.0
    current = time.time() if now is None else float(now)
    return (current - ts) >= AUTO_CHECK_INTERVAL


def mark_checked(settings: dict) -> None:
    """Фиксирует факт проверки в настройках (сохранение — на приложении)."""
    settings[SETTINGS_KEY_TS] = time.time()


# --------------------------------------------------------------------------- #
#  Версии
# --------------------------------------------------------------------------- #
def parse_version(text: str) -> tuple[int, ...] | None:
    """«v2.15.0» → (2, 15, 0). Битый/нецифровой тег → None."""
    match = re.search(r"(\d+(?:\.\d+)+)", str(text or ""))
    if not match:
        return None
    try:
        return tuple(int(part) for part in match.group(1).split("."))
    except ValueError:
        return None


def is_newer(candidate: str, current: str) -> bool:
    """True, если ``candidate`` — строго более новая версия, чем ``current``.

    Версии разной длины выравниваются нулями: «2.15» равно «2.15.0».
    Неразобираемые версии (None) обновлением не считаются — лучше пропустить
    проверку, чем открыть диалог по мусорному тегу.
    """
    candidate_version = parse_version(candidate)
    current_version = parse_version(current)
    if candidate_version is None or current_version is None:
        return False
    width = max(len(candidate_version), len(current_version))
    candidate_version += (0,) * (width - len(candidate_version))
    current_version += (0,) * (width - len(current_version))
    return candidate_version > current_version


# --------------------------------------------------------------------------- #
#  GitHub Releases API
# --------------------------------------------------------------------------- #
def parse_release_payload(payload: object) -> dict | None:
    """Разбирает JSON ответа GitHub в компактный словарь релиза.

    Возвращает None для несловарного ответа, draft-релиза и ответа без
    тега (например, если релизов ещё нет — GitHub вернёт 404, это ловится
    выше). ``installer_url`` — ссылка на ``*-setup.exe`` из ассетов, если
    установщик уже выложен.
    """
    if not isinstance(payload, dict):
        return None
    if payload.get("draft"):
        return None
    tag = str(payload.get("tag_name") or "").strip()
    if not tag:
        return None
    version = parse_version(tag)
    assets = [
        asset for asset in (payload.get("assets") or [])
        if isinstance(asset, dict)
    ]
    installer_url = ""
    for asset in assets:
        name = str(asset.get("name") or "").strip().lower()
        if name.endswith("-setup.exe"):
            installer_url = str(asset.get("browser_download_url") or "")
            break
    return {
        "tag": tag,
        "version": version,
        "version_str": ".".join(str(part) for part in version) if version else tag,
        "name": str(payload.get("name") or tag),
        "url": str(payload.get("html_url") or "").strip() or PAGE_URL,
        "notes": str(payload.get("body") or ""),
        "installer_url": installer_url,
    }


def human_error(exc: Exception) -> str:
    """Сетевую ошибку переводит в сообщение, понятное пользователю."""
    if isinstance(exc, urllib.error.HTTPError):
        if exc.code == 403 or exc.code == 429:
            return (
                "GitHub ограничил частоту запросов. Попробуйте проверить "
                "обновления позже."
            )
        if exc.code == 404:
            return (
                "Страница релизов недоступна (GitHub вернул 404). "
                "Релизы ещё не опубликованы."
            )
        return f"GitHub вернул ошибку {exc.code}."
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, TimeoutError) or "timed out" in str(reason).lower():
            return "Сервис GitHub не отвечает — проверьте интернет и попробуйте ещё раз."
        return "Нет соединения с github.com — проверьте интернет."
    if isinstance(exc, TimeoutError):
        return "Сервис GitHub не отвечает — проверьте интернет и попробуйте ещё раз."
    return f"Не удалось проверить обновления: {type(exc).__name__}"


def fetch_latest_release(timeout: float = TIMEOUT) -> dict | None:
    """Запрашивает последний релиз GitHub. None — релизов нет/ответ пуст.

    Исключения сети (HTTPError, URLError, TimeoutError) отдаются наружу —
    их переводит в понятный текст check_for_update (human_error).
    """
    request = urllib.request.Request(
        RELEASES_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": _USER_AGENT,
        },
        method="GET",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return parse_release_payload(payload)


def check_for_update(current: str, timeout: float = TIMEOUT) -> dict:
    """Единая точка проверки обновлений — не бросает исключений.

    Возвращает словарь:
      * ``status``: ``"update"`` — есть новая версия, ``"latest"`` — версия
        последняя, ``"error"`` — проверка не удалась;
      * ``message`` — текст, готовый для показа пользователю;
      * ``release`` — словарь релиза (parse_release_payload) или None.
    """
    try:
        release = fetch_latest_release(timeout=timeout)
    except Exception as exc:  # сеть/разбор: пользовательская ошибка, не traceback
        return {"status": "error", "message": human_error(exc), "release": None}
    if release is None or release["version"] is None:
        return {
            "status": "error",
            "message": "Ответ GitHub не удалось разобрать — попробуйте позже.",
            "release": None,
        }
    if is_newer(release["tag"], current):
        return {
            "status": "update",
            "message": f"Доступна новая версия {release['version_str']}.",
            "release": release,
        }
    return {
        "status": "latest",
        "message": f"Вы используете последнюю версию ({current}).",
        "release": release,
    }
