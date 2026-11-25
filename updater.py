"""Проверка обновлений EW.DocHub через GitHub Releases API.

Первый публичный релиз (2.15.0): приложение само проверяет наличие новых
версий и предлагает пользователю обновиться. Модуль намеренно без Tk —
только версии, HTTP и разбор ответа GitHub:

* ``parse_version`` / ``is_newer`` — сравнение версий «v2.15.0» ↔ «2.15.0»;
* ``fetch_latest_release`` — GET последнего релиза GitHub (stdlib urllib,
  никаких новых зависимостей — работает и в сборке PyInstaller);
* ``check_for_update`` — единая точка проверки с понятным пользователю
  результатом («доступна новая», «версия последняя», «ошибка сети»);
* ``auto_check_due`` / ``mark_checked`` — суточный троттлинг
  автопроверки через settings.json (ключ ``last_update_check``);
* 2.18.0 — автообновление: ``download_file`` (потоковое скачивание
  установщика с прогрессом и отменой), ``fetch_text``/``match_sums_line``
  (SHA256SUMS.txt) и ``verify_sha256`` (контроль целостности перед
  запуском установщика);
* M7 (И.2 плана миграции) — выбор ассета релиза по флейвору сборки:
  рядом с Tk-ассетом (``*-setup.exe`` без суффикса — историческое имя)
  в релизе лежит Qt-ассет (``*-qt-setup.exe``); ``parse_release_payload``
  разбирает оба, ``installer_asset_for_flavor`` отдаёт свой;
  ``run_installer_command`` — тихий запуск установщика Inno Setup
  (переехал из updater_controller.py: GUI-free модуль нужен и Qt-сборке,
  у которой стек Tk исключён из поставки).

Оркестрация с очередью и опросом из главного потока GUI — в
updater_controller.py (Tk) и updater_qt.py (Qt, этап M7): рабочий поток
не трогает интерфейс. Модуль не импортирует main — APP_VERSION передаётся
аргументом, поэтому верхнеуровневый импорт updater.py безопасен из любого
модуля.

Отключение проверки (приватность, offline-среды, тесты): переменная
окружения ``EWDOCHUB_NO_UPDATE_CHECK=1`` — приоритетнее всех настроек;
в tests/conftest.py и smoke_test.py она выставляется, поэтому тесты не
обращаются в сеть. Пользовательский переключатель — настройка
``auto_update_check`` в settings.json (по умолчанию включено).
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.parse
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

# Скачивание установщика (2.18.0): размер порции и имя ассета-сумм.
CHUNK_SIZE = 256 * 1024
SUMS_ASSET_NAME = "sha256sums.txt"

# И.2 (M7): суффиксы флейворов в имени ассета-установщика. Tk-сборка
# исторически без суффикса («EW.DocHub-2.28.2-setup.exe»), Qt-сборка —
# «EW.DocHub-2.28.2-qt-setup.exe». Сравнение версий — только по числовой
# части тега: Qt-сборка не «новее» Tk той же версии (parse_version/is_newer
# без изменений). Формат SHA256SUMS.txt (два пробела, hex64 + имя) не меняется.
TK_SETUP_SUFFIX = "-setup.exe"
QT_SETUP_SUFFIX = "-qt-setup.exe"
# Тег флейвора → суффикс имени ассета (единственный источник знания об
# именах: и release.yml, и installer_qt.iss собирают ассеты по этому правилу).
FLAVOR_SETUP_SUFFIX = {"tk": TK_SETUP_SUFFIX, "qt": QT_SETUP_SUFFIX}

# SEC-1 (D11): канал обновлений ходит только на хосты GitHub. Ссылки
# ассетов приходят из JSON релиза — компрометация релиза/перенаправление
# больше не подменяет источник установщика. objects.githubusercontent.com —
# фактический хост раздачи ассетов GitHub Releases (redirect от github.com).
ALLOWED_DOWNLOAD_HOSTS = frozenset({
    "api.github.com",
    "github.com",
    "objects.githubusercontent.com",
    "release-assets.githubusercontent.com",
})

# SEC-1 (D11): лимит размера скачивания. Установщик PyInstaller ~ 40–80 МБ;
# лимит — с запасом на рост. Превышение обрывает загрузку: усечённый файл
# всё равно не пройдёт проверку SHA-256, тянуть гигабайты бессмысленно.
MAX_DOWNLOAD_BYTES = 512 * 1024 * 1024
# SHA256SUMS.txt — маленький текстовый файл; большего не бывает.
MAX_TEXT_BYTES = 1024 * 1024

_USER_AGENT = "EW.DocHub-Updater"


class DownloadCancelled(Exception):
    """Пользователь отменил скачивание обновления (кнопка «Отмена»)."""


class DownloadNotAllowed(Exception):
    """URL вне allowlist хостов или превышен лимит размера (SEC-1, D11).

    Возбуждается ДО сетевого запроса для чужого хоста — скачивание не
    начинается вовсе; превышение размера обрывает загрузку посреди потока.
    """


# --------------------------------------------------------------------------- #
#  Включённость проверки
# --------------------------------------------------------------------------- #
def check_enabled() -> bool:
    """Общий рубильник проверки: False — никаких сетевых вызовов.

    Env-глушилка приоритетнее пользовательской настройки: она существует
    для тестов и офлайн-сред, где сетевой вызов недопустим в принципе.
    """
    return not os.environ.get(DISABLE_ENV)


def is_allowed_download_url(url: str) -> bool:
    """True, если URL можно скачивать: https и хост из allowlist (SEC-1, D11).

    Небезопасные схемы (http, file, ftp) и чужие хосты отклоняются.
    """
    try:
        parts = urllib.parse.urlparse(str(url or "").strip())
    except ValueError:
        return False
    if parts.scheme.lower() != "https" or not parts.hostname:
        return False
    return parts.hostname.lower() in ALLOWED_DOWNLOAD_HOSTS


def ensure_allowed_download_url(url: str) -> str:
    """Проверяет URL и возвращает его; чужой хост — исключение без запроса.

    Вызывается до построения Request: сторонний URL отклоняется БЕЗ
    скачивания (критерий карточки SEC-1).
    """
    if not is_allowed_download_url(url):
        host = ""
        try:
            host = urllib.parse.urlparse(str(url or "")).hostname or ""
        except ValueError:
            pass
        raise DownloadNotAllowed(
            "Скачивание разрешено только с хостов GitHub "
            f"({', '.join(sorted(ALLOWED_DOWNLOAD_HOSTS))}). "
            f"Получен адрес с хостом: {host or '(не определён)'}"
        )
    return str(url)


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
    выше). ``installer_url``/``installer_name`` — Tk-ассет, если он
    выложен (контракт 2.18.0, Tk-контроллер читает напрямую);
    ``installers`` — карта «флейвор → {url, name}» для выбора ассета
    своего флейвора (И.2, M7): чужой флейвор через неё не достать.
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
    sums_url = ""
    # И.2 (M7): установщики по флейворам. Один проход по ассетам;
    # дубликат ассета того же флейвора не перезаписывает первый.
    installers: dict[str, dict] = {}
    for asset in assets:
        name = str(asset.get("name") or "").strip()
        lowered = name.lower()
        flavor = parse_installer_name(lowered)
        if flavor is not None:
            asset_url = str(asset.get("browser_download_url") or "")
            if flavor not in installers:
                installers[flavor] = {"url": asset_url, "name": name}
        elif lowered == SUMS_ASSET_NAME:
            sums_url = str(asset.get("browser_download_url") or "")
    tk_asset = installers.get("tk") or {}
    installer_url = tk_asset.get("url", "")
    return {
        "tag": tag,
        "version": version,
        "version_str": ".".join(str(part) for part in version) if version else tag,
        "name": str(payload.get("name") or tag),
        "url": str(payload.get("html_url") or "").strip() or PAGE_URL,
        "notes": str(payload.get("body") or ""),
        "installer_url": installer_url,
        "sums_url": sums_url,
        "installer_name": os.path.basename(installer_url.split("?")[0])
        if installer_url else "",
        "installers": installers,
    }


def parse_installer_name(name: str) -> str | None:
    """Флейвор по имени ассета: «*-qt-setup.exe» → «qt», «*-setup.exe» → «tk».

    Tk-суффикс исторически без маркера флейвора, поэтому сначала
    проверяется более длинный Qt-суффикс: строка, оканчивающаяся на
    «-qt-setup.exe», — Qt-ассет, а не Tk. Не-установщик (SHA256SUMS.txt,
    исходники) → None.
    """
    lowered = str(name or "").strip().lower()
    if lowered.endswith(QT_SETUP_SUFFIX):
        return "qt"
    if lowered.endswith(TK_SETUP_SUFFIX):
        return "tk"
    return None


def installer_asset_for_flavor(release: dict, flavor: str) -> dict:
    """Ассет СВОЕГО флейвора из разобранного релиза: {"url", "name"} | {}.

    Критерий M7: чужой флейвор не скачивается — Qt-сборка при отсутствии
    Qt-ассета получает пустой словарь (диалог «установщик ещё не выложен»
    со ссылкой на страницу), а не Tk-установщик; и наоборот.
    """
    assets = (release or {}).get("installers")
    if not isinstance(assets, dict):
        return {}
    asset = assets.get(str(flavor or ""))
    if not isinstance(asset, dict) or not str(asset.get("url") or ""):
        return {}
    return {"url": str(asset["url"]), "name": str(asset.get("name") or "")}


def human_error(exc: Exception) -> str:
    """Сетевую ошибку переводит в сообщение, понятное пользователю."""
    if isinstance(exc, DownloadNotAllowed):
        # SEC-1 (D11): отказ allowlist/лимита — не «ошибка сети», а понятный
        # запрет с причиной; текст уже сформулирован в исключении.
        return str(exc)
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


# --------------------------------------------------------------------------- #
#  Скачивание установщика и контроль целостности (2.18.0)
# --------------------------------------------------------------------------- #
def download_file(url: str, dest_path: str, progress_cb=None,
                  cancel_event=None, timeout: float = 30.0) -> str:
    """Скачивает файл потоком в ``dest_path`` и возвращает путь.

    * ``progress_cb(done, total)`` вызывается на каждой порции (total может
      быть 0, если сервер не отдал Content-Length);
    * ``cancel_event`` — threading.Event: установка события прерывает
      скачивание исключением DownloadCancelled;
    * чистый stdlib urllib — работает и в сборке PyInstaller, никаких
      новых зависимостей.

    Файл пишется во временный ``.part`` и переименовывается по успехе,
    поэтому после отмены/ошибки не остаётся «полуфабриката» с целевым
    именем.

    SEC-1 (D11): URL проверяется по allowlist хостов ДО сетевого запроса;
    превышение ``MAX_DOWNLOAD_BYTES`` — по Content-Length до начала и
    по факту счёта порций во время загрузки — обрывает загрузку.
    """
    ensure_allowed_download_url(url)
    request = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    part_path = dest_path + ".part"
    os.makedirs(os.path.dirname(os.path.abspath(dest_path)), exist_ok=True)
    done = 0
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            total = int(response.headers.get("Content-Length") or 0)
            if total > MAX_DOWNLOAD_BYTES:
                raise DownloadNotAllowed(
                    f"Файл больше лимита {MAX_DOWNLOAD_BYTES} байт: {total}"
                )
            if callable(progress_cb):
                progress_cb(0, total)
            with open(part_path, "wb") as out:
                while True:
                    if cancel_event is not None and cancel_event.is_set():
                        raise DownloadCancelled()
                    chunk = response.read(CHUNK_SIZE)
                    if not chunk:
                        break
                    out.write(chunk)
                    done += len(chunk)
                    if callable(progress_cb):
                        progress_cb(done, total)
                    if done > MAX_DOWNLOAD_BYTES:
                        raise DownloadNotAllowed(
                            f"Загрузка превысила лимит {MAX_DOWNLOAD_BYTES} байт — оборвана"
                        )
        os.replace(part_path, dest_path)
        return dest_path
    finally:
        try:
            if os.path.exists(part_path):
                os.remove(part_path)
        except OSError:
            pass  # Windows может держать файл открытым до завершения потока


def fetch_text(url: str, timeout: float = TIMEOUT) -> str:
    """Скачивает небольшой текстовый файл целиком (SHA256SUMS.txt).

    SEC-1 (D11): allowlist хостов до запроса и лимит размера — читается
    не больше ``MAX_TEXT_BYTES``; больший ответ отклоняется.
    """
    ensure_allowed_download_url(url)
    request = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        data = response.read(MAX_TEXT_BYTES + 1)
    if len(data) > MAX_TEXT_BYTES:
        raise DownloadNotAllowed(
            f"Текстовый файл больше лимита {MAX_TEXT_BYTES} байт"
        )
    return data.decode("utf-8", errors="replace")


def match_sums_line(sums_text: str, filename: str) -> str:
    """Из содержимого SHA256SUMS.txt берёт хеш для ``filename``.

    Формат GitHub-релиза — ``<hex64>  <имя>`` (два пробела; допустим
    бинарный маркер ``*``). Возвращает hex-строку в нижнем регистре или
    пустую строку, если файла в списке нет.
    """
    target = str(filename or "").strip().lower()
    if not target:
        return ""
    for line in str(sums_text or "").splitlines():
        parts = line.strip().split(None, 1)
        if len(parts) != 2:
            continue
        digest, name = parts[0].strip().lower(), parts[1].strip()
        if name.startswith("*"):
            name = name[1:]
        if name.strip().lower() == target:
            return digest if re.fullmatch(r"[0-9a-f]{64}", digest) else ""
    return ""


def verify_sha256(path: str, expected_hex: str) -> bool:
    """Считает SHA-256 файла потоком и сравнивает с ожидаемым значением."""
    expected = str(expected_hex or "").strip().lower()
    if not expected:
        return False
    digest = hashlib.sha256()
    with open(path, "rb") as source:
        for chunk in iter(lambda: source.read(CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest() == expected


def run_installer_command(path: str) -> None:
    """Стартует установщик Inno Setup в тихом режиме (M7: GUI-free ядро).

    /VERYSILENT — без вопросов (обновление в ту же папку, профиль в
    %APPDATA% не трогается); /NORESTART — без перезагрузки;
    /SUPPRESSMSGBOXES и /FORCECLOSEAPPLICATIONS — тихо закрывают
    запущенную программу, если она не успела выйти (все данные
    сохраняются сразу при изменениях, риск минимальный); после установки
    приложение перезапускается само ([Run] skipifnotsilent). На других ОС
    функция не запускается (обновление — Windows-сценарий).

    Функция живёт в ядре (переехала из updater_controller.py в M7):
    updater_controller переэкспортирует имя для обратной совместимости
    тестов/smoke, updater_qt вызывает напрямую — стек Tk в Qt-сборке
    исключён, импорт Tk-контроллера из неё невозможен.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Установщик не найден: {path}")
    if not sys.platform.startswith("win"):
        raise OSError("Автообновление доступно только в Windows.")
    # creationflags: DETACHED_PROCESS — установщик живёт после закрытия
    # приложения; без консольного окна.
    detached = getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(
        [path, "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES",
         "/FORCECLOSEAPPLICATIONS"],
        close_fds=True,
        creationflags=detached,
    )
