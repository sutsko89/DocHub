"""Клиент DaData: автозаполнение реквизитов по ИНН.

Используется официальный API DaData (https://dadata.ru) — то же, что
подключается к 1С через сервис «1С:Контрагент». API-ключ выдаётся
бесплатно (лимит 10 000 запросов в день) и хранится локально
в профиле пользователя.

Для запросов используется только стандартная библиотека — никаких
дополнительных зависимостей, работает и в сборке PyInstaller.
"""

from __future__ import annotations

import json
import os
import urllib.error
import urllib.request

API_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party"
TIMEOUT = 10  # секунд

# Ключ в программу НЕ вшивается: пользователь вводит его в настройках.
# Это исключает распространение общего ключа вместе со сборкой и позволяет
# сменить ключ без пересборки. SECRET_KEY программе не нужен.
KEY_DIR_NAME = "DogovorGenerator"
LEGACY_KEY_DIR_NAME = "DogovorPostavki"  # прежнее расположение файла ключа
KEY_FILE_NAME = "dadata_key.txt"


def get_api_key() -> str:
    """Ключ для запросов: только сохранённый пользователем."""
    return load_api_key()


class DadataError(Exception):
    """Понятная пользователю ошибка запроса к DaData."""


def _profile_dir(name: str) -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, name)


def key_storage_path() -> str:
    """Путь к файлу с API-ключом (каталог данных пользователя)."""
    directory = _profile_dir(KEY_DIR_NAME)
    os.makedirs(directory, exist_ok=True)
    return os.path.join(directory, KEY_FILE_NAME)


def legacy_key_storage_path() -> str:
    """Прежний путь ключа (до переноса в общий каталог данных)."""
    return os.path.join(_profile_dir(LEGACY_KEY_DIR_NAME), KEY_FILE_NAME)


def load_api_key() -> str:
    """Читает ключ из нового расположения, затем из прежнего (миграция)."""
    for path in (key_storage_path(), legacy_key_storage_path()):
        try:
            with open(path, "r", encoding="utf-8") as file:
                key = file.read().strip()
        except OSError:
            continue
        if key:
            return key
    return ""


def save_api_key(key: str) -> bool:
    """Атомарно сохраняет ключ. Возвращает False, если записать не удалось."""
    key = (key or "").strip()
    if not key or any(ch.isspace() for ch in key):
        return False
    path = key_storage_path()
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as file:
            file.write(key)
        os.replace(tmp, path)
        return True
    except OSError:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        return False


def clear_api_key() -> bool:
    """Удаляет сохранённый ключ (включая прежнее расположение)."""
    removed = False
    for path in (key_storage_path(), legacy_key_storage_path()):
        try:
            os.remove(path)
            removed = True
        except OSError:
            pass
    return removed


def fetch_party(inn: str, api_key: str) -> dict:
    """Запрашивает карточку организации/ИП по ИНН.

    Возвращает словарь с разобранными полями; при проблемах raises
    :class:`DadataError` с сообщением, готовым для показа пользователю.
    """
    request = urllib.request.Request(
        API_URL,
        data=json.dumps({"query": inn}).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Accept": "application/json",
            "Authorization": f"Token {api_key}",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 401:
            raise DadataError("Неверный API-ключ DaData. Проверьте ключ на dadata.ru.") from exc
        if exc.code == 429:
            raise DadataError(
                "Превышен дневной лимит запросов DaData (бесплатный тариф — 10 000 в день)."
            ) from exc
        raise DadataError(f"Сервис DaData вернул ошибку {exc.code}.") from exc
    except urllib.error.URLError as exc:
        raise DadataError("Нет соединения с api.dadata.ru — проверьте интернет.") from exc
    except TimeoutError as exc:
        raise DadataError("Сервис DaData не отвечает — попробуйте ещё раз.") from exc

    suggestions = payload.get("suggestions") or []
    if not suggestions:
        raise DadataError("Организация с таким ИНН не найдена.")
    return _parse(suggestions[0].get("data") or {})


def _parse(data: dict) -> dict:
    name = data.get("name") or {}
    address = data.get("address") or {}
    management = data.get("management") or {}
    fio = data.get("fio") or {}
    bank = data.get("bank") or {}

    full_name = name.get("full_with_opf") or name.get("full") or ""
    short_name = name.get("short_with_opf") or name.get("short") or ""

    # Для ООО это ФИО руководителя, для ИП — ФИО самого предпринимателя.
    director = management.get("name", "")
    if not director and fio:
        director = " ".join(
            part
            for part in (fio.get("surname"), fio.get("name"), fio.get("patronymic"))
            if part
        )

    return {
        "full_name": full_name,
        "short_name": short_name,
        "director": director,
        "legal_address": address.get("value", ""),
        "inn": str(data.get("inn") or ""),
        "kpp": str(data.get("kpp") or ""),
        "ogrn": str(data.get("ogrn") or ""),  # для ИП это ОГРНИП (15 цифр)
        "bank_name": bank.get("name", ""),
        "bik": str(bank.get("bic") or ""),
        "correspondent_account": str(bank.get("correspondent_account") or ""),
    }
