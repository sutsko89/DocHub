"""Сверка реестра и папки документов (п. 2.4 плана).

Реестр (``database.db``) — единственный источник правды, но файлы на диске
живут отдельно: пользователь может удалить ``database.db``, скопировать
папку с другого компьютера или восстановить её из бэкапа. Модуль находит
файлы документов, которых нет в реестре, и умеет:

    * восстановить их в реестр (с привязкой к контрагенту по имени папки);
    * переместить в служебную папку ``_trash`` внутри папки документов.

Модуль не знает про Tk: интерфейс передаёт свою функцию нормализации имён
(``safe_filename``), чтобы папки сравнивались так же, как они создавались.
"""

from __future__ import annotations

import os
import re
import shutil
from datetime import datetime

import applog
import storage

TRASH_DIR_NAME = "_trash"
MAX_SCAN_DEPTH = 3
DOCUMENT_EXTENSIONS = {".docx", ".xlsx", ".xlsm", ".doc", ".xls"}

_NUMBER_RE = re.compile(r"№\s*([^\s,;()]+)")


def _norm(path: str) -> str:
    return os.path.normcase(os.path.normpath(str(path)))


def _iter_files(directory: str, depth: int = 0):
    if depth > MAX_SCAN_DEPTH:
        return
    try:
        with os.scandir(directory) as entries:
            for entry in entries:
                try:
                    if entry.is_dir(follow_symlinks=False):
                        if entry.name == TRASH_DIR_NAME:
                            continue
                        yield from _iter_files(entry.path, depth + 1)
                    elif entry.is_file(follow_symlinks=False):
                        yield entry.path
                except OSError:
                    continue
    except OSError:
        return


def scan_document_files(base_dir: str) -> list[str]:
    """Все файлы документов в папке (кроме служебной ``_trash``)."""
    if not base_dir or not os.path.isdir(base_dir):
        return []
    result = [
        path for path in _iter_files(base_dir)
        if os.path.splitext(path)[1].lower() in DOCUMENT_EXTENSIONS
    ]
    return sorted(result, key=lambda value: value.lower())


def registered_files(store, source_files: "dict | None" = None) -> set[str]:
    """Нормализованные пути файлов, на которые ссылаются записи реестра.

    LAZY-1: по умолчанию тела ленивых записей не читаются — вызывающий
    код передаёт карту ``store.bulk_source_files()`` (один read-only
    запрос, без загрузки тяжёлых тел). Без карты используется тело из
    памяти записи (совместимость со старыми вызовами).
    """
    result: set[str] = set()
    if source_files is None:
        for rec in store.documents:
            source = str((rec.get("extra") or {}).get("source_file") or "")
            if source:
                result.add(_norm(source))
        return result
    for source in source_files.values():
        if source:
            result.add(_norm(source))
    return result


def find_orphans(base_dir: str, store, source_files: "dict | None" = None) -> list[str]:
    """Файлы документов, которых нет в реестре."""
    if not base_dir or not os.path.isdir(base_dir):
        return []
    registered = registered_files(store, source_files=source_files)
    return [
        path for path in scan_document_files(base_dir)
        if _norm(path) not in registered
    ]


def guess_kind(filename: str) -> str:
    """Определяет тип документа по имени файла (для восстановления)."""
    lowered = (filename or "").casefold()
    if "спецификац" in lowered:
        return storage.KIND_SPEC
    if "соглаш" in lowered and "доп" in lowered:
        return storage.KIND_ADDITIONAL
    if "договор" in lowered:
        return storage.KIND_CONTRACT
    return storage.KIND_UNKNOWN


def extract_number(filename: str) -> str:
    """Номер документа из имени файла: «… № 516 …» → «516»."""
    match = _NUMBER_RE.search(filename or "")
    if not match:
        return ""
    return match.group(1).strip(".,;:()[]{}\"'").strip()


def _folder_key(name: str, name_key=None) -> str:
    value = name_key(name) if name_key else name
    return str(value or "").strip().casefold()


def guess_party_id(path: str, base_dir: str, store, name_key=None):
    """Контрагент по имени папки, в которой лежит файл (если удаётся сопоставить)."""
    try:
        relative = os.path.relpath(os.path.dirname(path), base_dir)
    except ValueError:
        return None
    if not relative or relative == ".":
        return None
    folder = relative.split(os.sep)[0]
    if folder == "unknown":
        return None
    key = _folder_key(folder, name_key)
    if not key:
        return None
    for rec in store.active_counterparties():
        if _folder_key(storage.display_name(rec), name_key) == key:
            return rec.get("id")
    return None


def restore_orphans(store, base_dir: str, paths: list[str], name_key=None) -> int:
    """Создаёт записи реестра для найденных файлов. Возвращает число записей.

    D1: раньше в цикле проверялся ``store.last_error`` после
    ``add_document(save=False)`` — при save=False он не обновляется, поэтому
    проверка либо ложно прерывала цикл по устаревшей ошибке (восстановление
    обрывалось после первой записи), либо не проверяла ничего. Теперь ошибки
    добавления собираются локально, а финальный ``save()`` проверяется
    отдельно: реальный сбой записи даёт честное сообщение, и пользователь
    видит, сколько файлов восстановлено и сколько нет.
    """
    log = applog.get_logger("reconcile")
    # Сбрасываем устаревшую ошибку прошлых операций: при save=False она
    # не перезаписывается и не должна прерывать восстановление.
    store.last_error = ""
    added = 0
    failed: list[str] = []
    for path in paths:
        filename = os.path.basename(path)
        try:
            party_id = guess_party_id(path, base_dir, store, name_key)
            party = store.get_counterparty(party_id) if party_id is not None else None
            fields = dict((party or {}).get("fields") or {})
            rec = store.add_document(
                guess_kind(filename),
                party_id,
                extract_number(filename),
                "",
                data=fields,
                extra={
                    "imported": True,
                    "source_file": path,
                    "restored": True,
                    "is_ip": (party or {}).get("type") == "ip",
                },
                text="",
                save=False,
            )
            if rec is None:
                raise RuntimeError("add_document не вернул запись")
        except Exception as exc:  # noqa: BLE001 — один битый файл не роняет цикл
            failed.append(filename)
            log.error(
                "Не удалось добавить запись восстановления (%s): %s",
                filename, type(exc).__name__,
            )
            continue
        added += 1
    if added:
        store.save()
    if store.last_error:
        log.error(
            "Не удалось сохранить восстановленные записи (%d шт.): %s",
            added, store.last_error,
        )
    if failed:
        log.warning("Не восстановлено файлов: %d из %d", len(failed), len(paths))
    if added:
        log.info("В реестр восстановлено файлов: %d", added)
    return added


def move_orphans_to_trash(base_dir: str, paths: list[str]) -> int:
    """Переносит файлы в ``<папка документов>/_trash/<дата-время>/``."""
    if not paths:
        return 0
    log = applog.get_logger("reconcile")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    target_root = os.path.join(base_dir, TRASH_DIR_NAME, stamp)
    moved = 0
    for path in paths:
        try:
            relative = os.path.relpath(path, base_dir)
        except ValueError:
            relative = os.path.basename(path)
        if relative.startswith(".."):
            relative = os.path.basename(path)
        dest = os.path.join(target_root, relative)
        base_norm = _norm(target_root)
        if not _norm(dest).startswith(base_norm):
            continue
        try:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            candidate = dest
            counter = 2
            stem, ext = os.path.splitext(dest)
            while os.path.exists(candidate):
                candidate = f"{stem} ({counter}){ext}"
                counter += 1
            shutil.move(path, candidate)
            moved += 1
        except OSError as exc:
            log.warning(
                "Не удалось переместить файл в корзину (%s): %s",
                os.path.splitext(path)[1].lower(), type(exc).__name__,
            )
    if moved:
        log.info("Файлов перемещено в _trash: %d", moved)
    return moved
