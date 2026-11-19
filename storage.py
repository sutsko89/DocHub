"""Хранилище данных программы: контрагенты и созданные документы.

Реестр ведётся в SQLite-базе в профиле пользователя
(%APPDATA%/DogovorGenerator/database.db), поэтому работают и в разработке,
и внутри собранного PyInstaller .exe (п. 11 этапа 5 плана).

    * WAL-журнал: чтение не блокирует запись, второй экземпляр программы
      безопасно работает параллельно (busy_timeout, транзакции);
    * каждая запись — строка таблицы, есть индексы: по ИНН
      (counterparties.inn), по контрагенту и типу документа (documents);
    * запись идёт одной транзакцией: либо весь реестр, либо ничего —
      при сбое предыдущая версия базы не теряется;
    * при временных сбоях (антивирус, индексатор) — до 3 попыток;
    * старый JSON-реестр database.json переносится в базу при первом
      запуске и сохраняется как database.json.migrated.

Структура записей в памяти не изменилась:
    контрагент: {"id", "type": "ooo"|"ip", "fields": {...}, "created_at"}
    документ:   {"id", "kind": "contract"|"additional"|"spec",
                 "party_id", "number", "date", "data": {...},
                 "extra": {...}, "text", "created_at"}

Корзина (п. 2.6): удаление помечает запись полем ``deleted_at``, а не стирает
её. Восстановление снимает метку, окончательная очистка удаляет физически.
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
import time
from datetime import datetime, timedelta

import applog

# Единый источник имён полей (п. 9.6): storage пишет и читает ключи так же,
# как их задаёт карточка контрагента и печатные формы. Цикла импорта нет:
# document_builder не зависит от storage.
from document_builder import F

SETTINGS_DIR_NAME = "DogovorGenerator"

DB_FILENAME = "database.db"              # источник правды (SQLite)
LEGACY_JSON_FILENAME = "database.json"   # реестр прежних версий (только чтение)

# v2.13.0: опечатка «(Покупь)» в имени поля исправлена на «(Покупатель)».
# Каноническое имя — document_builder.F.FULL_NAME; старая строка нужна
# только для миграции записей при загрузке (см. _normalize_field_names).
LEGACY_FULL_NAME_FIELD = "Полное название компании (Покупь)"


def _rename_field_key(payload: dict) -> None:
    """Переименовывает один ключ в словаре полей: старый → канонический.

    Если канонический ключ уже заполнен (обе версии в одной записи),
    остаётся новое значение — старое отбрасывается, чтобы предпросмотр
    и печатная форма не увидели расхождения.
    """
    if LEGACY_FULL_NAME_FIELD not in payload:
        return
    value = payload.pop(LEGACY_FULL_NAME_FIELD)
    payload.setdefault(F.FULL_NAME, value)


def _normalize_field_names(rec: dict) -> None:
    """Миграция имён полей v2.13.0 для записи реестра (идемпотентно).

    Затрагивает карточку контрагента (``fields``) и сохранённые данные
    документа (``data`` — копия полей карточки на момент сохранения).
    """
    fields = rec.get("fields")
    if isinstance(fields, dict):
        _rename_field_key(fields)
    data = rec.get("data")
    if isinstance(data, dict):
        _rename_field_key(data)


SQLITE_SCHEMA_VERSION = 1
SCHEMA_VERSION = 2              # схема старого JSON-реестра (для переноса)
BUSY_TIMEOUT_SECONDS = 5.0      # ждём освобождения базы вместо мгновенного сбоя
SAVE_RETRIES = 3
RETRY_DELAY_SECONDS = 0.25
TRASH_RETENTION_DAYS = 30
BAD_COPY_RETENTION_DAYS = 30   # сколько дней хранить database.db.bad-*

KIND_CONTRACT = "contract"
KIND_ADDITIONAL = "additional"
KIND_SPEC = "spec"
KIND_UNKNOWN = "unknown"   # импортированный файл, который не удалось опознать

KIND_LABELS = {
    KIND_CONTRACT: "Договор поставки",
    KIND_ADDITIONAL: "Доп. соглашение",
    KIND_SPEC: "Спецификация",
    KIND_UNKNOWN: "Неопознанный документ",
}

# Ключи записи, у которых в базе есть собственные колонки. Всё остальное
# (запас на будущее) сохраняется в колонке extra_keys как JSON — перенос
# в обе стороны не теряет ни одного ключа.
_PARTY_KEYS = ("id", "type", "fields", "created_at", "deleted_at")
_DOCUMENT_KEYS = ("id", "kind", "party_id", "number", "date", "data", "extra",
                  "text", "created_at", "deleted_at", "deleted_with_party")

_SCHEMA_SQL = (
    """
    CREATE TABLE IF NOT EXISTS counterparties (
        id         INTEGER PRIMARY KEY,
        type       TEXT NOT NULL DEFAULT '',
        fields     TEXT NOT NULL DEFAULT '{}',
        created_at TEXT NOT NULL DEFAULT '',
        deleted_at TEXT,
        inn        TEXT NOT NULL DEFAULT '',
        extra_keys TEXT NOT NULL DEFAULT '{}'
    )
    """,
    "CREATE INDEX IF NOT EXISTS idx_counterparties_inn ON counterparties(inn)",
    "CREATE INDEX IF NOT EXISTS idx_counterparties_deleted ON counterparties(deleted_at)",
    """
    CREATE TABLE IF NOT EXISTS documents (
        id                 INTEGER PRIMARY KEY,
        kind               TEXT NOT NULL DEFAULT '',
        party_id           INTEGER,
        number             TEXT NOT NULL DEFAULT '',
        date               TEXT NOT NULL DEFAULT '',
        data               TEXT NOT NULL DEFAULT '{}',
        extra              TEXT NOT NULL DEFAULT '{}',
        text               TEXT NOT NULL DEFAULT '',
        created_at         TEXT NOT NULL DEFAULT '',
        deleted_at         TEXT,
        deleted_with_party INTEGER NOT NULL DEFAULT 0,
        extra_keys         TEXT NOT NULL DEFAULT '{}'
    )
    """,
    "CREATE INDEX IF NOT EXISTS idx_documents_party ON documents(party_id)",
    "CREATE INDEX IF NOT EXISTS idx_documents_kind ON documents(kind, party_id)",
    "CREATE INDEX IF NOT EXISTS idx_documents_deleted ON documents(deleted_at)",
    """
    CREATE TABLE IF NOT EXISTS meta (
        key   TEXT PRIMARY KEY,
        value TEXT NOT NULL
    )
    """,
)


def data_dir() -> str:
    """Каталог данных программы (создаётся при необходимости)."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, SETTINGS_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def db_path() -> str:
    return os.path.join(data_dir(), DB_FILENAME)


def legacy_json_path() -> str:
    """Старый JSON-реестр: читается один раз для переноса в базу."""
    return os.path.join(data_dir(), LEGACY_JSON_FILENAME)


def _connect_db(path: str) -> sqlite3.Connection:
    """Открывает базу: WAL (параллельное чтение), FULL-синхронизация, схема."""
    conn = sqlite3.connect(path, timeout=BUSY_TIMEOUT_SECONDS)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=FULL")
    for statement in _SCHEMA_SQL:
        conn.execute(statement)
    conn.commit()
    return conn


def _safe_int(value, default: int = 0) -> int:
    """Безопасно приводит значение к int (null/строка не роняют запуск)."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _now_iso() -> str:
    """Метка времени удаления с микросекундами — уникальна для отката."""
    return datetime.now().isoformat(timespec="microseconds")


def is_deleted(rec: dict) -> bool:
    """Запись находится в корзине."""
    return bool(rec.get("deleted_at"))


def active_records(records: list[dict]) -> list[dict]:
    return [rec for rec in records if not is_deleted(rec)]


class DataStore:
    """SQLite-хранилище контрагентов и документов.

    В памяти реестр выглядит как раньше — два списка словарей
    (``counterparties`` и ``documents``); на диск он целиком пишется
    одной транзакцией в database.db.
    """

    def __init__(self):
        self.counterparties: list[dict] = []
        self.documents: list[dict] = []
        self._seq = 1
        self.schema_version = SQLITE_SCHEMA_VERSION
        # Причина последней неудачной записи: UI обязан проверять её перед
        # сообщением «сохранено».
        self.last_error = ""
        # Предупреждение о проблемах при чтении файла (показывается пользователю).
        self.load_warning = ""
        self._load()

    # ------------------------------------------------------------------ #
    #  Загрузка / сохранение
    # ------------------------------------------------------------------ #
    def _load(self) -> None:
        log = applog.get_logger("storage")
        path = db_path()
        self._purge_stale_bad_copies(path)
        try:
            with applog.timed("Чтение реестра", 500, "storage"):
                conn = _connect_db(path)
                try:
                    parties, documents, seq = self._read_all(conn)
                finally:
                    conn.close()
        except sqlite3.DatabaseError as exc:
            if self._quarantine_corrupt_db(path, exc):
                return  # повреждённая база убрана в .bad, предупреждение выставлено
            self.load_warning = (
                "Не удалось прочитать реестр — программа запущена с пустым "
                f"реестром.\nОшибка: {exc}\nПуть: {path}"
            )
            log.error("Реестр не прочитан: %s", type(exc).__name__)
            return
        except OSError as exc:
            self.load_warning = (
                "Не удалось открыть базу реестра — программа запущена с "
                f"пустым реестром.\nОшибка: {exc}\nПуть: {path}"
            )
            log.error("База реестра не открыта: %s", type(exc).__name__)
            return

        if seq is None and not parties and not documents:
            # База пуста и ни разу не сохранялась: переносим старый JSON.
            self._migrate_from_legacy_json()
            return

        self.counterparties = parties
        self.documents = documents
        self._seq = max(
            [_safe_int(seq, 1)]
            + [_safe_int(rec.get("id"), 0) for rec in self.counterparties]
            + [_safe_int(rec.get("id"), 0) for rec in self.documents]
        ) + 1
        log.info(
            "Реестр загружен: контрагентов=%d, документов=%d, в корзине=%d",
            len(self.counterparties), len(self.documents),
            sum(1 for rec in self.counterparties + self.documents if is_deleted(rec)),
        )

    # ------------------- Перенос старого JSON-реестра ------------------ #
    def _migrate_from_legacy_json(self) -> None:
        """Переносит старый JSON-реестр в базу при первом запуске (п. 11)."""
        log = applog.get_logger("storage")
        legacy = self._read_legacy_json()
        if legacy is None:
            return  # файла нет или он не читается — предупреждение выставлено
        self.counterparties = legacy["counterparties"]
        self.documents = legacy["documents"]
        self._seq = legacy["seq"]
        # Старый JSON тоже несёт ключ «(Покупь)» — нормализуем до первой записи.
        for rec in self.counterparties:
            _normalize_field_names(rec)
        for rec in self.documents:
            _normalize_field_names(rec)
        if not self.save():
            self.load_warning = (
                "Реестр прочитан из старого файла database.json, но записать "
                "новую базу не удалось: " + (self.last_error or "неизвестная ошибка")
            )
            return  # JSON оставлен на месте, данные не потеряны
        json_path = legacy_json_path()
        try:
            os.replace(json_path, json_path + ".migrated")
        except OSError:
            log.warning("Старый JSON-реестр не переименован в .migrated")
        log.info(
            "Реестр перенесён из JSON в базу: контрагентов=%d, документов=%d",
            len(self.counterparties), len(self.documents),
        )

    def _read_legacy_json(self) -> dict | None:
        """Читает старый JSON-реестр. None — файла нет или он не читается."""
        log = applog.get_logger("storage")
        path = legacy_json_path()
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except FileNotFoundError:
            return None  # старого реестра нет — чистый старт
        except (OSError, ValueError) as exc:
            self.load_warning = (
                "Старый файл реестра database.json повреждён и не был "
                "перенесён в новую базу. Файл сохранён без изменений — "
                "скопируйте его, если хотите восстановить записи вручную.\n"
                f"Путь: {path}\nОшибка: {exc}"
            )
            log.error("Старый JSON-реестр не прочитан: %s", type(exc).__name__)
            return None

        if not isinstance(data, dict):
            self.load_warning = (
                "Старый файл реестра имеет неверную структуру и не был "
                f"перенесён в новую базу.\nПуть: {path}"
            )
            return None

        raw_parties = data.get("counterparties", [])
        raw_documents = data.get("documents", [])
        if not isinstance(raw_parties, list) or not isinstance(raw_documents, list):
            self.load_warning = (
                "Старый файл реестра повреждён: разделы «counterparties»/"
                "«documents» имеют неверный формат, перенос не выполнен.\n"
                f"Путь: {path}"
            )
            return None

        parties = [rec for rec in raw_parties if isinstance(rec, dict)]
        documents = [rec for rec in raw_documents if isinstance(rec, dict)]
        dropped = (len(raw_parties) - len(parties)) + (
            len(raw_documents) - len(documents)
        )
        if dropped:
            log.warning("При переносе реестра отброшено битых записей: %d", dropped)

        # Дубли id (файл могли править вручную): база не допустит второй
        # записи с тем же PRIMARY KEY, поэтому переносим первое вхождение.
        seen_ids: set[int] = set()
        parties = self._dedupe_by_id(parties, seen_ids)
        documents = self._dedupe_by_id(documents, seen_ids)

        seq = max(
            [_safe_int(data.get("seq"), 1)]
            + [_safe_int(rec.get("id"), 0) for rec in parties]
            + [_safe_int(rec.get("id"), 0) for rec in documents]
        ) + 1
        return {"counterparties": parties, "documents": documents, "seq": seq}

    @staticmethod
    def _dedupe_by_id(records: list[dict], seen_ids: set[int]) -> list[dict]:
        unique: list[dict] = []
        for rec in records:
            rid = _safe_int(rec.get("id"), 0)
            if rid and rid in seen_ids:
                applog.get_logger("storage").warning(
                    "Дубликат id=%s пропущен при переносе", rid,
                )
                continue
            if rid:
                seen_ids.add(rid)
            unique.append(rec)
        return unique

    # --------------------------- Чтение базы --------------------------- #
    @staticmethod
    def _loads(text, fallback):
        try:
            return json.loads(text or "")
        except ValueError:
            return fallback

    @classmethod
    def _party_from_row(cls, row) -> dict | None:
        fields = cls._loads(row[2], None)
        if not isinstance(fields, dict):
            return None
        rec: dict = {
            "id": row[0],
            "type": row[1] or "",
            "fields": fields,
            "created_at": row[3] or "",
        }
        if row[4]:
            rec["deleted_at"] = row[4]
        extra = cls._loads(row[5], None)
        if isinstance(extra, dict):
            for key, value in extra.items():
                if key not in _PARTY_KEYS:
                    rec[key] = value
        return rec

    @classmethod
    def _document_from_row(cls, row) -> dict | None:
        data = cls._loads(row[5], None)
        extra = cls._loads(row[6], None)
        if not isinstance(data, dict) or not isinstance(extra, dict):
            return None
        rec: dict = {
            "id": row[0],
            "kind": row[1] or "",
            "party_id": row[2],
            "number": row[3] or "",
            "date": row[4] or "",
            "data": data,
            "extra": extra,
            "text": row[7] or "",
            "created_at": row[8] or "",
        }
        if row[9]:
            rec["deleted_at"] = row[9]
        if row[10]:
            rec["deleted_with_party"] = True
        extra_keys = cls._loads(row[11], None)
        if isinstance(extra_keys, dict):
            for key, value in extra_keys.items():
                if key not in _DOCUMENT_KEYS:
                    rec[key] = value
        return rec

    @classmethod
    def _read_all(cls, conn) -> tuple[list[dict], list[dict], int | None]:
        parties: list[dict] = []
        documents: list[dict] = []
        dropped = 0
        for row in conn.execute(
            "SELECT id, type, fields, created_at, deleted_at, extra_keys "
            "FROM counterparties"
        ):
            rec = cls._party_from_row(row)
            if rec is None:
                dropped += 1
            else:
                _normalize_field_names(rec)
                parties.append(rec)
        for row in conn.execute(
            "SELECT id, kind, party_id, number, date, data, extra, text, "
            "created_at, deleted_at, deleted_with_party, extra_keys FROM documents"
        ):
            rec = cls._document_from_row(row)
            if rec is None:
                dropped += 1
            else:
                _normalize_field_names(rec)
                documents.append(rec)
        if dropped:
            applog.get_logger("storage").warning(
                "При чтении реестра отброшено битых записей: %d", dropped,
            )
        seq_row = conn.execute("SELECT value FROM meta WHERE key = 'seq'").fetchone()
        seq = _safe_int(seq_row[0], None) if seq_row else None
        return parties, documents, seq

    # ------------------------ Повреждённая база ------------------------ #
    def _quarantine_corrupt_db(self, path: str, exc: Exception) -> bool:
        """Повреждённую базу убирает в ``.bad-копию`` и предупреждает.

        Возвращает True, если ошибка похожа на повреждение файла и реестр
        начат с пустой базы. Блокировки и другие временные ошибки — False.
        """
        message = str(exc).lower()
        markers = ("not a database", "malformed", "encrypted", "corrupt")
        if not any(marker in message for marker in markers):
            return False
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        moved: list[str] = []
        failed: list[str] = []
        for suffix in ("", "-wal", "-shm"):
            src = path + suffix
            if not os.path.exists(src):
                continue
            destination = f"{path}.bad-{stamp}{suffix}"
            moved_ok = False
            last_exc: OSError | None = None
            # На Windows только что записанный файл может быть на короткое
            # время заблокирован антивирусным сканом (Defender): os.replace
            # падает с PermissionError. Замена ретраится с растущей паузой
            # (суммарно ~0,4 с) — операция разовая, старт приложения это не
            # тормозит. 2.15.0: без ретраев CI-сборка на Windows-раннере
            # стабильно теряла .bad-копию (файл занят сразу после записи).
            for attempt in range(5):
                try:
                    os.replace(src, destination)
                    moved_ok = True
                    break
                except OSError as replace_exc:
                    last_exc = replace_exc
                    time.sleep(0.02 * (attempt + 1))
            if moved_ok:
                moved.append(os.path.basename(src))
            else:
                failed.append(os.path.basename(src))
                if last_exc is not None:
                    applog.get_logger("storage").error(
                        "Не удалось убрать повреждённый файл %s: %s",
                        src, last_exc,
                    )
        self.load_warning = (
            "Файл базы реестра повреждён — программа запущена с пустым "
            f"реестром.\nОшибка: {exc}"
        )
        if moved:
            self.load_warning += (
                f"\nПовреждённый файл сохранён как:\n{path}.bad-{stamp}"
            )
        if failed:
            self.load_warning += (
                "\nНе удалось убрать повреждённые файлы: " + ", ".join(failed)
            )
        applog.get_logger("storage").error(
            "База реестра повреждена (%s); начат пустой реестр",
            type(exc).__name__,
        )
        return True

    def _purge_stale_bad_copies(self, path: str) -> None:
        """Удаляет ``.bad-*``-копии старше BAD_COPY_RETENTION_DAYS (аудит 2.0.1).

        Копии повреждённых файлов раньше копились без ограничения возраста
        и не попадали ни в какую очистку — теперь уборка идёт при каждом
        запуске, свежие копии не трогаются. Чистятся и старые
        ``database.json.bad-*`` от прежних версий программы.
        """
        directory = os.path.dirname(path)
        try:
            names = os.listdir(directory)
        except OSError:
            return
        prefixes = (
            os.path.basename(path) + ".bad-",
            LEGACY_JSON_FILENAME + ".bad-",
        )
        deadline = time.time() - BAD_COPY_RETENTION_DAYS * 86400
        removed = 0
        for name in names:
            if not name.startswith(prefixes):
                continue
            full = os.path.join(directory, name)
            try:
                if os.path.getmtime(full) < deadline:
                    os.remove(full)
                    removed += 1
            except OSError:
                continue
        if removed:
            applog.get_logger("storage").info(
                "Убраны старые копии повреждённых файлов: %d", removed,
            )

    # --------------------------- Запись базы --------------------------- #
    def _party_row(self, rec: dict) -> tuple:
        fields = rec.get("fields") or {}
        extra = {k: v for k, v in rec.items() if k not in _PARTY_KEYS}
        return (
            _safe_int(rec.get("id"), 0),
            str(rec.get("type") or ""),
            json.dumps(fields, ensure_ascii=False),
            str(rec.get("created_at") or ""),
            rec.get("deleted_at"),
            str(fields.get("ИНН", "") or "").strip(),
            json.dumps(extra, ensure_ascii=False),
        )

    def _document_row(self, rec: dict) -> tuple:
        extra = {k: v for k, v in rec.items() if k not in _DOCUMENT_KEYS}
        return (
            _safe_int(rec.get("id"), 0),
            str(rec.get("kind") or ""),
            rec.get("party_id"),
            str(rec.get("number") or ""),
            str(rec.get("date") or ""),
            json.dumps(rec.get("data") or {}, ensure_ascii=False),
            json.dumps(rec.get("extra") or {}, ensure_ascii=False),
            str(rec.get("text") or ""),
            str(rec.get("created_at") or ""),
            rec.get("deleted_at"),
            1 if rec.get("deleted_with_party") else 0,
            json.dumps(extra, ensure_ascii=False),
        )

    def _write_all(self, path: str, party_rows: list, document_rows: list) -> None:
        """Записывает весь реестр одной транзакцией: либо всё, либо ничего."""
        os.makedirs(os.path.dirname(path), exist_ok=True)
        conn = _connect_db(path)
        try:
            with conn:
                conn.execute("DELETE FROM counterparties")
                conn.execute("DELETE FROM documents")
                conn.execute("DELETE FROM meta")
                conn.executemany(
                    "INSERT INTO counterparties (id, type, fields, created_at, "
                    "deleted_at, inn, extra_keys) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    party_rows,
                )
                conn.executemany(
                    "INSERT INTO documents (id, kind, party_id, number, date, "
                    "data, extra, text, created_at, deleted_at, "
                    "deleted_with_party, extra_keys) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    document_rows,
                )
                conn.executemany(
                    "INSERT INTO meta (key, value) VALUES (?, ?)",
                    (
                        ("schema_version", str(SQLITE_SCHEMA_VERSION)),
                        ("seq", str(self._seq)),
                    ),
                )
        finally:
            conn.close()

    def save(self) -> bool:
        """Атомарно записывает реестр транзакцией. Возвращает False и заполняет
        last_error при сбое; предыдущая версия базы при этом не теряется.

        Долгая запись (более 500 мс — обычно антивирус или медленный диск)
        попадает в журнал (п. 10 этапа 5).
        """
        log = applog.get_logger("storage")
        self.last_error = ""
        # Сериализация до открытия базы: битые данные не роняют файл.
        try:
            party_rows = [self._party_row(rec) for rec in self.counterparties]
            document_rows = [self._document_row(rec) for rec in self.documents]
        except (TypeError, ValueError) as exc:
            self.last_error = f"Данные реестра не удалось сериализовать: {exc}"
            log.error("Сериализация реестра не удалась: %s", type(exc).__name__)
            return False

        path = db_path()
        last_exc: Exception | None = None
        with applog.timed("Запись реестра", 500, "storage"):
            for attempt in range(1, SAVE_RETRIES + 1):
                try:
                    self._write_all(path, party_rows, document_rows)
                    self.schema_version = SQLITE_SCHEMA_VERSION
                    log.info(
                        "Реестр записан: контрагентов=%d, документов=%d",
                        len(self.counterparties), len(self.documents),
                    )
                    return True
                except (sqlite3.Error, OSError) as exc:
                    last_exc = exc
                    if attempt < SAVE_RETRIES:
                        log.warning(
                            "Запись реестра не удалась (попытка %d из %d): %s",
                            attempt, SAVE_RETRIES, type(exc).__name__,
                        )
                        time.sleep(RETRY_DELAY_SECONDS * attempt)

        self.last_error = f"Не удалось записать реестр: {last_exc}"
        log.error("Реестр не записан после %d попыток: %s", SAVE_RETRIES, last_exc)
        return False

    def _next_id(self) -> int:
        next_id = self._seq
        self._seq += 1
        return next_id

    # ------------------------------------------------------------------ #
    #  Контрагенты
    # ------------------------------------------------------------------ #
    def add_counterparty(self, ptype: str, fields: dict, save: bool = True) -> dict:
        rec = {
            "id": self._next_id(),
            "type": ptype,  # "ooo" | "ip"
            "fields": dict(fields),
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }
        # Защита от старых callers, передающих ключ «(Покупь)» (v2.13.0).
        _normalize_field_names(rec)
        self.counterparties.append(rec)
        if save:
            self.save()
        return rec

    def active_counterparties(self) -> list[dict]:
        return active_records(self.counterparties)

    def trashed_counterparties(self) -> list[dict]:
        return [rec for rec in self.counterparties if is_deleted(rec)]

    def get_counterparty(self, party_id, include_deleted: bool = False) -> dict | None:
        for rec in self.counterparties:
            if rec.get("id") == party_id:
                if not include_deleted and is_deleted(rec):
                    return None
                return rec
        return None

    def update_counterparty(self, party_id, fields: dict) -> dict | None:
        """Обновляет карточку контрагента и сохраняет реестр.

        Документы контрагента намеренно не изменяются: каждый документ
        хранит собственную копию реквизитов на момент создания.
        """
        rec = self.get_counterparty(party_id)
        if rec is None:
            return None
        previous = rec.get("fields")
        rec["fields"] = dict(fields)
        # Защита от старых callers, передающих ключ «(Покупь)» (v2.13.0).
        _normalize_field_names(rec)
        if not self.save():
            rec["fields"] = previous
            return None
        return rec

    def find_by_inn(self, inn: str) -> dict | None:
        inn = (inn or "").strip()
        if not inn:
            return None
        for rec in self.counterparties:
            if is_deleted(rec):
                continue
            if str(rec.get("fields", {}).get("ИНН", "")).strip() == inn:
                return rec
        return None

    @staticmethod
    def _normalize_name(value) -> str:
        """Название для сравнения: без регистра, кавычек и лишних пробелов."""
        text = str(value or "").lower()
        text = re.sub(r"[«»\"“”„']", "", text)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    def find_by_name(self, name: str) -> dict | None:
        """Поиск контрагента по названию (п. 3.3).

        Нужен, когда в импортированном документе нет ИНН: если название уже
        есть в реестре, документ привязывается к существующей карточке,
        а не создаёт «полукопию».
        """
        target = self._normalize_name(name)
        if not target:
            return None
        keys = (
            "Сокращенное название компании",
            F.FULL_NAME,
            "ФИО ИП",
        )
        for rec in self.counterparties:
            if is_deleted(rec):
                continue
            fields = rec.get("fields") or {}
            for key in keys:
                if self._normalize_name(fields.get(key)) == target:
                    return rec
        return None

    # ------------------------------------------------------------------ #
    #  Документы
    # ------------------------------------------------------------------ #
    def add_document(
        self,
        kind: str,
        party_id,
        number: str,
        date_text: str,
        data: dict,
        extra: dict | None = None,
        text: str = "",
        save: bool = True,
    ) -> dict:
        rec = {
            "id": self._next_id(),
            "kind": kind,
            "party_id": party_id,
            "number": str(number or ""),
            "date": str(date_text or ""),
            "data": dict(data or {}),
            "extra": dict(extra or {}),
            "text": str(text or ""),
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.documents.append(rec)
        if save:
            self.save()
        return rec

    def active_documents(self) -> list[dict]:
        return active_records(self.documents)

    def trashed_documents(self) -> list[dict]:
        return [rec for rec in self.documents if is_deleted(rec)]

    def get_document(self, doc_id, include_deleted: bool = False) -> dict | None:
        for rec in self.documents:
            if rec.get("id") == doc_id:
                if not include_deleted and is_deleted(rec):
                    return None
                return rec
        return None

    def delete_document(self, doc_id) -> bool:
        """Окончательное (физическое) удаление записи документа."""
        before = len(self.documents)
        self.documents = [rec for rec in self.documents if rec.get("id") != doc_id]
        removed = len(self.documents) < before
        if removed:
            return self.save()
        return False

    def delete_counterparty(self, party_id) -> int:
        """Окончательно удаляет контрагента вместе со всеми его документами.

        Возвращает число удалённых документов; изменения сразу сохраняются.
        """
        self.counterparties = [
            rec for rec in self.counterparties if rec.get("id") != party_id
        ]
        before = len(self.documents)
        self.documents = [
            rec for rec in self.documents if rec.get("party_id") != party_id
        ]
        removed = before - len(self.documents)
        self.save()
        return removed

    def documents_for(self, party_id, include_deleted: bool = False) -> list[dict]:
        return [
            rec for rec in self.documents
            if rec.get("party_id") == party_id
            and (include_deleted or not is_deleted(rec))
        ]

    def contracts_for(self, party_id) -> list[dict]:
        """Договоры контрагента в порядке создания (без корзины)."""
        return [
            rec
            for rec in self.documents
            if rec.get("party_id") == party_id
            and rec.get("kind") == KIND_CONTRACT
            and not is_deleted(rec)
        ]

    def party_name(self, party_id) -> str:
        rec = self.get_counterparty(party_id, include_deleted=True)
        return display_name(rec) if rec else "—"

    # ------------------------------------------------------------------ #
    #  Корзина
    # ------------------------------------------------------------------ #
    def trash_document(self, doc_id) -> bool:
        """Перемещает документ в корзину (мягкое удаление)."""
        rec = self.get_document(doc_id, include_deleted=True)
        if rec is None or is_deleted(rec):
            return False
        rec["deleted_at"] = _now_iso()
        if not self.save():
            rec.pop("deleted_at", None)
            return False
        applog.get_logger("storage").info("Документ перемещён в корзину: id=%s", doc_id)
        return True

    def trash_counterparty(self, party_id) -> int:
        """Перемещает карточку и её документы в корзину. Возвращает число документов."""
        party = self.get_counterparty(party_id, include_deleted=True)
        if party is None or is_deleted(party):
            return 0
        stamp = _now_iso()
        party["deleted_at"] = stamp
        touched: list[dict] = []
        for rec in self.documents:
            if rec.get("party_id") == party_id and not is_deleted(rec):
                rec["deleted_at"] = stamp
                rec["deleted_with_party"] = True
                touched.append(rec)
        if not self.save():
            party.pop("deleted_at", None)
            for rec in touched:
                rec.pop("deleted_at", None)
                rec.pop("deleted_with_party", None)
            return 0
        applog.get_logger("storage").info(
            "Контрагент перемещён в корзину: id=%s, документов=%d", party_id, len(touched),
        )
        return len(touched)

    def restore_document(self, doc_id) -> bool:
        """Возвращает документ из корзины."""
        rec = self.get_document(doc_id, include_deleted=True)
        if rec is None or not is_deleted(rec):
            return False
        stamp = rec.pop("deleted_at")
        with_party = rec.pop("deleted_with_party", None)
        if not self.save():
            rec["deleted_at"] = stamp
            if with_party is not None:
                rec["deleted_with_party"] = with_party
            return False
        applog.get_logger("storage").info("Документ восстановлен: id=%s", doc_id)
        return True

    def restore_counterparty(self, party_id) -> int:
        """Возвращает карточку и документы, удалённые вместе с ней."""
        party = self.get_counterparty(party_id, include_deleted=True)
        if party is None or not is_deleted(party):
            return 0
        stamp = party.pop("deleted_at")
        restored: list[dict] = []
        for rec in self.documents:
            if (
                rec.get("party_id") == party_id
                and rec.get("deleted_with_party")
                and is_deleted(rec)
            ):
                rec.pop("deleted_at", None)
                rec.pop("deleted_with_party", None)
                restored.append(rec)
        if not self.save():
            party["deleted_at"] = stamp
            for rec in restored:
                rec["deleted_at"] = stamp
                rec["deleted_with_party"] = True
            return 0
        applog.get_logger("storage").info(
            "Контрагент восстановлен: id=%s, документов=%d", party_id, len(restored),
        )
        return len(restored)

    def purge_document(self, doc_id) -> bool:
        """Удаляет запись документа из корзины окончательно."""
        before = len(self.documents)
        self.documents = [rec for rec in self.documents if rec.get("id") != doc_id]
        if len(self.documents) == before:
            return False
        return self.save()

    def purge_counterparty(self, party_id) -> int:
        """Удаляет карточку и все её документы окончательно."""
        before = len(self.documents)
        self.counterparties = [
            rec for rec in self.counterparties if rec.get("id") != party_id
        ]
        self.documents = [
            rec for rec in self.documents if rec.get("party_id") != party_id
        ]
        removed = before - len(self.documents)
        self.save()
        return removed

    def trash_count(self) -> int:
        return sum(
            1 for rec in self.counterparties + self.documents if is_deleted(rec)
        )

    def purge_all_trash(self) -> int:
        """Очищает корзину полностью. Возвращает число удалённых записей."""
        parties = sum(1 for rec in self.counterparties if is_deleted(rec))
        docs = sum(1 for rec in self.documents if is_deleted(rec))
        if not parties and not docs:
            return 0
        self.counterparties = [rec for rec in self.counterparties if not is_deleted(rec)]
        self.documents = [rec for rec in self.documents if not is_deleted(rec)]
        if not self.save():
            return 0
        applog.get_logger("storage").info(
            "Корзина очищена: контрагентов=%d, документов=%d", parties, docs,
        )
        return parties + docs

    def cleanup_trash(self, retention_days: int = TRASH_RETENTION_DAYS) -> int:
        """Удаляет записи, пролежавшие в корзине дольше срока хранения."""
        cutoff = datetime.now() - timedelta(days=max(1, retention_days))

        def expired(rec: dict) -> bool:
            stamp = rec.get("deleted_at")
            if not stamp:
                return False
            try:
                return datetime.fromisoformat(str(stamp)) < cutoff
            except (TypeError, ValueError):
                return False

        parties = [rec for rec in self.counterparties if expired(rec)]
        docs = [rec for rec in self.documents if expired(rec)]
        if not parties and not docs:
            return 0
        expired_party_ids = {rec.get("id") for rec in parties}
        self.counterparties = [rec for rec in self.counterparties if not expired(rec)]
        self.documents = [
            rec for rec in self.documents
            if not expired(rec) and rec.get("party_id") not in expired_party_ids
        ]
        removed = len(parties) + len(docs)
        if not self.save():
            return 0
        applog.get_logger("storage").info(
            "Автоочистка корзины: удалено записей=%d (срок %d дн.)",
            removed, retention_days,
        )
        return removed


def display_name(rec: dict) -> str:
    """Короткое имя контрагента для списков: «ООО Ромашка» / «ИП Иванов»."""
    fields = rec.get("fields", {})
    if rec.get("type") == "ip":
        return f"ИП {fields.get('ФИО ИП', '').strip()}".strip()
    return fields.get("Сокращенное название компании", "").strip() or \
        fields.get(F.FULL_NAME, "").strip() or "Без названия"


def display_with_inn(rec: dict) -> str:
    """Имя для выпадающих списков: название + ИНН."""
    inn = str(rec.get("fields", {}).get("ИНН", "")).strip()
    name = display_name(rec)
    return f"{name} · ИНН {inn}" if inn else name
