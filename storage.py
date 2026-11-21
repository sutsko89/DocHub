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
import shutil
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


def _rename_field_key(payload: dict) -> bool:
    """Переименовывает один ключ в словаре полей: старый → канонический.

    Если канонический ключ уже заполнен (обе версии в одной записи),
    остаётся новое значение — старое отбрасывается, чтобы предпросмотр
    и печатная форма не увидели расхождения. Возвращает True, если
    словарь изменился (STORE-1A: загрузка помечает такую запись dirty,
    чтобы нормализация доезжала до диска).
    """
    if LEGACY_FULL_NAME_FIELD not in payload:
        return False
    value = payload.pop(LEGACY_FULL_NAME_FIELD)
    payload.setdefault(F.FULL_NAME, value)
    return True


def _normalize_field_names(rec: dict) -> bool:
    """Миграция имён полей v2.13.0 для записи реестра (идемпотентно).

    Затрагивает карточку контрагента (``fields``) и сохранённые данные
    документа (``data`` — копия полей карточки на момент сохранения).
    Возвращает True, если запись изменилась (STORE-1A: загрузка помечает
    её dirty — нормализация доезжает до диска следующей save()).
    """
    changed = False
    fields = rec.get("fields")
    if isinstance(fields, dict):
        changed = _rename_field_key(fields) or changed
    data = rec.get("data")
    if isinstance(data, dict):
        changed = _rename_field_key(data) or changed
    return changed


SQLITE_SCHEMA_VERSION = 1
SCHEMA_VERSION = 2              # схема старого JSON-реестра (для переноса)
BUSY_TIMEOUT_SECONDS = 5.0      # ждём освобождения базы вместо мгновенного сбоя
SAVE_RETRIES = 3
RETRY_DELAY_SECONDS = 0.25
# Ретраи чтения при временной блокировке файла базы (антивирус/бэкап).
LOAD_RETRIES = 5
LOAD_RETRY_DELAY = 0.25
TRASH_RETENTION_DAYS = 30
BAD_COPY_RETENTION_DAYS = 30   # сколько дней хранить database.db.bad-*

# Автокопии реестра (BKUP-1): ротация 7 копий + отдельные снимки перед
# обновлением (до 3); снимки не вытесняют обычные копии и в лимит 7 не входят.
SETTINGS_FILENAME = "settings.json"
BACKUPS_DIR_NAME = "backups"
SNAPSHOTS_DIR_NAME = "snapshots"
BACKUP_FILE_PREFIX = "registry_"
BACKUP_MARKER_FILENAME = "last_backup.json"
BACKUP_SETTINGS_SUFFIX = "_settings.json"
BACKUP_RETENTION = 7
SNAPSHOT_RETENTION = 3
# Ключ — секрет пользователя: в копию настроек не включается никогда
# (включение в копию «по явному выбору» не предусмотрено, ключ хранится
# отдельным файлом и в settings.json не пишется).
_SECRET_KEY_MARKERS = ("dadata",)

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
    # STORE-1A (R-05a): покрывающий индекс для COUNT(*) статусной строки —
    # подсчёт активных документов и договоров идёт index-only сканом
    # (deleted_at ведущий колонкой, kind — внутри индекса).
    "CREATE INDEX IF NOT EXISTS idx_documents_status ON documents(deleted_at, kind)",
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


def settings_json_path() -> str:
    """Файл настроек программы (тот же путь, что в main.settings_path)."""
    return os.path.join(data_dir(), SETTINGS_FILENAME)


def backups_dir() -> str:
    """Каталог автокопий реестра (BKUP-1): <профиль>/backups/."""
    return os.path.join(data_dir(), BACKUPS_DIR_NAME)


def snapshots_dir() -> str:
    """Каталог снимков перед обновлением: backups/snapshots/ (своя ротация)."""
    return os.path.join(backups_dir(), SNAPSHOTS_DIR_NAME)


def backup_marker_path() -> str:
    """Отметка о последней успешной копии (ревизия реестра на момент копии)."""
    return os.path.join(backups_dir(), BACKUP_MARKER_FILENAME)


def _connect_db(path: str) -> sqlite3.Connection:
    """Открывает базу: WAL (параллельное чтение), FULL-синхронизация, схема.

    При ошибке инициализации (например, «file is not a database» на
    PRAGMA) соединение закрывается ДО выхода из функции: иначе объект
    Connection живёт в traceback исключения, держит дескриптор файла и
    os.replace/os.remove в quarantine на Windows падают с WinError 32
    («файл занят другим процессом») — битая база оставалась на месте и
    quarantine повторялся при каждом запуске (2.15.0).
    """
    conn = sqlite3.connect(path, timeout=BUSY_TIMEOUT_SECONDS)
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=FULL")
        for statement in _SCHEMA_SQL:
            conn.execute(statement)
        conn.commit()
        return conn
    except BaseException:
        try:
            conn.close()
        except Exception:
            pass
        raise


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
        # Ревизия реестра: увеличивается каждой успешной записью и хранится
        # в meta — триггер автокопии «изменено с последней копии» (BKUP-1)
        # сравнивает ревизию базы с отметкой последней копии.
        self._rev = 0
        # Флаг «реестр менялся с последней успешной копии»: ставится каждым
        # успешным save(), снимается успешной копией (BKUP-1).
        self.dirty_since_backup = False
        # STORE-1A (R-01): dirty-наборы инкрементальной записи. Мутаторы
        # (add/update/trash/restore/purge…) помечают id, save() пишет
        # ТОЛЬКО эти строки (INSERT OR REPLACE) и удаляет перечисленные
        # (DELETE). Наборы очищаются только после подтверждения транзакции:
        # сбой записи оставляет изменения «грязными», повторный save()
        # их дописывает. Полный путь (перезапись всех строк) — save_full():
        # откат и самопроверка, перенос старого JSON.
        self._dirty_parties: set[int] = set()
        self._dirty_documents: set[int] = set()
        self._deleted_party_ids: set[int] = set()
        self._deleted_document_ids: set[int] = set()
        # Причина последней неудачной записи: UI обязан проверять её перед
        # сообщением «сохранено».
        self.last_error = ""
        # Предупреждение о проблемах при чтении файла (показывается пользователю).
        self.load_warning = ""
        self._load()

    @property
    def registry_rev(self) -> int:
        """Текущая ревизия реестра (количество успешных записей в базе)."""
        return self._rev

    # ------------------------------------------------------------------ #
    #  Загрузка / сохранение
    # ------------------------------------------------------------------ #
    def _load(self) -> None:
        log = applog.get_logger("storage")
        path = db_path()
        self._purge_stale_bad_copies(path)
        # На Windows свежесозданный файл базы может быть несколько секунд
        # заблокирован антивирусным сканом: sqlite3.connect падает с
        # OperationalError «unable to open database file» / «locked». Это
        # временная ошибка — чтение повторяется с растущей паузой (2.15.0:
        # без ретраев CI-сборка на Windows-раннере получала пустой реестр
        # вместо миграции, а пользователь мог увидеть пустой реестр при
        # запуске сразу после установки).
        try:
            for attempt in range(1, LOAD_RETRIES + 1):
                try:
                    with applog.timed("Чтение реестра", 500, "storage"):
                        conn = _connect_db(path)
                        try:
                            parties, documents, seq, rev, migrated_ids = self._read_all(conn)
                        finally:
                            conn.close()
                    break
                except sqlite3.OperationalError as exc:
                    message = str(exc).lower()
                    transient = (
                        "unable to open" in message
                        or "locked" in message
                        or "busy" in message
                    )
                    if not transient or attempt >= LOAD_RETRIES:
                        raise
                    log.warning(
                        "Чтение реестра повторяется (попытка %d из %d): %s",
                        attempt, LOAD_RETRIES, type(exc).__name__,
                    )
                    time.sleep(LOAD_RETRY_DELAY * attempt)
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
        self._rev = _safe_int(rev, 0)
        # STORE-1A (R-01): записи, нормализованные при загрузке (миграция
        # имён полей v2.13.0), помечаются dirty — их перезапишет следующая
        # save(), и старый ключ уйдёт с диска без полной перезаписи.
        for rec in self.counterparties:
            if _safe_int(rec.get("id"), 0) in migrated_ids:
                self._mark_party_dirty(rec)
        for rec in self.documents:
            if _safe_int(rec.get("id"), 0) in migrated_ids:
                self._mark_document_dirty(rec)
        if migrated_ids:
            log.info(
                "Миграция имён полей при загрузке: записей=%d — будут "
                "перезаписаны при ближайшем сохранении", len(migrated_ids),
            )
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
        # STORE-1A: полный путь — перенос идёт мимо мутаторов, dirty-наборы
        # пусты; save() посчитал бы реестр «не изменённым» и ничего не записал.
        if not self.save_full():
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
    def _read_all(cls, conn) -> tuple[list[dict], list[dict], int | None, int | None, set[int]]:
        parties: list[dict] = []
        documents: list[dict] = []
        migrated_ids: set[int] = set()
        dropped = 0
        for row in conn.execute(
            "SELECT id, type, fields, created_at, deleted_at, extra_keys "
            "FROM counterparties"
        ):
            rec = cls._party_from_row(row)
            if rec is None:
                dropped += 1
            else:
                # STORE-1A: миграция имён полей при загрузке помечает запись
                # dirty — следующая save() перезапишет только её и старый
                # ключ уйдёт с диска (раньше полный save() переписывал всё).
                if _normalize_field_names(rec):
                    migrated_ids.add(_safe_int(rec.get("id"), 0))
                parties.append(rec)
        for row in conn.execute(
            "SELECT id, kind, party_id, number, date, data, extra, text, "
            "created_at, deleted_at, deleted_with_party, extra_keys FROM documents"
        ):
            rec = cls._document_from_row(row)
            if rec is None:
                dropped += 1
            else:
                if _normalize_field_names(rec):
                    migrated_ids.add(_safe_int(rec.get("id"), 0))
                documents.append(rec)
        if dropped:
            applog.get_logger("storage").warning(
                "При чтении реестра отброшено битых записей: %d", dropped,
            )
        seq_row = conn.execute("SELECT value FROM meta WHERE key = 'seq'").fetchone()
        seq = _safe_int(seq_row[0], None) if seq_row else None
        rev_row = conn.execute("SELECT value FROM meta WHERE key = 'rev'").fetchone()
        rev = _safe_int(rev_row[0], None) if rev_row else None
        return parties, documents, seq, rev, migrated_ids

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
            # На Windows только что записанный файл может быть несколько
            # секунд заблокирован антивирусным сканом (Defender): os.replace
            # падает с PermissionError / WinError 32. Замена ретраится с
            # растущей паузой (суммарно ~2,75 с) — операция разовая, старт
            # приложения это не тормозит. 2.15.0: без ретраев CI-сборка на
            # Windows-раннере стабильно теряла .bad-копию (файл занят сразу
            # после записи).
            for attempt in range(10):
                try:
                    os.replace(src, destination)
                    moved_ok = True
                    break
                except OSError as replace_exc:
                    last_exc = replace_exc
                    time.sleep(0.05 * (attempt + 1))
            if not moved_ok:
                # Последний шанс: чтение не требует эксклюзивной блокировки —
                # копируем содержимое в .bad-копию и удаляем оригинал.
                # moved_ok ставится только при УСПЕШНОМ удалении: если битый
                # файл остался на месте базы, quarantine повторится при
                # следующем запуске (потери данных нет, копия уже создана).
                try:
                    shutil.copyfile(src, destination)
                    try:
                        os.remove(src)
                        moved_ok = True
                    except OSError as remove_exc:
                        last_exc = remove_exc
                except OSError as copy_exc:
                    last_exc = copy_exc
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
                        # Ревизия каждой успешной записи — триггер автокопий (BKUP-1).
                        # Схема не меняется: старые версии читают meta без этой строки.
                        ("rev", str(self._rev)),
                    ),
                )
        finally:
            conn.close()

    def _has_dirty(self) -> bool:
        """Есть ли неподтверждённые изменения (R-01)."""
        return bool(
            self._dirty_parties or self._dirty_documents
            or self._deleted_party_ids or self._deleted_document_ids
        )

    def _mark_party_dirty(self, rec: dict) -> None:
        rec_id = _safe_int(rec.get("id"), 0)
        if rec_id > 0:
            self._dirty_parties.add(rec_id)

    def _mark_document_dirty(self, rec: dict) -> None:
        rec_id = _safe_int(rec.get("id"), 0)
        if rec_id > 0:
            self._dirty_documents.add(rec_id)

    def _mark_party_deleted(self, rec_id) -> None:
        rec_id = _safe_int(rec_id, 0)
        if rec_id > 0:
            self._deleted_party_ids.add(rec_id)
            self._dirty_parties.discard(rec_id)

    def _mark_document_deleted(self, rec_id) -> None:
        rec_id = _safe_int(rec_id, 0)
        if rec_id > 0:
            self._deleted_document_ids.add(rec_id)
            self._dirty_documents.discard(rec_id)

    def _clear_dirty(self) -> None:
        """Сброс dirty-наборов — ТОЛЬКО после подтверждения транзакции (R-01)."""
        self._dirty_parties.clear()
        self._dirty_documents.clear()
        self._deleted_party_ids.clear()
        self._deleted_document_ids.clear()

    def mark_dirty(self, rec: dict) -> None:
        """Публичная пометка записи «изменена мимо мутатора» (R-01).

        Для редких случаев прямого изменения словаря записи вне add/update-
        методов (миграция при загрузке, сценарии восстановления): следующая
        save() перезапишет только эту строку. Полная перезапись — save_full().
        """
        # Сравнение по тождеству (is): dict.__eq__ для больших карточек
        # дорог и может упасть на нестандартном содержимом.
        if any(rec is item for item in self.documents):
            self._mark_document_dirty(rec)
        else:
            self._mark_party_dirty(rec)

    def _collect_dirty_changes(self):
        """Готовит точечные изменения (R-01). None — сбой сериализации.

        Dirty id без записи в памяти — «фантом», удалённый из памяти вне
        мутатора (services.save_document снимает запись при сбое сохранения,
        counterparty_controller убирает фантом после сбоя add): такой id
        становится DELETE — точечная запись не воскресит строку.
        """
        log = applog.get_logger("storage")
        try:
            party_index = {_safe_int(rec.get("id"), 0): rec for rec in self.counterparties}
            document_index = {_safe_int(rec.get("id"), 0): rec for rec in self.documents}
            for rec_id in list(self._dirty_parties):
                if rec_id not in party_index:
                    self._dirty_parties.discard(rec_id)
                    self._deleted_party_ids.add(rec_id)
            for rec_id in list(self._dirty_documents):
                if rec_id not in document_index:
                    self._dirty_documents.discard(rec_id)
                    self._deleted_document_ids.add(rec_id)
            party_rows = [self._party_row(party_index[rec_id]) for rec_id in self._dirty_parties]
            document_rows = [
                self._document_row(document_index[rec_id]) for rec_id in self._dirty_documents
            ]
            return (
                party_rows, document_rows,
                set(self._deleted_party_ids), set(self._deleted_document_ids),
            )
        except (TypeError, ValueError) as exc:
            self.last_error = f"Данные реестра не удалось сериализовать: {exc}"
            log.error("Сериализация реестра не удалась: %s", type(exc).__name__)
            return None

    def _write_incremental(self, path: str, party_rows: list, document_rows: list,
                           deleted_party_ids: set, deleted_document_ids: set) -> None:
        """Точечная запись dirty-строк одной транзакцией: либо всё, либо ничего (R-01).

        Порядок — сначала DELETE, затем INSERT OR REPLACE: даже если бы id
        одновременно попал в удалённые и изменённые, финальное состояние
        базы соответствует памяти. seq обновляется при добавлении записи
        (мутатор уже сдвинул _seq), rev — каждой успешной записью (триггер
        автокопий BKUP-1). Остальные ключи meta не трогаются.
        """
        os.makedirs(os.path.dirname(path), exist_ok=True)
        conn = _connect_db(path)
        try:
            with conn:
                if deleted_party_ids:
                    conn.executemany(
                        "DELETE FROM counterparties WHERE id = ?",
                        [(rec_id,) for rec_id in deleted_party_ids],
                    )
                if deleted_document_ids:
                    conn.executemany(
                        "DELETE FROM documents WHERE id = ?",
                        [(rec_id,) for rec_id in deleted_document_ids],
                    )
                if party_rows:
                    conn.executemany(
                        "INSERT OR REPLACE INTO counterparties "
                        "(id, type, fields, created_at, deleted_at, inn, extra_keys) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        party_rows,
                    )
                if document_rows:
                    conn.executemany(
                        "INSERT OR REPLACE INTO documents "
                        "(id, kind, party_id, number, date, data, extra, text, "
                        "created_at, deleted_at, deleted_with_party, extra_keys) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        document_rows,
                    )
                conn.executemany(
                    "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
                    (
                        ("schema_version", str(SQLITE_SCHEMA_VERSION)),
                        ("seq", str(self._seq)),
                        # Ревизия каждой успешной записи — триггер автокопий (BKUP-1).
                        ("rev", str(self._rev)),
                    ),
                )
        finally:
            conn.close()

    def save(self) -> bool:
        """Инкрементальная запись реестра (STORE-1A, R-01).

        Пишет только изменённые строки (INSERT OR REPLACE по изменённым id)
        и удаляет изъятые (DELETE по id) — одной транзакцией. Мутаторы
        помечают id в dirty-наборах; наборы очищаются ТОЛЬКО после
        подтверждения транзакции, поэтому сбой записи не теряет изменения.
        При отсутствии изменений запись не выполняется вовсе: ревизия
        не растёт, автокопия BKUP-1 не запускается.

        Полный путь (перезапись всех строк) остаётся в save_full(): откат
        и самопроверка, перенос старого JSON. Ошибка записи доходит до
        last_error и UI; длительная запись (>500 мс) — в журнал.
        """
        log = applog.get_logger("storage")
        self.last_error = ""
        if not self._has_dirty():
            log.info("Реестр не изменён с последней записи — сохранение не требуется")
            return True

        changes = self._collect_dirty_changes()
        if changes is None:
            return False
        party_rows, document_rows, deleted_party_ids, deleted_document_ids = changes

        path = db_path()
        last_exc: Exception | None = None
        # Ревизия увеличивается каждой УСПЕШНОЙ записью (сбой — откат счётчика).
        self._rev += 1
        with applog.timed("Запись реестра (инкрементально)", 500, "storage"):
            for attempt in range(1, SAVE_RETRIES + 1):
                try:
                    self._write_incremental(
                        path, party_rows, document_rows,
                        deleted_party_ids, deleted_document_ids,
                    )
                    self.schema_version = SQLITE_SCHEMA_VERSION
                    self.dirty_since_backup = True  # BKUP-1: «изменено с последней копии»
                    log.info(
                        "Реестр записан инкрементально: записано контрагентов=%d, "
                        "документов=%d, удалено контрагентов=%d, документов=%d",
                        len(party_rows), len(document_rows),
                        len(deleted_party_ids), len(deleted_document_ids),
                    )
                    self._clear_dirty()
                    return True
                except (sqlite3.Error, OSError) as exc:
                    last_exc = exc
                    if attempt < SAVE_RETRIES:
                        log.warning(
                            "Запись реестра не удалась (попытка %d из %d): %s",
                            attempt, SAVE_RETRIES, type(exc).__name__,
                        )
                        time.sleep(RETRY_DELAY_SECONDS * attempt)

        self._rev -= 1  # ревизия не записана в базу — счётчик откатывается
        self.last_error = f"Не удалось записать реестр: {last_exc}"
        # SEC-1 (D12): сообщение sqlite содержит полный путь базы —
        # в журнале корень профиля маскируется.
        log.error(
            "Реестр не записан после %d попыток: %s",
            SAVE_RETRIES, applog.mask_path(str(last_exc)),
        )
        return False

    def save_full(self) -> bool:
        """Полная перезапись реестра одной транзакцией: либо всё, либо ничего.

        Путь отката и самопроверки (STORE-1A): используется переносом старого
        JSON (_migrate_from_legacy_json) и доступна вызывающему коду для
        принудительной сверкой памяти и диска. Повседневная запись — save().
        Долгая запись (>500 мс — обычно антивирус или медленный диск)
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
        # Ревизия увеличивается каждой УСПЕШНОЙ записью (сбой — откат счётчика).
        self._rev += 1
        with applog.timed("Запись реестра", 500, "storage"):
            for attempt in range(1, SAVE_RETRIES + 1):
                try:
                    self._write_all(path, party_rows, document_rows)
                    self.schema_version = SQLITE_SCHEMA_VERSION
                    self.dirty_since_backup = True  # BKUP-1: «изменено с последней копии»
                    log.info(
                        "Реестр записан полностью: контрагентов=%d, документов=%d",
                        len(self.counterparties), len(self.documents),
                    )
                    # Полная запись подтверждает ВСЕ изменения памяти.
                    self._clear_dirty()
                    return True
                except (sqlite3.Error, OSError) as exc:
                    last_exc = exc
                    if attempt < SAVE_RETRIES:
                        log.warning(
                            "Запись реестра не удалась (попытка %d из %d): %s",
                            attempt, SAVE_RETRIES, type(exc).__name__,
                        )
                        time.sleep(RETRY_DELAY_SECONDS * attempt)

        self._rev -= 1  # ревизия не записана в базу — счётчик откатывается
        self.last_error = f"Не удалось записать реестр: {last_exc}"
        # SEC-1 (D12): сообщение sqlite содержит полный путь базы —
        # в журнале корень профиля маскируется.
        log.error(
            "Реестр не записан после %d попыток: %s",
            SAVE_RETRIES, applog.mask_path(str(last_exc)),
        )
        return False

    def backup_db(self, dest: str, retries: int = 3,
                  retry_delay: float = 0.25) -> bool:
        """Консистентная копия базы через sqlite backup API (BKUP-1).

        Backup API переносит и содержимое WAL: копия соответствует состоянию
        базы на момент вызова, останавливать программу не нужно. Ретраи —
        против секундной блокировки файла антивирусом (как в _load/save).
        Ошибка копирования не поднимается наружу: возвращается False,
        подробности — в журнале.
        """
        log = applog.get_logger("storage")
        source_path = db_path()
        if not os.path.exists(source_path):
            log.warning("Копия реестра: файл базы ещё не создан — копировать нечего")
            return False
        last_exc: Exception | None = None
        for attempt in range(1, retries + 1):
            try:
                os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
                source = sqlite3.connect(source_path, timeout=BUSY_TIMEOUT_SECONDS)
                try:
                    target = sqlite3.connect(dest)
                    try:
                        source.backup(target)
                    finally:
                        target.close()
                finally:
                    source.close()
                return True
            except (sqlite3.Error, OSError) as exc:
                last_exc = exc
                if attempt < retries:
                    log.warning(
                        "Копия реестра не удалась (попытка %d из %d): %s",
                        attempt, retries, type(exc).__name__,
                    )
                    time.sleep(retry_delay * attempt)
        # SEC-1 (D12): путь источника в сообщении исключения маскируется.
        log.error(
            "Копия реестра не создана после %d попыток: %s",
            retries, applog.mask_path(str(last_exc)),
        )
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
        self._mark_party_dirty(rec)
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
        self._mark_party_dirty(rec)
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
        self._mark_document_dirty(rec)
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
            self._mark_document_deleted(doc_id)
            return self.save()
        return False

    def delete_counterparty(self, party_id) -> int:
        """Окончательно удаляет контрагента вместе со всеми его документами.

        Возвращает число удалённых документов; изменения сразу сохраняются.
        """
        party_exists = any(rec.get("id") == party_id for rec in self.counterparties)
        self.counterparties = [
            rec for rec in self.counterparties if rec.get("id") != party_id
        ]
        before = len(self.documents)
        removed_doc_ids = [rec.get("id") for rec in self.documents if rec.get("party_id") == party_id]
        self.documents = [
            rec for rec in self.documents if rec.get("party_id") != party_id
        ]
        removed = before - len(self.documents)
        if party_exists:
            self._mark_party_deleted(party_id)
        for rec_id in removed_doc_ids:
            self._mark_document_deleted(rec_id)
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
        self._mark_document_dirty(rec)
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
        self._mark_party_dirty(party)
        for rec in touched:
            self._mark_document_dirty(rec)
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
        self._mark_document_dirty(rec)
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
        self._mark_party_dirty(party)
        for rec in restored:
            self._mark_document_dirty(rec)
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
        self._mark_document_deleted(doc_id)
        return self.save()

    def purge_counterparty(self, party_id) -> int:
        """Удаляет карточку и все её документы окончательно."""
        before = len(self.documents)
        party_exists = any(rec.get("id") == party_id for rec in self.counterparties)
        self.counterparties = [
            rec for rec in self.counterparties if rec.get("id") != party_id
        ]
        removed_doc_ids = [
            rec.get("id") for rec in self.documents if rec.get("party_id") == party_id
        ]
        self.documents = [
            rec for rec in self.documents if rec.get("party_id") != party_id
        ]
        removed = before - len(self.documents)
        if party_exists:
            self._mark_party_deleted(party_id)
        for rec_id in removed_doc_ids:
            self._mark_document_deleted(rec_id)
        self.save()
        return removed

    def trash_count(self) -> int:
        return sum(
            1 for rec in self.counterparties + self.documents if is_deleted(rec)
        )

    # ------------------------------------------------------------------ #
    #  Счётчики статусной строки (STORE-1A, R-05a)
    # ------------------------------------------------------------------ #
    def status_counts(self) -> dict[str, int]:
        """Счётчики статусной строки тремя COUNT(*) по покрывающим индексам (R-05a).

        Раньше _update_status строила списки активных записей и трижды
        проходила их в памяти (П-аудит P-10) — доля UI-времени каждого
        refresh при 10k+. Теперь: активные контрагенты; активные документы
        и договоры ОДНИМ проходом; корзина. Запросы идут index-only сканом
        по idx_counterparties_deleted и idx_documents_status(deleted_at,
        kind). Если база недоступна (первый запуск до создания файла,
        блокировка) — честный откат к подсчёту по памяти: статусная строка
        не остаётся пустой.
        """
        path = db_path()
        if not os.path.exists(path):
            return self._status_counts_from_memory()
        conn = None
        try:
            conn = sqlite3.connect(path, timeout=BUSY_TIMEOUT_SECONDS)
            parties = conn.execute(
                "SELECT COUNT(*) FROM counterparties WHERE deleted_at IS NULL"
            ).fetchone()[0]
            documents, contracts = conn.execute(
                "SELECT COUNT(*), COUNT(CASE WHEN kind = ? THEN 1 END) "
                "FROM documents WHERE deleted_at IS NULL",
                (KIND_CONTRACT,),
            ).fetchone()
            trash_parties = conn.execute(
                "SELECT COUNT(*) FROM counterparties WHERE deleted_at IS NOT NULL"
            ).fetchone()[0]
            trash_documents = conn.execute(
                "SELECT COUNT(*) FROM documents WHERE deleted_at IS NOT NULL"
            ).fetchone()[0]
            return {
                "parties": int(parties),
                "documents": int(documents),
                "contracts": int(contracts),
                "trash": int(trash_parties + trash_documents),
            }
        except sqlite3.Error:
            return self._status_counts_from_memory()
        finally:
            if conn is not None:
                conn.close()

    def _status_counts_from_memory(self) -> dict[str, int]:
        """Резервный подсчёт по памяти — та же семантика, что у статусной строки."""
        docs = active_records(self.documents)
        return {
            "parties": sum(1 for rec in self.counterparties if not is_deleted(rec)),
            "documents": sum(1 for rec in docs if rec.get("kind") != KIND_UNKNOWN),
            "contracts": sum(1 for rec in docs if rec.get("kind") == KIND_CONTRACT),
            "trash": sum(
                1 for rec in self.counterparties + self.documents if is_deleted(rec)
            ),
        }

    def purge_all_trash(self) -> int:
        """Очищает корзину полностью. Возвращает число удалённых записей."""
        parties = [rec for rec in self.counterparties if is_deleted(rec)]
        docs = [rec for rec in self.documents if is_deleted(rec)]
        if not parties and not docs:
            return 0
        self.counterparties = [rec for rec in self.counterparties if not is_deleted(rec)]
        self.documents = [rec for rec in self.documents if not is_deleted(rec)]
        for rec in parties:
            self._mark_party_deleted(rec.get("id"))
        for rec in docs:
            self._mark_document_deleted(rec.get("id"))
        if not self.save():
            return 0
        applog.get_logger("storage").info(
            "Корзина очищена: контрагентов=%d, документов=%d", len(parties), len(docs),
        )
        return len(parties) + len(docs)

    def _trash_expired_records(self, retention_days: int) -> tuple[list[dict], list[dict]]:
        """Записи в корзине старше срока хранения: (контрагенты, документы)."""
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
        return parties, docs

    def count_expired_trash(self, retention_days: int = TRASH_RETENTION_DAYS) -> int:
        """Сколько записей удалит автоочистка корзины — сводка TRASH-WARN.

        Вызывается ДО cleanup_trash: «будет удалено N записей» показывается
        пользователю и пишется в журнал до фактического удаления.
        """
        parties, docs = self._trash_expired_records(retention_days)
        return len(parties) + len(docs)

    def cleanup_trash(self, retention_days: int = TRASH_RETENTION_DAYS) -> int:
        """Удаляет записи, пролежавшие в корзине дольше срока хранения."""
        parties, docs = self._trash_expired_records(retention_days)
        if not parties and not docs:
            return 0
        expired_party_ids = {rec.get("id") for rec in parties}
        expired_ids = {id(rec) for rec in parties + docs}
        # STORE-1A (R-01): id удаляемых записей помечаются ДО фильтрации —
        # каскад документы просроченных контрагентов уходит вместе с ними.
        for rec in parties:
            self._mark_party_deleted(rec.get("id"))
        for rec in docs:
            self._mark_document_deleted(rec.get("id"))
        for rec in self.documents:
            if (
                id(rec) not in expired_ids
                and rec.get("party_id") in expired_party_ids
            ):
                self._mark_document_deleted(rec.get("id"))
        self.counterparties = [
            rec for rec in self.counterparties if id(rec) not in expired_ids
        ]
        self.documents = [
            rec for rec in self.documents
            if id(rec) not in expired_ids
            and rec.get("party_id") not in expired_party_ids
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


# --------------------------------------------------------------------------- #
#  Автокопии реестра и снимки перед обновлением (BKUP-1)
# --------------------------------------------------------------------------- #
def _backup_files() -> list[str]:
    """Копии базы в backups/ (только файлы ``registry_*.db``, снимки не считаются)."""
    directory = backups_dir()
    try:
        names = os.listdir(directory)
    except OSError:
        return []
    return sorted(
        os.path.join(directory, name) for name in names
        if name.startswith(BACKUP_FILE_PREFIX) and name.endswith(".db")
    )


def _snapshot_files() -> list[str]:
    """Снимки перед обновлением в backups/snapshots/."""
    directory = snapshots_dir()
    try:
        names = os.listdir(directory)
    except OSError:
        return []
    return sorted(
        os.path.join(directory, name) for name in names
        if name.startswith(BACKUP_FILE_PREFIX) and name.endswith(".db")
    )


def _read_backup_marker() -> dict:
    """Отметка последней успешной копии; {} — отметки нет (первый запуск)."""
    try:
        with open(backup_marker_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _write_backup_marker(state: dict) -> bool:
    try:
        os.makedirs(backups_dir(), exist_ok=True)
        tmp = backup_marker_path() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(state, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, backup_marker_path())
        return True
    except (OSError, TypeError, ValueError):
        return False


def _strip_secret_keys(payload: dict) -> dict:
    """Копия настроек без ключей-секретов (ключ DaData — секрет пользователя)."""
    out = {}
    for key, value in payload.items():
        lowered = str(key).lower()
        if any(marker in lowered for marker in _SECRET_KEY_MARKERS):
            continue
        out[key] = value
    return out


def _copy_settings_snapshot(dest_path: str) -> bool:
    """settings.json в копию — без ключа DaData. False — файла нет или сбой."""
    try:
        with open(settings_json_path(), "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        if not isinstance(payload, dict):
            return False
        os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
        with open(dest_path, "w", encoding="utf-8") as fh:
            json.dump(_strip_secret_keys(payload), fh, ensure_ascii=False, indent=2)
        return True
    except (OSError, ValueError):
        return False


def _rotate(directory: str, keep: int) -> None:
    """Вытесняет самые старые копии (по времени создания) сверх ``keep``.

    Файл настроек-сателлит (``*_settings.json``) удаляется вместе со своей
    копией базы.
    """
    try:
        names = sorted(
            name for name in os.listdir(directory)
            if name.startswith(BACKUP_FILE_PREFIX) and name.endswith(".db")
        )
    except OSError:
        return
    stats = []
    for name in names:
        try:
            stats.append((os.path.getmtime(os.path.join(directory, name)), name))
        except OSError:
            continue
    stats.sort(reverse=True)  # свежие первыми
    for _mtime, name in stats[keep:]:
        try:
            os.remove(os.path.join(directory, name))
        except OSError:
            continue
        sidecar = os.path.join(directory, name[:-3] + BACKUP_SETTINGS_SUFFIX)
        if os.path.exists(sidecar):
            try:
                os.remove(sidecar)
            except OSError:
                pass


def create_startup_backup(store, version_label: str = "") -> dict:
    """Автокопия реестра при запуске (BKUP-1): до 7 копий, только при изменениях.

    Триггер: первый запуск (папки копий или отметки нет) либо реестр менялся
    с момента последней успешной копии (ревизия в meta отличается от отметки,
    или флаг dirty_since_backup). Содержимое копии — состояние базы на конец
    предыдущего прогона: копия создаётся до изменений текущего сеанса.
    Повторный запуск без изменений новых файлов не создаёт.

    Отказ копирования не блокирует старт: предупреждение в журнале и
    настройках, флаг изменений сохраняется — следующая попытка создаст копию.

    Возвращает ``{'status': 'created'|'skipped'|'failed', 'path', 'message'}``.
    """
    log = applog.get_logger("storage")
    # Карантин повреждённой базы / непрочитанный реестр: копировать пустое
    # состояние нельзя — автокопии вытеснили бы хорошие копии.
    if getattr(store, "load_warning", ""):
        return {
            "status": "skipped", "path": "",
            "message": "Реестр не прочитан — автокопия отложена до чистого старта.",
        }
    marker = _read_backup_marker()
    first_run = not marker or not _backup_files()
    changed = bool(getattr(store, "dirty_since_backup", False)) or (
        not first_run and marker.get("rev") != getattr(store, "registry_rev", 0)
    )
    if not first_run and not changed:
        return {
            "status": "skipped", "path": "",
            "message": "Изменений с последней копии нет — копия не создаётся.",
        }

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = os.path.join(backups_dir(), f"{BACKUP_FILE_PREFIX}{stamp}.db")
    if not store.backup_db(dest):
        return {
            "status": "failed", "path": dest,
            "message": (
                "Копия реестра не создана (подробности в журнале). "
                "Старт продолжен; попытка повторится при следующем запуске."
            ),
        }
    settings_copied = _copy_settings_snapshot(
        os.path.join(backups_dir(), f"{BACKUP_FILE_PREFIX}{stamp}{BACKUP_SETTINGS_SUFFIX}")
    )
    _rotate(backups_dir(), BACKUP_RETENTION)
    state = {
        "rev": getattr(store, "registry_rev", 0),
        "time": datetime.now().isoformat(timespec="seconds"),
        "version": str(version_label or ""),
        "file": os.path.basename(dest),
        "settings": bool(settings_copied),
    }
    if not _write_backup_marker(state):
        log.warning("Отметка о копии реестра не записана: %s", backup_marker_path())
    # Флаг снимается только успешной копией (карточка BKUP-1).
    store.dirty_since_backup = False
    log.info(
        "Автокопия реестра создана: %s (ревизия %d, settings=%s)",
        dest, state["rev"], settings_copied,
    )
    message = "Копия реестра создана."
    if not settings_copied:
        message = "Копия реестра создана; settings.json скопировать не удалось."
    return {"status": "created", "path": dest, "message": message}


def create_update_snapshot(store, version_label: str = "") -> str:
    """Снимок реестра перед обновлением (BKUP-1): отдельная папка, ротация до 3.

    Снимок создаётся до скачивания/запуска установщика и до любых изменений
    базы новой версией. Снимок ≠ автоматический откат версии приложения:
    восстановление базы в более старую/новую версию — ручная операция,
    совместимость данных проверяется владельцем. Снимки не вытесняют обычные
    копии и в лимит 7 не входят.
    """
    log = applog.get_logger("storage")
    if getattr(store, "load_warning", ""):
        log.warning("Снимок перед обновлением не создан: реестр не прочитан")
        return ""
    safe_label = re.sub(r"[^0-9A-Za-z.\-]", "", str(version_label or "")) or "unknown"
    dest = os.path.join(
        snapshots_dir(), f"{BACKUP_FILE_PREFIX}v{safe_label}_pre-update.db"
    )
    if not store.backup_db(dest):
        log.error("Снимок реестра перед обновлением не создан: %s", dest)
        return ""
    _rotate(snapshots_dir(), SNAPSHOT_RETENTION)
    log.info("Снимок реестра перед обновлением создан: %s", dest)
    return dest


def backup_status_text() -> str:
    """Состояние копий для блока настроек (без персональных данных)."""
    marker = _read_backup_marker()
    files = _backup_files()
    if not marker and not files:
        return ("Копий реестра пока нет — копия создаётся при запуске "
                "после изменений в реестре.")
    parts = []
    if marker.get("time"):
        parts.append(f"Последняя копия: {marker['time']}")
    if marker.get("version"):
        parts.append(f"версия {marker['version']}")
    parts.append(f"копий в папке: {len(files)}")
    if marker and not marker.get("settings", True):
        parts.append("settings.json в последней копии нет")
    return ", ".join(parts) + "."
