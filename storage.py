"""Хранилище данных программы: контрагенты и созданные документы.

Данные хранятся в одном JSON-файле в профиле пользователя
(%APPDATA%/DogovorGenerator/database.json), поэтому работают и в
разработке, и внутри собранного PyInstaller .exe.

Структура записей:
    контрагент: {"id", "type": "ooo"|"ip", "fields": {...}, "created_at"}
    документ:   {"id", "kind": "contract"|"additional"|"spec",
                 "party_id", "number", "date", "data": {...},
                 "extra": {...}, "text", "created_at"}

Надёжность записи (п. 2.1 плана):
    * запись идёт в уникальный ``database.json.<pid>.tmp`` с ``fsync``;
    * предыдущая версия сохраняется в ``database.json.bak`` (до 5 копий);
    * при временных ``OSError`` (антивирус, индексатор) — до 3 попыток;
    * в файле хранится ``schema_version``.

Корзина (п. 2.6): удаление помечает запись полем ``deleted_at``, а не стирает
её. Восстановление снимает метку, окончательная очистка удаляет физически.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime, timedelta

import applog

SETTINGS_DIR_NAME = "DogovorGenerator"

SCHEMA_VERSION = 2
BACKUP_COUNT = 5          # database.json.bak + .bak.1 … .bak.4
SAVE_RETRIES = 3
RETRY_DELAY_SECONDS = 0.25
TRASH_RETENTION_DAYS = 30
BAD_COPY_RETENTION_DAYS = 30   # сколько дней хранить database.json.bad-* (аудит 2.0.1)

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


def data_dir() -> str:
    """Каталог данных программы (создаётся при необходимости)."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, SETTINGS_DIR_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def db_path() -> str:
    return os.path.join(data_dir(), "database.json")


def _safe_int(value, default: int = 0) -> int:
    """Безопасно приводит значение из JSON к int (null/строка не роняют запуск)."""
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
    """Простое JSON-хранилище контрагентов и документов."""

    def __init__(self):
        self.counterparties: list[dict] = []
        self.documents: list[dict] = []
        self._seq = 1
        self.schema_version = SCHEMA_VERSION
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
                with open(path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
        except FileNotFoundError:
            log.info("Файл реестра не найден — создан пустой реестр")
            return  # первый запуск — пустой реестр
        except (OSError, ValueError) as exc:
            data = self._recover_from_broken_file(path, exc)
            if data is None:
                return

        if not isinstance(data, dict):
            self.load_warning = (
                "Файл реестра имеет неверную структуру и не был загружен.\n"
                f"Путь: {path}"
            )
            return

        raw_parties = data.get("counterparties", [])
        raw_documents = data.get("documents", [])
        if not isinstance(raw_parties, list) or not isinstance(raw_documents, list):
            # null/объект вместо списка: не роняем запуск, но честно предупреждаем.
            self.load_warning = (
                "Файл реестра повреждён: разделы «counterparties»/«documents» "
                "имеют неверный формат, поэтому реестр показан пустым.\n"
                f"Путь: {path}"
            )
            return

        self.schema_version = _safe_int(data.get("schema_version"), 1)
        self.counterparties = [rec for rec in raw_parties if isinstance(rec, dict)]
        self.documents = [rec for rec in raw_documents if isinstance(rec, dict)]

        dropped = (len(raw_parties) - len(self.counterparties)) + (
            len(raw_documents) - len(self.documents)
        )
        if dropped:
            log.warning("При чтении реестра отброшено битых записей: %d", dropped)

        self._seq = _safe_int(data.get("seq"), 1)
        max_id = [self._seq]
        max_id.extend(_safe_int(rec.get("id"), 0) for rec in self.counterparties)
        max_id.extend(_safe_int(rec.get("id"), 0) for rec in self.documents)
        self._seq = max(max_id) + 1

        if self.schema_version < SCHEMA_VERSION:
            log.info(
                "Реестр прочитан в схеме %d; при следующем сохранении будет схема %d",
                self.schema_version, SCHEMA_VERSION,
            )
        log.info(
            "Реестр загружен: контрагентов=%d, документов=%d, в корзине=%d",
            len(self.counterparties), len(self.documents),
            sum(1 for rec in self.counterparties + self.documents if is_deleted(rec)),
        )

    @staticmethod
    def _tmp_candidates(path: str) -> list[str]:
        """Временные файлы записи (старый ``.tmp`` и новые ``.<pid>.tmp``)."""
        directory = os.path.dirname(path)
        base = os.path.basename(path)
        candidates: list[str] = []
        try:
            names = os.listdir(directory)
        except OSError:
            return candidates
        for name in names:
            if name == base + ".tmp" or (
                name.startswith(base + ".") and name.endswith(".tmp")
            ):
                candidates.append(os.path.join(directory, name))

        def mtime(value: str) -> float:
            try:
                return os.path.getmtime(value)
            except OSError:
                return 0.0

        candidates.sort(key=mtime, reverse=True)
        return candidates

    def _recover_from_broken_file(self, path: str, exc: Exception) -> dict | None:
        """Спасает реестр: сначала пробует временные файлы, затем делает
        .bad-копию повреждённого файла. Возвращает прочитанные данные или None."""
        data: dict | None = None
        recovered_from = ""
        tmp_errors: list[str] = []
        for tmp_path in self._tmp_candidates(path):
            try:
                with open(tmp_path, "r", encoding="utf-8") as fh:
                    candidate = json.load(fh)
                if isinstance(candidate, dict):
                    data = candidate
                    recovered_from = os.path.basename(tmp_path)
                    break
            except (OSError, ValueError) as tmp_exc:
                tmp_errors.append(f"{os.path.basename(tmp_path)}: {tmp_exc}")

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        bad_path = f"{path}.bad-{stamp}"
        try:
            os.replace(path, bad_path)
            moved = f"\nПовреждённый файл сохранён как:\n{bad_path}"
        except OSError as move_exc:
            moved = f"\nНе удалось сохранить копию повреждённого файла: {move_exc}"

        tmp_note = ""
        if tmp_errors:
            tmp_note = "\nВременные файлы тоже не читаются: " + "; ".join(tmp_errors[:3])

        if data is not None:
            self.load_warning = (
                "Основной файл реестра повреждён, данные восстановлены из "
                f"{recovered_from}.\n"
                f"Ошибка: {exc}{tmp_note}{moved}"
            )
            applog.get_logger("storage").warning(
                "Реестр восстановлен из временного файла; повреждённый сохранён как .bad"
            )
            return data

        self.load_warning = (
            "Не удалось прочитать файл реестра — программа запущена с пустым "
            f"реестром.\nОшибка: {exc}{tmp_note}{moved}\n\n"
            "До первого сохранения скопируйте папку данных, если хотите "
            "попробовать восстановить записи вручную."
        )
        applog.get_logger("storage").error(
            "Реестр не прочитан; повреждённый файл сохранён как .bad"
        )
        return None

    def _purge_stale_bad_copies(self, path: str) -> None:
        """Удаляет ``.bad-*``-копии старше BAD_COPY_RETENTION_DAYS (аудит 2.0.1).

        Копии повреждённых файлов раньше копились без ограничения возраста
        и не попадали ни в какую очистку — теперь уборка идёт при каждом
        запуске, свежие копии не трогаются.
        """
        directory = os.path.dirname(path)
        try:
            names = os.listdir(directory)
        except OSError:
            return
        prefix = os.path.basename(path) + ".bad-"
        deadline = time.time() - BAD_COPY_RETENTION_DAYS * 86400
        removed = 0
        for name in names:
            if not name.startswith(prefix):
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

    @staticmethod
    def _remove_silent(path: str) -> None:
        try:
            if os.path.exists(path):
                os.remove(path)
        except OSError:
            pass

    def _rotate_backups(self, path: str) -> None:
        """Сдвигает ``.bak``-копии и сохраняет текущую версию как ``.bak``."""
        log = applog.get_logger("storage")
        try:
            oldest = f"{path}.bak.{BACKUP_COUNT - 1}"
            if os.path.exists(oldest):
                os.remove(oldest)
            for index in range(BACKUP_COUNT - 2, 0, -1):
                src = f"{path}.bak.{index}"
                if os.path.exists(src):
                    os.replace(src, f"{path}.bak.{index + 1}")
            if os.path.exists(f"{path}.bak"):
                os.replace(f"{path}.bak", f"{path}.bak.1")
            if os.path.exists(path):
                shutil.copy2(path, f"{path}.bak")
                try:
                    # Бэкап — «вторая линия обороны»: он обязан пережить сбой
                    # питания, поэтому сбрасывается на диск явно (аудит 2.0.1).
                    with open(f"{path}.bak", "rb") as fh:
                        os.fsync(fh.fileno())
                except OSError:
                    pass  # fsync недоступен — бэкап всё равно лучше, чем ничего
        except OSError as exc:
            # Бэкап — страховка, но не причина срывать сохранение.
            log.warning("Не удалось обновить резервные копии реестра: %s", exc)

    def _cleanup_stale_tmp(self, path: str) -> None:
        """Удаляет временные файлы предыдущих неудачных записей."""
        own = f"{path}.{os.getpid()}.tmp"
        for candidate in self._tmp_candidates(path):
            if os.path.normcase(candidate) != os.path.normcase(own):
                self._remove_silent(candidate)

    def save(self) -> bool:
        """Атомарно записывает реестр. Возвращает False и заполняет last_error при сбое.

        Долгая запись (более 500 мс — обычно антивирус или медленный диск)
        попадает в журнал (п. 10 этапа 5).
        """
        log = applog.get_logger("storage")
        self.last_error = ""
        payload = {
            "schema_version": SCHEMA_VERSION,
            "seq": self._seq,
            "counterparties": self.counterparties,
            "documents": self.documents,
        }
        try:
            text = json.dumps(payload, ensure_ascii=False, indent=2)
        except (TypeError, ValueError) as exc:
            self.last_error = f"Данные реестра не удалось сериализовать: {exc}"
            log.error("Сериализация реестра не удалась: %s", type(exc).__name__)
            return False

        path = db_path()
        directory = os.path.dirname(path)
        tmp = f"{path}.{os.getpid()}.tmp"
        last_exc: Exception | None = None
        with applog.timed("Запись реестра", 500, "storage"):
            for attempt in range(1, SAVE_RETRIES + 1):
                try:
                    os.makedirs(directory, exist_ok=True)
                    with open(tmp, "w", encoding="utf-8") as fh:
                        fh.write(text)
                        fh.flush()
                        os.fsync(fh.fileno())
                    self._rotate_backups(path)
                    os.replace(tmp, path)
                    self._cleanup_stale_tmp(path)
                    self.schema_version = SCHEMA_VERSION
                    log.info(
                        "Реестр записан: контрагентов=%d, документов=%d",
                        len(self.counterparties), len(self.documents),
                    )
                    return True
                except OSError as exc:
                    last_exc = exc
                    self._remove_silent(tmp)
                    if attempt < SAVE_RETRIES:
                        log.warning(
                            "Запись реестра не удалась (попытка %d из %d): %s",
                            attempt, SAVE_RETRIES, type(exc).__name__,
                        )
                        time.sleep(RETRY_DELAY_SECONDS * attempt)

        self.last_error = f"Не удалось записать файл реестра: {last_exc}"
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
            "Полное название компании (Покупь)",
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
        fields.get("Полное название компании (Покупь)", "").strip() or "Без названия"


def display_with_inn(rec: dict) -> str:
    """Имя для выпадающих списков: название + ИНН."""
    inn = str(rec.get("fields", {}).get("ИНН", "")).strip()
    name = display_name(rec)
    return f"{name} · ИНН {inn}" if inn else name
