"""Безопасность и пути (п. 7.7): safe_filename, allowlist импорта, MOTW.

``..`` и запрещённые символы, зарезервированные имена Windows (CON, COM1),
длина, кириллица, регистр; allowlist расширений импорта; отсутствие
``os.startfile`` в импортере; копирование с Mark-of-the-Web.
"""
from __future__ import annotations

import inspect
import os

import pytest

import document_importer as di
import services


# ---------- safe_filename ----------

def test_traversal_is_neutralized():
    name = services.safe_filename("../../etc/passwd")
    assert ".." not in name and "/" not in name and "\\" not in name


def test_dotdot_alone_falls_back():
    assert services.safe_filename("..", "Документ") == "Документ"
    assert services.safe_filename(".", "Документ") == "Документ"


@pytest.mark.parametrize("reserved", ["CON", "PRN", "NUL", "COM1", "LPT9"])
def test_reserved_windows_names_prefixed(reserved):
    assert services.safe_filename(reserved, "").startswith("_")


def test_reserved_name_case_insensitive():
    assert services.safe_filename("con", "").startswith("_")
    assert services.safe_filename("Com1.txt", "").startswith("_")


def test_forbidden_chars_removed():
    name = services.safe_filename('файл<>:"/\\|?*версия')
    for char in '<>:"/\\|?*':
        assert char not in name
    assert name.startswith("файл")


def test_cyrillic_preserved():
    assert services.safe_filename("ООО «Тест»") == "ООО «Тест»"


def test_long_name_truncated_with_hash_suffix():
    name = services.safe_filename("Д" * 200)
    assert len(name) <= services.MAX_FILENAME_LENGTH
    assert "~" in name


def test_control_characters_stripped():
    assert services.safe_filename("a\x01b\x1fc") == "abc"


def test_empty_name_falls_back():
    assert services.safe_filename("", "Документ") == "Документ"
    assert services.safe_filename("???", "Документ") == "Документ"


# ---------- Ограничение длины пути ----------

def test_fit_document_path_short_is_unchanged(tmp_path):
    path = services.fit_document_path(str(tmp_path), "Договор №516", ".docx")
    assert path == os.path.join(str(tmp_path), "Договор №516.docx")


def test_fit_document_path_long_dir_raises(tmp_path):
    long_dir = str(tmp_path)
    while len(long_dir) <= services.MAX_PATH_LENGTH:
        long_dir = os.path.join(long_dir, "очень-длинное-имя-вложенной-папки")
    with pytest.raises(OSError):
        services.fit_document_path(long_dir, "Договор", ".docx")


# ---------- Allowlist импорта ----------

def test_import_allowlist_is_docx_xlsx_only():
    import main  # константа allowlist живёт рядом с диалогом импорта

    assert main.IMPORT_ALLOWED_EXTENSIONS == {".docx", ".xlsx"}


def test_doc_extension_rejected_by_importer(tmp_path):
    path = tmp_path / "old.doc"
    path.write_bytes(b"stub")
    result = di.analyze_file(str(path))
    assert result["ok"] is False
    assert ".doc" in result["error"]


def test_importer_source_has_no_startfile():
    source = inspect.getsource(di)
    assert "startfile" not in source, (
        "импортер не должен сам открывать файлы — только копировать их "
        "в папку документов (п. 1.2)"
    )


def test_unknown_kind_is_never_ok(tmp_path):
    path = tmp_path / "random.docx"
    import zipfile
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml",
                         "<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>"
                         "<w:r><w:t>случайный текст без реквизитов</w:t></w:r></w:p>")
    result = di.analyze_file(str(path))
    assert result["ok"] is False
    assert result["kind"] == "unknown" or result["error"]


# ---------- Mark-of-the-Web ----------

def test_copy_with_zone_identifier_copies_content(tmp_path):
    src = tmp_path / "source.docx"
    dst = tmp_path / "dest.docx"
    src.write_bytes(b"PK\x03\x04 document bytes")

    services.copy_with_zone_identifier(str(src), str(dst))

    assert dst.read_bytes() == b"PK\x03\x04 document bytes"


def test_copy_with_zone_identifier_missing_source_stream(tmp_path):
    # На не-Windows потока Zone.Identifier нет — копирование всё равно работает.
    src = tmp_path / "a.docx"
    dst = tmp_path / "b.docx"
    src.write_bytes(b"data")
    services.copy_with_zone_identifier(str(src), str(dst))
    assert dst.read_bytes() == b"data"


# ---------- SEC-1 (D12): маскирование пути профиля в журнале ----------

def test_mask_path_hides_profile_root(appdata, monkeypatch):
    import applog

    db = os.path.join(str(appdata), "DogovorGenerator", "database.db")
    message = f"unable to open database file: {db}"
    masked = applog.mask_path(message)
    assert str(appdata) not in masked, "путь профиля не раскрывается"
    assert "<профиль>" in masked
    assert "DogovorGenerator\\database.db" in masked or \
        "DogovorGenerator/database.db" in masked, \
        "хвост внутри профиля виден для диагностики"


def test_mask_path_keeps_foreign_paths_intact(appdata):
    import applog

    foreign = "C:\\Program Files\\EW.DocHub\\EW.DocHub.exe"
    assert applog.mask_path(f"ошибка: {foreign}") == f"ошибка: {foreign}", \
        "пути вне профиля не искажаются"


def test_mask_path_prefers_longest_root(appdata, monkeypatch):
    """APPDATA лежит внутри USERPROFILE: замена начинается с длинного корня,
    иначе внутри остаётся хвост с AppData/Roaming."""
    import applog

    monkeypatch.setattr(applog, "_profile_roots", lambda: [str(appdata)])
    deep = os.path.join(str(appdata), "logs", "app.log")
    assert applog.mask_path(deep) == "<профиль>\\logs\\app.log" or \
        applog.mask_path(deep).endswith("<профиль>" + os.sep + "logs" + os.sep + "app.log")


def test_mask_path_empty_is_empty():
    import applog

    assert applog.mask_path("") == ""


def test_storage_save_failure_logs_masked_path(appdata, caplog, monkeypatch):
    """Запись реестра в занятый файл: путь базы в строке лога маскирован."""
    import logging

    import storage

    store = storage.DataStore()
    # STORE-1A: изменение с несохранённой записью (save=False) — иначе
    # инкрементальный save() без dirty-изменений не писал бы в базу вовсе.
    store.add_counterparty("ooo", {"ИНН": "2540123456"}, save=False)
    # Уводим путь базы на неписуемое место через monkeypatch db_path.
    locked = os.path.join(str(appdata), "locked-dir", "database.db")
    os.makedirs(os.path.dirname(locked), exist_ok=True)
    monkeypatch.setattr(storage, "db_path", lambda: locked)
    # Каталог оставляем, но заменяем файл на каталог: sqlite вернёт ошибку с путём.
    os.makedirs(locked, exist_ok=True)

    logger = logging.getLogger(storage.applog.LOGGER_NAME + ".storage")
    records: list[logging.LogRecord] = []

    class _Capture(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            records.append(record)

    handler = _Capture(level=logging.ERROR)
    logger.addHandler(handler)
    try:
        store.save()
    finally:
        logger.removeHandler(handler)

    errors = [r.getMessage() for r in records if r.levelno == logging.ERROR]
    assert errors, "сбой записи дошёл до журнала"
    assert all(str(appdata) not in message for message in errors), \
        "полный путь профиля в строке лога отсутствует"
