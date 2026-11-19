"""Быстрая проверка интерфейса и логики без ручного запуска."""
from __future__ import annotations

import ctypes
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile

import applog
import dadata_client as dadata
import document_builder
import single_instance
import win_shell
import sqlite3
from datetime import datetime

# Настройки приложения и база данных (контрагенты/документы) читаются из APPDATA:
# изолируем тест от реальных данных пользователя, чтобы результат был детерминированным.
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "dogovor_smoke_appdata")

# Тест не ходит в сеть: авто-проверка обновлений (2.15.0) глушится целиком;
# сетевые функции updater в этом тесте не подменяются и не вызываются.
os.environ["EWDOCHUB_NO_UPDATE_CHECK"] = "1"

# Windows-консоль может быть в cp1252/cp1251 (без полной поддержки
# кириллицы или вообще без неё): форсируем UTF-8 для вывода сводки,
# иначе первый же print с русским текстом падает с UnicodeEncodeError
# (2.15.0, Windows-раннер CI).
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import main as m
import reconcile
import settings_controller
import storage
import updater
import document_importer as di
from docx import Document as _Doc

# Чистая база на каждый запуск: файл мог остаться от прошлого прогона теста.
try:
    os.remove(storage.db_path())
except OSError:
    pass

# Модальные алерты и подтверждения в тесте никто не закроет — подменяем заглушками.
ALERTS: list[tuple] = []
m.ui.alert = lambda *args, **kwargs: ALERTS.append(args)
m.ui.confirm = lambda *args, **kwargs: True
# В тесте не открываем Word и не запускаем файлы.
OPENED: list[str] = []
_real_open_file = m.open_file  # реальная функция до стаба ниже (нужна п. 6.2)
m.open_file = lambda path: OPENED.append(path)


# Аудит 2.0.1: check() больше не завершает процесс на первой же неудаче —
# неудачи копятся и печатаются сводкой в конце (в CI видно всё сразу,
# а не одну ошибку за прогон). Код возврата — числу неудач (до 100).
FAILURES: list[str] = []
_TOTAL_CHECKS = 0


def check(condition, message):
    global _TOTAL_CHECKS
    _TOTAL_CHECKS += 1
    if condition:
        print(f"ok: {message}")
    else:
        FAILURES.append(message)
        print(f"FAIL: {message}")


# ---------- Чистая логика ----------
check(m.format_phone("9123456789") == "+7 (912) 345-67-89", "маска телефона")
check(m.format_phone("89121234567") == "+7 (912) 123-45-67", "маска телефона: 8 -> +7")
check(m.format_phone("") == "", "маска телефона: пусто")
check(m.number_to_words_ru(1500000) == "один миллион пятьсот тысяч", "число прописью 1 500 000")
check(m.number_to_words_ru(21) == "двадцать один", "число прописью 21")
check(m.number_to_words_ru(3000) == "три тысячи", "число прописью 3000")
check(m.number_to_words_ru(700) == "семьсот", "число прописью: сотни")
check(m.amount_to_words_ru(150000) == "сто пятьдесят тысяч рублей", "сумма прописью")
check(m.format_number_grouped("1500000") == "1 500 000", "разделители разрядов")
check(m.validate_day_input("31") and not m.validate_day_input("32"), "валидатор дня 1-31")
check(m.validate_term_input("45") and not m.validate_term_input("abc"), "валидатор отсрочки")
check(m.validate_limit_input("1 500") and not m.validate_limit_input("15a"), "валидатор лимита")
check(m.get_spec_settings()["spec_year"] == datetime.now().year,
      "настройки: год спецификации берётся на момент вызова, а не импорта модуля")

errors = m.validate_party_fields({}, is_ip=False)
check("ИНН" in errors and "Расчётный счёт" in errors, "валидация контрагента: пустые поля")
check(m.validate_party_fields({
    "ИНН": "7701234567", "КПП": "770101001", "ОГРН": "1037739010221",
    "Юридический адрес": "Москва", "Email": "a@b.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012", "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "Банк", "БИК": "044525225",
    document_builder.F.FULL_NAME: "ООО", "Сокращенное название компании": "ООО",
    "ФИО директора": "Иванов И.И.",
}, is_ip=False) == {}, "валидация контрагента: валидные данные")

# ---------- 5.6: один проход замены плейсхолдеров ----------
def _para(*runs_text):
    from docx import Document as _Doc
    document = _Doc()
    paragraph = document.add_paragraph()
    for chunk in runs_text:
        paragraph.add_run(chunk)
    return paragraph

_p = _para("Привет, {{NAME}}!", " И {{NAME}} снова.")
m.replace_in_paragraph(_p, {"{{NAME}}": "ООО Ромашка"})
check(_p.runs[0].text == "Привет, ООО Ромашка!" and _p.runs[1].text == " И ООО Ромашка снова.",
      "5.6: плейсхолдер в двух runs заменён, остальные тексты не тронуты")

_p = _para("Поставщик: {{", "PARTY", "}} и {{PARTY}} в конце.")
m.replace_in_paragraph(_p, {"{{PARTY}}": "ИП Иванов"})
check(_p.text == "Поставщик: ИП Иванов и ИП Иванов в конце.",
      "5.6: плейсхолдер, разорванный по runs, собирается и заменяется")

_p = _para("{{A}}{{A}}{{A}}")
m.replace_in_paragraph(_p, {"{{A}}": "X"})
check(_p.text == "XXX", "5.6: три одинаковых плейсхолдера подряд")

_p = _para("до {{A}} после")
m.replace_in_paragraph(_p, {"{{A}}": "Y{{A}}Z"})
check(_p.text == "до Y{{A}}Z после", "5.6: подстановка, содержащая сам плейсхолдер, не зацикливает замену")

_p = _para("А{{X}}Б{{Y}}В")
m.replace_in_paragraph(_p, {"{{X}}": "1", "{{Y}}": "2"})
check(_p.text == "А1Б2В", "5.6: два разных плейсхолдера в одной строке")

# ---------- DaData: ключ только через настройки (п. 1.3) ----------
# Ключ мог остаться в песочнице от прошлого прогона — чистим для детерминизма.
for _stale_key in (dadata.key_storage_path(), dadata.legacy_key_storage_path()):
    try:
        os.remove(_stale_key)
    except OSError:
        pass
check(dadata.get_api_key() == "", "dadata: без сохранённого ключа get_api_key пуст")
check(not dadata.save_api_key(""), "dadata: пустой ключ не сохраняется")
check(not dadata.save_api_key("ключ с пробелом"), "dadata: ключ с пробелом отклонён")
check(dadata.save_api_key("smoke-test-key"), "dadata: ключ сохранён")
check(dadata.get_api_key() == "smoke-test-key", "dadata: сохранённый ключ читается")
# Аудит 2.0.1: на Windows ключ на диске зашифрован (DPAPI); в dev/CI — открытый текст.
with open(dadata.key_storage_path(), encoding="utf-8") as _kf:
    _stored_key = _kf.read()
if dadata._dpapi_available():
    check(_stored_key.startswith(dadata.KEY_ENC_PREFIX),
          "dadata: на Windows ключ хранится зашифрованным (DPAPI)")
else:
    check(not _stored_key.startswith(dadata.KEY_ENC_PREFIX),
          "dadata: без DPAPI (dev/CI) ключ хранится открытым текстом")

# Миграция ключа из прежнего расположения (DogovorPostavki).
_appdata_backup = os.environ["APPDATA"]
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_dadata_legacy")
os.makedirs(os.path.dirname(dadata.legacy_key_storage_path()), exist_ok=True)
with open(dadata.legacy_key_storage_path(), "w", encoding="utf-8") as _kf:
    _kf.write("legacy-key-123")
check(dadata.get_api_key() == "legacy-key-123",
      "dadata: ключ из прежнего расположения подхвачен")
check(dadata.clear_api_key(), "dadata: ключ удаляется")
check(dadata.get_api_key() == "", "dadata: после удаления ключа get_api_key пуст")
os.environ["APPDATA"] = _appdata_backup

# В исходнике нет вшитых ключей: присваивания строковых литералов запрещены.
with open("dadata_client.py", "r", encoding="utf-8") as _df:
    _dadata_src = _df.read()
check(re.search(r"(API_KEY|SECRET_KEY)\s*=\s*[\"'][^\"']+[\"']", _dadata_src) is None,
      "dadata: в исходнике нет вшитых ключей")

# ---------- Санитизация имён и путей (п. 2.7) ----------
check(m.safe_filename("..") == "", "пути: «..» не превращается в имя файла")
check(m.safe_filename("CON").startswith("_"), "пути: зарезервированное имя CON обезврежено")
check(len(m.safe_filename("Очень длинное имя " * 20)) <= m.MAX_FILENAME_LENGTH,
      "пути: имя файла ограничено по длине")
check(":" not in m.safe_filename('a:b/c\\d'), "пути: запрещённые символы удалены")
# Проверка вложенности путей зависит от семантики ОС: на Windows «C:\docs»,
# на Linux/macOS «/docs». Логика одна, разделители и корни разные — поэтому
# smoke-тест можно гонять и вне Windows (CI, Linux-машина).
if sys.platform == "win32":
    _path_base, _path_inside, _path_outside = r"C:\docs", r"C:\docs\party", r"C:\other"
else:
    _path_base, _path_inside, _path_outside = "/docs", "/docs/party", "/other"
check(settings_controller.SettingsController.is_inside_directory(_path_base, _path_inside),
      "пути: вложенная папка разрешена")
check(not settings_controller.SettingsController.is_inside_directory(_path_base, _path_outside),
      "пути: выход за папку документов запрещён")

# ---------- Хранилище ----------
# v2.13.0: опечатка «(Покупь)» в имени поля исправлена (п. 9.6).
check(document_builder.F.FULL_NAME == "Полное название компании (Покупатель)",
      "имя поля: опечатка «Покупь» исправлена на «Покупатель»")
_legacy = {"fields": {storage.LEGACY_FULL_NAME_FIELD: "ООО Ромашка"}, "data": {}}
storage._normalize_field_names(_legacy)
check(_legacy["fields"].get(document_builder.F.FULL_NAME) == "ООО Ромашка"
      and storage.LEGACY_FULL_NAME_FIELD not in _legacy["fields"],
      "миграция поля: старый ключ «(Покупь)» переименован при загрузке")

# ---------- DPI и хуки окна (п. 6.3) ----------
check(win_shell.dpi_scale_from_wparam((144 << 16) | 144) == 1.5
      and win_shell.dpi_scale_from_wparam((120 << 16) | 120) == 1.25,
      "DPI: масштаб интерфейса из wparam WM_DPICHANGED (144 → 1.5, 120 → 1.25)")
_rc = win_shell._RECT(100, 50, 1700, 950)
check(win_shell.suggested_rect_from_lparam(ctypes.addressof(_rc)) == (100, 50, 1600, 900),
      "DPI: прямоугольник-подсказка Windows из lparam WM_DPICHANGED")
check(bool(win_shell.documents_known_folder()),
      "DPI: папка «Документы» определяется (п. 6.2)")
win_shell.enable_dpi_awareness()
check(True, "DPI: per-monitor awareness установлен без ошибок")

os.makedirs(storage.data_dir(), exist_ok=True)
store = storage.DataStore()
party_rec = store.add_counterparty("ooo", {
    "ИНН": "7701234567", "КПП": "770101001", "ОГРН": "1037739010221",
    "Полное название компании (Покупь)": "Общество с ограниченной ответственностью Ромашка",
    "Сокращенное название компании": "ООО Ромашка",
    "ФИО директора": "Иванов Иван Иванович", "Основание действий": "Устав",
    "Юридический адрес": "г. Москва, ул. Ленина, д. 1",
    "Фактический адрес": "г. Москва, ул. Ленина, д. 1",
    "Email": "info@romashka.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012", "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "ПАО Сбербанк", "БИК": "044525225",
})
check(store.get_counterparty(party_rec["id"]) is not None, "хранилище: контрагент сохранён")
check(store.find_by_inn("7701234567") is not None, "хранилище: поиск по ИНН")

# ---------- Надёжность записи реестра (п. 2.1, п. 11 этапа 5) ----------
check(store.save(), "хранилище: повторная запись реестра")
check(os.path.exists(storage.db_path()), "хранилище: база SQLite создана")
with sqlite3.connect(storage.db_path()) as _conn:
    _meta = dict(_conn.execute("SELECT key, value FROM meta"))
check(int(_meta.get("schema_version", "0")) == storage.SQLITE_SCHEMA_VERSION,
      "хранилище: schema_version записан в базу")
check(storage.is_deleted({"deleted_at": "2026-01-01T00:00:00"}),
      "корзина: запись с deleted_at считается удалённой")
check(not storage.is_deleted({}), "корзина: запись без deleted_at активна")

# ---------- 2.1: сбой записи не теряет предыдущую версию реестра ----------
_appdata_p1 = os.environ["APPDATA"]
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_p1_storage")
shutil.rmtree(storage.data_dir(), ignore_errors=True)
os.makedirs(storage.data_dir(), exist_ok=True)
_p1 = storage.DataStore()
_p1.add_counterparty("ooo", {
    "ИНН": "7709991111", "Сокращенное название компании": "ООО Первая",
}, save=False)
check(_p1.save(), "2.1: первая запись реестра успешна")

_orig_connect = storage._connect_db

def _locked_connect(path, *args, **kwargs):
    raise sqlite3.OperationalError("database is locked (имитация антивируса)")

storage._connect_db = _locked_connect
_orig_sleep = time.sleep
time.sleep = lambda _seconds: None  # ускоряем повторы записи
try:
    check(_p1.save() is False, "2.1: сбой записи повторён 3 раза и возвращён False")
finally:
    storage._connect_db = _orig_connect
    time.sleep = _orig_sleep
check(bool(_p1.last_error), "2.1: причина сбоя доступна пользователю")
_reloaded = storage.DataStore()
check(len(_reloaded.counterparties) == 1,
      "2.1: при сбое записи предыдущая версия реестра не потеряна")

# Временный сбой (файл занят сканером): запись повторами всё же проходит.
_transients = {"left": 2}

def _flaky_connect(path, *args, **kwargs):
    if _transients["left"] > 0:
        _transients["left"] -= 1
        raise sqlite3.OperationalError("database is locked (имитация сканера)")
    return _orig_connect(path, *args, **kwargs)

time.sleep = lambda _seconds: None
storage._connect_db = _flaky_connect
try:
    check(_p1.save(), "2.1: повтор записи переживает 2 временных сбоя")
finally:
    storage._connect_db = _orig_connect
    time.sleep = _orig_sleep
check(_transients["left"] == 0, "2.1: обе запасные попытки понадобились")
check(not [n for n in os.listdir(storage.data_dir()) if n.endswith(".tmp")],
      "2.1: временные файлы после записи не остаются")
os.environ["APPDATA"] = _appdata_p1

# ---------- Уборка старых .bad-* копий (аудит 2.0.1) ----------
_bad_old = storage.db_path() + ".bad-20200101-000000"
_bad_fresh = storage.db_path() + ".bad-29991231-235959"
for _bad_path in (_bad_old, _bad_fresh):
    with open(_bad_path, "w", encoding="utf-8") as _fh:
        _fh.write("{}")
_bad_past = time.time() - (storage.BAD_COPY_RETENTION_DAYS + 10) * 86400
os.utime(_bad_old, (_bad_past, _bad_past))
storage.DataStore()  # загрузка запускает уборку
check(not os.path.exists(_bad_old) and os.path.exists(_bad_fresh),
      "хранилище: .bad-копии старше лимита убираются, свежие остаются")
try:
    os.remove(_bad_fresh)
except OSError:
    pass

# ---------- Блокировка второго экземпляра (п. 2.2) ----------
_lock = single_instance.InstanceLock()
check(_lock.acquire(), "2.2: первый запуск захватывает блокировку")
_lock.release()
check(single_instance.InstanceLock().acquire(),
      "2.2: после release блокировка снова свободна")
check(single_instance.mutex_name().startswith("Global\\")
      and "EWDocHub" in single_instance.mutex_name(),
      "2.2: мьютекс в Global-пространстве (защита и от второй RDP-сессии)")
check(bool(single_instance.MUTEX_NAME) and "EWDocHub" in single_instance.MUTEX_NAME,
      "2.2: именованный mutex настроен")
with open("main.py", encoding="utf-8") as _mf:
    _main_src = _mf.read()
# Исходники контроллеров декомпозиции ContractApp (аудит 2.0.1, P2) —
# часть «приложения» для строковых проверок ниже.
with open("import_controller.py", encoding="utf-8") as _icf:
    _app_src = _main_src + _icf.read()
check("if not lock.acquire()" in _main_src
      and "single_instance.show_already_running" in _main_src,
      "2.2: старт приложения gated захватом блокировки")

# ---------- Логирование без реквизитов (п. 2.3) ----------
_ewdochub_logger = logging.getLogger(applog.LOGGER_NAME)
_saved_handlers = _ewdochub_logger.handlers[:]
applog._configured = False
for _handler in _saved_handlers:
    _ewdochub_logger.removeHandler(_handler)
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_p1_log")
shutil.rmtree(applog.log_dir(), ignore_errors=True)
applog.setup_logging("1.4.0-smoke")
try:
    _log_store = storage.DataStore()
    _log_party = _log_store.add_counterparty("ooo", {
        "ИНН": "7701234567", "Расчётный счёт": "40702810123456789012",
        "ФИО директора": "Иванов Иван Иванович",
        "Сокращенное название компании": "ООО Ромашка",
    }, save=False)
    _log_doc = _log_store.add_document(
        storage.KIND_CONTRACT, _log_party["id"], "516", "12.01.2026",
        {"ИНН": "7701234567"},
        text="Секретный текст договора поставки № 516", save=False,
    )
    check(_log_store.save(), "2.3: операции хранилища работают с включённым логом")
    check(_log_store.trash_document(_log_doc["id"]), "2.3: событие корзины отработало")
    with applog.timed("Запись документа", 0, "documents"):
        pass  # порог 0 мс — «медленная» операция точно попадёт в лог
    try:
        with applog.timed("Опасная операция", 0, "documents"):
            raise RuntimeError("имитация сбоя")
    except RuntimeError:
        pass
    for _handler in _ewdochub_logger.handlers:
        _handler.flush()
    _log_text = open(applog.log_path(), encoding="utf-8").read()
    check("Запуск EW.DocHub" in _log_text and "1.4.0-smoke" in _log_text,
          "2.3: в логе есть старт и версия приложения")
    check("Реестр записан" in _log_text and "перемещён в корзину" in _log_text,
          "2.3: операции хранилища диагностируются по логу")
    check("заняла" in _log_text, "2.3: длительные операции фиксируются (timed)")
    check("прервана ошибкой" in _log_text,
          "2.3: прерванная операция попадает в лог")
    check(applog.MAX_BYTES >= 1024 * 1024 and applog.BACKUP_COUNT >= 3,
          "2.3: ротация лога настроена (1+ МБ, 3+ копии)")
    # Приёмка 2.3: в логе нет реквизитов, имён и текстов документов.
    for _secret in ("7701234567", "40702810123456789012", "Иванов Иван Иванович",
                    "ООО Ромашка", "Секретный текст договора поставки"):
        check(_secret not in _log_text,
              f"2.3: в логе нет реквизитов ({_secret[:14]}…)")
    check(applog.redact("7701234567") == "77********",
          "2.3: redact() маскирует значение для логов")
finally:
    for _handler in _ewdochub_logger.handlers[:]:
        try:
            _handler.close()
        except Exception:
            pass
        _ewdochub_logger.removeHandler(_handler)
    for _handler in _saved_handlers:
        _ewdochub_logger.addHandler(_handler)
    applog._configured = True
    os.environ["APPDATA"] = _appdata_p1

# ---------- Корзина: документы, файлы, просрочка (п. 2.5, 2.6) ----------
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_p1_trash")
shutil.rmtree(storage.data_dir(), ignore_errors=True)
os.makedirs(storage.data_dir(), exist_ok=True)
_tstore = storage.DataStore()
_tparty = _tstore.add_counterparty("ooo", {
    "ИНН": "7719998888", "Сокращенное название компании": "ООО Корзина",
}, save=False)
_disk_doc = os.path.join(tempfile.gettempdir(), "smoke_p1_disk.docx")
with open(_disk_doc, "wb") as _fh:
    _fh.write(b"PK\x03\x04 stub")
_tdoc = _tstore.add_document(
    storage.KIND_CONTRACT, _tparty["id"], "7", "01.03.2026", {},
    extra={"source_file": _disk_doc}, save=False,
)
check(_tstore.save(), "корзина: подготовка реестра выполнена")
check(_tstore.trash_document(_tdoc["id"]), "2.6: документ перемещён в корзину")
check(_tstore.get_document(_tdoc["id"]) is None and _tstore.trashed_documents(),
      "2.6: удалённый документ неактивен, но доступен в корзине")
check(os.path.exists(_disk_doc),
      "2.5: перемещение в корзину не трогает файл на диске")
check(_tstore.trash_document(_tdoc["id"]) is False,
      "2.6: повторное перемещение в корзину отклонено")
check(_tstore.restore_document(_tdoc["id"]), "2.6: документ восстановлен из корзины")
check(_tstore.get_document(_tdoc["id"]) is not None,
      "2.6: восстановленный документ снова активен")
check(_tstore.restore_document(_tdoc["id"]) is False,
      "2.6: восстановление активного документа отклонено")
check(_tstore.trash_document(_tdoc["id"]), "2.6: документ снова в корзине")
_tdoc["deleted_at"] = "2020-01-01T00:00:00"
check(_tstore.cleanup_trash(retention_days=1) == 1,
      "2.6: просроченная запись удаляется автоматически")
check(_tstore.get_document(_tdoc["id"], include_deleted=True) is None,
      "2.6: после автоочистки записи нет даже в корзине")
check(os.path.exists(_disk_doc),
      "2.6: автоочистка корзины не трогает файлы на диске")
os.remove(_disk_doc)
os.environ["APPDATA"] = _appdata_p1

# ---------- Защита чтения реестра (п. 1.4, п. 11 этапа 5) ----------
_appdata_main = os.environ["APPDATA"]
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_storage_bad")
# Песочница могла остаться от прошлого прогона (в т.ч. .bad файлы) — сбрасываем.
shutil.rmtree(storage.data_dir(), ignore_errors=True)
os.makedirs(storage.data_dir(), exist_ok=True)

# Повреждённая база: убирается в .bad-копию, пользователь предупреждён.
with open(storage.db_path(), "wb") as _fh:
    _fh.write(b"{\xe4\xee\xea\xf3\xec\xe5\xed\xf2\xfb \xe1\xe5\xe7 \xe1\xe0\xe7\xfb")
_broken = storage.DataStore()
check(bool(_broken.load_warning), "реестр: при битой базе пользователь предупреждён")
check(not _broken.counterparties and not _broken.documents,
      "реестр: битая база — реестр пуст, запуск не падает")
_bad_copies = [n for n in os.listdir(storage.data_dir())
               if n.startswith("database.db.bad-")]
check(len(_bad_copies) == 1,
      f"реестр: повреждённая база сохранена как .bad ({_bad_copies})")

# Битый старый JSON-реестр при переносе: файл не трогаем, реестр пуст.
with open(storage.legacy_json_path(), "w", encoding="utf-8") as _fh:
    _fh.write("{ это не JSON !!!")
_legacy = storage.DataStore()
check(bool(_legacy.load_warning), "реестр: при битом legacy-JSON предупреждение")
check(not _legacy.counterparties and not _legacy.documents,
      "реестр: перенос не выполнен, запуск не падает")
check(os.path.exists(storage.legacy_json_path()),
      "реестр: повреждённый JSON сохранён без изменений")

# null вместо списков не роняет запуск и не блокирует перенос (п. 1.4).
with open(storage.legacy_json_path(), "w", encoding="utf-8") as _fh:
    _fh.write('{"counterparties": null, "documents": null}')
_null_store = storage.DataStore()
check(bool(_null_store.load_warning) and not _null_store.counterparties,
      "реестр: null вместо разделов — предупреждение без падения")

# Валидный legacy-JSON переносится в базу при первом запуске.
with open(storage.legacy_json_path(), "w", encoding="utf-8") as _fh:
    json.dump({"schema_version": storage.SCHEMA_VERSION, "seq": 4,
               "counterparties": [{"id": 3, "type": "ooo",
                                   "fields": {"ИНН": "7701234567"}}],
               "documents": []}, _fh, ensure_ascii=False)
_migrated = storage.DataStore()
check([rec["id"] for rec in _migrated.counterparties] == [3],
      "реестр: данные перенесены из старого JSON в базу")
check(os.path.exists(storage.legacy_json_path() + ".migrated"),
      "реестр: перенесённый JSON сохранён как .migrated")
_reloaded = storage.DataStore()
check([rec["id"] for rec in _reloaded.counterparties] == [3],
      "реестр: перенесённые данные читаются из базы повторным запуском")

# ---------- Сбой записи реестра (п. 1.5) ----------
_ro = storage.DataStore()
_ro_party = _ro.add_counterparty("ooo", {"ИНН": "7709999999"}, save=False)
os.remove(storage.db_path())   # файл создаётся при загрузке — убираем
os.makedirs(storage.db_path())  # каталог вместо файла: база не открывается
check(_ro.save() is False and _ro.last_error,
      "save(): сбой записи возвращает False и last_error, а не ложный успех")
_upd_id = _ro_party["id"]
check(_ro.update_counterparty(_upd_id, {"ИНН": "7700000001"}) is None,
      "update_counterparty: при сбое записи правки откатываются (None)")
check(_ro.get_counterparty(_upd_id)["fields"]["ИНН"] == "7709999999",
      "update_counterparty: данные не изменились после сбоя")
os.rmdir(storage.db_path())
os.environ["APPDATA"] = _appdata_main

# ---------- Атомарная запись .docx и настроек (п. 1.7) ----------
_atomic_dir = os.path.join(tempfile.gettempdir(), "smoke_atomic")
os.makedirs(_atomic_dir, exist_ok=True)
_atomic_ok = os.path.join(_atomic_dir, "Документ.docx")
m.save_docx_atomic(_Doc(), _atomic_ok)
check(os.path.exists(_atomic_ok) and not os.path.exists(_atomic_ok + ".tmp"),
      "atomic: .docx записан, tmp-файл не остался")

class _BrokenDoc:
    def save(self, path):
        with open(path, "wb") as _fh:
            _fh.write("это не zip-архив".encode("utf-8"))

_broken_target = os.path.join(_atomic_dir, "Битый.docx")
try:
    m.save_docx_atomic(_BrokenDoc(), _broken_target)
    check(False, "atomic: повреждённый docx должен отклоняться")
except Exception:
    check(True, "atomic: повреждённый docx отклонён с ошибкой")
check(not os.path.exists(_broken_target) and not os.path.exists(_broken_target + ".tmp"),
      "atomic: при сбое нет ни итогового файла, ни tmp")

check(m.write_settings({"theme": "dark"}), "настройки: атомарная запись")
check(m.load_settings() == {"theme": "dark"}, "настройки: прочитаны после записи")
os.remove(m.settings_path())
os.makedirs(m.settings_path())  # каталог вместо файла: os.replace падает
check(not m.write_settings({"theme": "light"}), "настройки: сбой записи -> False")
check(not m.load_settings(), "настройки: при сбое читаются пустыми без падения")
os.rmdir(m.settings_path())

# ---------- Сверка «реестр ↔ папка документов» (п. 2.4) ----------
_reconcile_root = os.path.join(tempfile.gettempdir(), "smoke_reconcile")
os.makedirs(os.path.join(_reconcile_root, "ООО Ромашка"), exist_ok=True)
_orphan_path = os.path.join(_reconcile_root, "ООО Ромашка", "Договор поставки № 99.docx")
with open(_orphan_path, "wb") as _fh:
    _fh.write(b"stub")
_orphans = reconcile.find_orphans(_reconcile_root, store)
check(_orphan_path in _orphans, "сверка: файл вне реестра найден")
_restored = reconcile.restore_orphans(store, _reconcile_root, _orphans, m.safe_filename)
check(_restored == 1, "сверка: файл восстановлен в реестр")
_restored_rec = [rec for rec in store.documents
                 if (rec.get("extra") or {}).get("restored")]
check(bool(_restored_rec) and _restored_rec[0]["kind"] == storage.KIND_CONTRACT,
      "сверка: тип документа определён по имени файла")
check(_restored_rec[0]["party_id"] == party_rec["id"],
      "сверка: контрагент сопоставлен по имени папки")
check(not reconcile.find_orphans(_reconcile_root, store),
      "сверка: после восстановления файл не считается сиротой")
store.purge_document(_restored_rec[0]["id"])
try:
    os.remove(_orphan_path)
except OSError:
    pass

# ---------- Импорт готовых документов ----------
check(di.F is m.F, "импорт: единый источник имён полей (9.6 — F из document_builder)")

check(di.detect_kind("СПЕЦИФИКАЦИЯ № 12 к Договору поставки") == storage.KIND_SPEC,
      "импорт: тип — спецификация")
check(di.detect_kind("ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ № 1") == storage.KIND_ADDITIONAL,
      "импорт: тип — доп. соглашение")
check(di.detect_kind("Договор поставки № 516") == storage.KIND_CONTRACT,
      "импорт: тип — договор поставки")

_import_docx = os.path.join(tempfile.gettempdir(), "smoke_import_contract.docx")
_doc = _Doc()
_doc.add_paragraph("ДОГОВОР ПОСТАВКИ № 777")
_doc.add_paragraph("г. Владивосток")
_doc.add_paragraph("«10» января 2026 г.")
_doc.add_paragraph("ООО «Ромашка-Импорт», именуемое в дальнейшем «Покупатель»,")
_doc.add_paragraph("ИНН 2540123456, КПП 254001001, ОГРН 1035001234567")
_doc.add_paragraph("Адрес: 690000, г. Владивосток, ул. Светланская, д. 1")
_doc.add_paragraph("Тел.: +7 (423) 123-45-67, e-mail: buy@romashka-import.ru")
_doc.add_paragraph("Р/с 40702810123456789012 в ПАО Банк, БИК 044525225, к/с 30101810100000000225")
_doc.add_paragraph("Генеральный директор Петров Пётр Петрович")
_doc.save(_import_docx)

_res = di.analyze_file(_import_docx)
check(_res["ok"] and _res["kind"] == storage.KIND_CONTRACT, "импорт: договор опознан")
check(_res["number"] == "777", "импорт: номер договора 777")
check(_res["date"] == "10.01.2026", "импорт: дата договора 10.01.2026")
check(_res["party"] and _res["party"]["type"] == "ooo", "импорт: контрагент — ООО")
check(_res["party"]["fields"]["ИНН"] == "2540123456", "импорт: ИНН контрагента")
check(_res["party"]["fields"]["КПП"] == "254001001", "импорт: КПП контрагента")
check(_res["party"]["fields"]["ОГРН"] == "1035001234567", "импорт: ОГРН контрагента")
check(_res["party"]["fields"]["Расчётный счёт"] == "40702810123456789012", "импорт: расчётный счёт")
check(_res["party"]["fields"]["ФИО директора"] == "Петров Пётр Петрович", "импорт: директор")
check("Ромашка-Импорт" in _res["party"]["display"], "импорт: название контрагента")

_unrecognized = os.path.join(tempfile.gettempdir(), "smoke_import_unknown.docx")
_doc2 = _Doc()
_doc2.add_paragraph("Список покупок: молоко, хлеб, сыр.")
_doc2.save(_unrecognized)
_res2 = di.analyze_file(_unrecognized)
check(not _res2["ok"] and _res2["kind"] == storage.KIND_UNKNOWN, "импорт: нераспознанный файл -> unknown")
check("опознать" in _res2["error"].lower(), f"импорт: причина отказа понятна ({_res2['error']})")

# ---------- Качество импорта: счета, названия, критерий recognized (этап 3) ----------
_party_text = (
    "ООО «Ромашка-Импорт», именуемое в дальнейшем «Покупатель»,\n"
    "ИНН 2540123456, КПП 254001001\n"
    "Р/счет № 40702810123456789012 в ПАО Банк\n"
    "к/счет № 30101810100000000225\n"
    "БИК 044525225"
)
_party = di.extract_party(_party_text)
check(_party is not None and _party["recognized"], "импорт: карточка с ИНН распознана")
check(_party["fields"][m.F.ACCOUNT] == "40702810123456789012",
      f"счёт: «Р/счет №» извлечён ({_party['fields'][m.F.ACCOUNT]!r})")
check(_party["fields"][m.F.CORR_ACCOUNT] == "30101810100000000225",
      "счёт: корр. счёт извлечён")
check(_party["fields"][m.F.ACCOUNT] != _party["fields"][m.F.CORR_ACCOUNT],
      "счёт: корр. счёт не попал в расчётный")

_account_variants = (
    ("Р/с 40702810123456789012", "р/с"),
    ("Р/счёт № 40702810123456789012", "р/счёт"),
    ("р/сч 40702810123456789012", "р/сч"),
    ("расчётный счёт 40702810123456789012", "расчётный счёт"),
    ("расч. счёт 40702810123456789012", "расч. счёт"),
    ("р.с. 40702810123456789012", "р.с."),
    ("счет № 40702810123456789012", "счет №"),
)
for _line, _label in _account_variants:
    _variant = di.extract_party(f"ООО «Тест», ИНН 7701234567\n{_line}")
    check(_variant["fields"][m.F.ACCOUNT] == "40702810123456789012",
          f"счёт: вариант «{_label}» извлечён")

_cell_account = di.extract_party(
    "ООО «Тест», ИНН 7701234567\nБИК 044525225\n40702810999999999991"
)
check(_cell_account["fields"][m.F.ACCOUNT] == "40702810999999999991",
      "счёт: 20 цифр в отдельной ячейке без подписи")

_named = di.extract_party(
    'ООО "ЛИДЕР ДВ", именуемый в дальнейшем «Покупатель», в лице директора '
    "Иванова Ивана Ивановича"
)
check(_named["display"] == "ООО «ЛИДЕР ДВ»",
      f"название: хвост «именуемый» обрезан ({_named['display']!r})")
check("именуем" not in _named["display"].lower() and "в лице" not in _named["display"].lower(),
      "название: служебные фразы не попали в карточку")

_no_inn = di.extract_party(
    "ООО «Ромашка», именуемое в дальнейшем «Покупатель»\n"
    "Юридический адрес: г. Москва, ул. Ленина, д. 1"
)
check(_no_inn is not None and not _no_inn["recognized"],
      "recognized: карточка без ИНН не создаётся автоматически")
check(_no_inn["missing"] == "ИНН", "recognized: причина — не найден ИНН")

_orig_file_limit = di.MAX_FILE_SIZE
di.MAX_FILE_SIZE = 10
try:
    try:
        di.check_file_limits(_import_docx)
        check(False, "лимиты: файл больше лимита должен отклоняться")
    except ValueError as _exc:
        check("слишком большой" in str(_exc).lower(), "лимиты: причина отказа понятна")
finally:
    di.MAX_FILE_SIZE = _orig_file_limit

_orig_zip_limit = di.MAX_ZIP_UNCOMPRESSED_SIZE
di.MAX_ZIP_UNCOMPRESSED_SIZE = 100
try:
    try:
        di.check_file_limits(_import_docx)
        check(False, "лимиты: ZIP-бомба должна отклоняться")
    except ValueError as _exc:
        check("архивная бомба" in str(_exc).lower(), "лимиты: ZIP-бомба отклонена")
finally:
    di.MAX_ZIP_UNCOMPRESSED_SIZE = _orig_zip_limit

_bad_ext = os.path.join(tempfile.gettempdir(), "smoke_import_bad.txt")
with open(_bad_ext, "w", encoding="utf-8") as _fh:
    _fh.write("просто текст")
_bad_res = di.analyze_file(_bad_ext)
check(not _bad_res["ok"] and "Неподдерживаемый формат" in _bad_res["error"],
      "импорт: неподдерживаемый формат — понятная причина")

# Потенциально опасные файлы запрещены на входе (п. 1.2): ни анализа,
# ни копирования — .exe/.bat/.cmd/.vbs/.lnk в программу не попадают.
for _ext, _label in (
    (".exe", "исполняемый"), (".bat", "batch"), (".cmd", "cmd"),
    (".vbs", "скрипт"), (".lnk", "ярлык"),
):
    _danger_path = os.path.join(tempfile.gettempdir(), "smoke_import_danger" + _ext)
    with open(_danger_path, "wb") as _fh:
        _fh.write(b"MZ stub")
    _danger_res = di.analyze_file(_danger_path)
    check(not _danger_res["ok"], f"импорт: {_ext} ({_label}) отклонён")
    try:
        os.remove(_danger_path)
    except OSError:
        pass

_sample_additional = os.path.join("образцы", "Доп. соглашение ООО «ЛИДЕР ДВ».docx")
if os.path.exists(_sample_additional):
    _sample_res = di.analyze_file(_sample_additional)
    check(_sample_res["party"]["display"] == "ООО «ЛИДЕР ДВ»",
          f"образец доп. соглашения: название без хвоста ({_sample_res['party']['display']!r})")
    check(not _sample_res["recognized"] and _sample_res["warning"],
          "образец доп. соглашения: карточка требует подтверждения (нет ИНН)")

# ---------- Лимиты и защита парсера (п. 3.4) — реальная zip-бомба ----------
_zb_path = os.path.join(tempfile.gettempdir(), "smoke_zipbomb.docx")
with zipfile.ZipFile(_zb_path, "w", zipfile.ZIP_DEFLATED) as _zb:
    _zb.writestr("word/document.xml", b"\0" * (30 * 1024 * 1024))
_zb_bytes = os.path.getsize(_zb_path)
try:
    di.check_file_limits(_zb_path)
    check(False, "3.4: реальная zip-бомба должна отклоняться")
except ValueError as _exc:
    check("сжат подозрительно сильно" in str(_exc).lower(),
          f"3.4: zip-бомба {_zb_bytes} Б -> 30 МБ распаковки отклонена ({_exc})")

_orig_member_limit = di.MAX_ZIP_MEMBERS
di.MAX_ZIP_MEMBERS = 2
try:
    try:
        di.check_file_limits(_import_docx)
        check(False, "3.4: превышение числа частей должно отклоняться")
    except ValueError as _exc:
        check("част" in str(_exc).lower(),
              "3.4: офисный архив с избытком частей отклонён")
finally:
    di.MAX_ZIP_MEMBERS = _orig_member_limit
check(di.MAX_FILES_PER_BATCH > 0 and di.MAX_TOTAL_SIZE > 0,
      "3.4: лимиты пакета импорта настроены")
check("IMPORT_MAX_FILES" in _main_src and "IMPORT_MAX_TOTAL_SIZE" in _main_src
      and "disk_usage" in _app_src,
      "3.4: пакет импорта и свободное место проверяются в приложении")

# ---------- Причины отказа импорта (п. 3.5) ----------
_broken_office = os.path.join(tempfile.gettempdir(), "smoke_broken_office.docx")
with open(_broken_office, "wb") as _fh:
    _fh.write(b"this is definitely not a zip archive")
_broken_res = di.analyze_file(_broken_office)
check(not _broken_res["ok"] and "повреждён" in _broken_res["error"].lower(),
      f"3.5: битый офисный файл — понятная причина ({_broken_res['error']})")

check(store.find_by_name("ООО «Ромашка»") is not None,
      "хранилище: поиск контрагента по названию без учёта кавычек")
check(store.find_by_name("нет такого") is None,
      "хранилище: поиск по названию не находит лишнего")

# ---------- п.10/5.1: производительность реестра на 600 контрагентах ----------
_perf_dir = tempfile.mkdtemp(prefix="smoke_perf_")
_appdata_perf_backup = os.environ["APPDATA"]
os.environ["APPDATA"] = _perf_dir
try:
    _t0 = time.perf_counter()
    _perf_store = storage.DataStore()
    for _i in range(600):
        _perf_store.add_counterparty("ooo", {
            "ИНН": f"7701{_i:06d}",
            "Сокращенное название компании": f"ООО Нагрузка-{_i}",
        }, save=False)
    _perf_store.save()
    _t1 = time.perf_counter()
    _perf_store = storage.DataStore()
    _t2 = time.perf_counter()
    check(len(_perf_store.counterparties) == 600,
          "п.10: реестр на 600 контрагентов загружается целиком")
    check((_t1 - _t0) < 5.0,
          f"п.10: запись 600 контрагентов одним save() заняла {(_t1 - _t0):.2f} с (<5 с)")
    check((_t2 - _t1) < 5.0,
          f"п.10: загрузка 600 контрагентов заняла {(_t2 - _t1):.2f} с (<5 с)")
finally:
    os.environ["APPDATA"] = _appdata_perf_backup
    shutil.rmtree(_perf_dir, ignore_errors=True)

# ---------- Интерфейс ----------
app = m.ContractApp()
try:
    app.root.update_idletasks()

    # п.9: при старте реестр не отрисовывается — ленивый рендер.
    check(getattr(app, "_registry_rendered", True) is False,
          "п.9: при старте реестр не отрисовывается (ленивый рендер)")
    # 6.2: открытие файла подменяем детерминированным стабом (без xdg-open).
    _orig_open_file = m.open_file
    m.open_file = lambda path: True

    check(hasattr(app, "nav_holder") and len(app.nav_items) == 6,
          "кнопки разделов в шапке, включая «Настройки»")
    check(set(app.pages) == {"contract", "additional", "spec", "registry", "counterparty", "settings"},
          f"страницы: {sorted(app.pages)}")
    check(hasattr(app, "contract_preview") and hasattr(app, "addl_preview")
          and hasattr(app, "spec_text"),
          "панели предпросмотра: договор и доп. соглашение — есть (спецификация — есть)")
    check(hasattr(app, "status_label"), "строка статуса в футере")
    check(app.save_button.cget("text") == "Сохранить договор", "кнопка футера: договор")

    # Регрессия: переключатель «Юр. лицо / ИП» должен обновлять _cp_type_var.
    # Раньше переменная не была привязана к сегменту (variable=), тип оставался
    # «Юр. лицо»: «Заполнить по ИНН» у ИП читала пустое скрытое поле юрлица и
    # требовала 10 цифр — поиск по ИНН для ИП не работал.
    _seg = getattr(app, "_cp_type_segmented", None)
    check(_seg is not None and app._cp_type_var.get() == "Юр. лицо",
          "контрагент: сегмент типа найден, по умолчанию «Юр. лицо»")
    if _seg is not None:
        _ip_button = getattr(_seg, "_buttons_dict", {}).get("ИП")
        if _ip_button is not None:
            # Имитируем реальный клик по сегменту «ИП».
            _ip_button.invoke()
            app.root.update()
            check(app._cp_type_var.get() == "ИП",
                  "контрагент: клик по сегменту «ИП» обновляет тип карточки")

            # «Заполнить по ИНН» в режиме ИП: подменяем DaData и проверяем
            # подстановку карточки ЕГРИП в поля ИП.
            #
            # Поток подменяем синхронным исполнителем: в smoke-среде нет
            # mainloop(), и root.after() из рабочего потока падает с
            # RuntimeError — в работающем приложении поток реален.
            _orig_key, _orig_fetch = m.dadata.get_api_key, m.dadata.fetch_party
            _orig_thread = m.threading.Thread

            class _SyncThread:
                def __init__(self, target=None, args=(), daemon=None, **_kw):
                    self._target, self._args = target, args

                def start(self):
                    self._target(*self._args)

            m.dadata.get_api_key = lambda: "test-key"
            m.threading.Thread = _SyncThread
            _ip_card = {
                "full_name": "ИП Петров Пётр Петрович",
                "short_name": "ИП Петров",
                "director": "Петров Пётр Петрович",
                "legal_address": "г. Москва, ул. Тестовая, д. 1, оф. 2",
                "inn": "312301234567", "kpp": "", "ogrn": "304500100000340",
                "bank_name": "Тест-Банк", "bik": "044525000",
                "correspondent_account": "30101810000000000225",
            }
            m.dadata.fetch_party = lambda inn, key: _ip_card
            try:
                _inn_widget = app._cp_entries["ip"][m.F.INN]
                _inn_widget.delete(0, "end")
                _inn_widget.insert(0, "312301234567")
                ALERTS.clear()
                app._dadata_flow.fill_from_inn()
                for _ in range(60):
                    app.root.update()
                    time.sleep(0.02)
                check(not any("ИНН должен содержать" in str(a) for a in ALERTS),
                      "контрагент ИП: «Заполнить по ИНН» не требует 10 цифр")
                check(app._cp_entries["ip"][m.F.IP_NAME].get().strip() == "Петров Пётр Петрович"
                      and app._cp_entries["ip"][m.F.OGRNIP].get().strip() == "304500100000340"
                      and app._cp_entries["ip"][m.F.LEGAL_ADDRESS].get().strip() != "",
                      "контрагент ИП: реквизиты ЕГРИП подставлены в поля ИП")
                check(app._cp_status_var.get() == "Реквизиты заполнены из ЕГРЮЛ/ЕГРИП",
                      "контрагент ИП: статус подтверждает заполнение")
            finally:
                m.dadata.get_api_key, m.dadata.fetch_party = _orig_key, _orig_fetch
                m.threading.Thread = _orig_thread
                ALERTS.clear()
                # Очищаем поля ИП и статус, чтобы форма осталась «чистой»
                # для последующих проверок несохранённых данных.
                for _key in (m.F.INN, m.F.IP_NAME, m.F.OGRNIP, m.F.LEGAL_ADDRESS,
                             m.F.BANK, m.F.BIK, m.F.CORR_ACCOUNT):
                    _widget = app._cp_entries["ip"].get(_key)
                    if _widget is not None:
                        _widget.delete(0, "end")
                app._cp_status_var.set("")
                app.root.update()

            # Обратно на «Юр. лицо»: тип снова «Юр. лицо» (режим ooo как раньше).
            _seg._buttons_dict["Юр. лицо"].invoke()
            app.root.update()
            check(app._cp_type_var.get() == "Юр. лицо",
                  "контрагент: возврат на «Юр. лицо» обновляет тип")

    # ---------- Аудит 2.0.1: re-entry импорта и повторные клики DaData ----------
    _guard_thread_orig = m.threading.Thread
    _guard_threads: list = []

    class _CountingThread:
        def __init__(self, target=None, args=(), daemon=None, **_kw):
            _guard_threads.append(self)
            self._target, self._args = target, args

        def start(self):
            pass  # в guard-проверках воркеры не должны даже стартовать

    m.threading.Thread = _CountingThread
    try:
        _guard_docx = os.path.join(tempfile.gettempdir(), "smoke_guard.docx")
        _Doc().save(_guard_docx)
        _threads_before = len(_guard_threads)

        # Первый импорт «выполняется»: второй запуск (как при drag-and-drop)
        # обязан быть заблокирован, не сбрасывая состояние очереди.
        app._import_flow.running = True
        app._import_flow.queue = None
        app._import_flow.win = None
        app._import_flow.start([_guard_docx])
        check(len(_guard_threads) == _threads_before
              and app._import_flow.queue is None and app._import_flow.win is None,
              "импорт: повторный запуск во время первого заблокирован, "
              "очередь первого не тронута")
        app._import_flow.running = False

        app._import_flow.start([])
        check(len(_guard_threads) == _threads_before,
              "импорт: пустой список файлов не запускает воркер")

        # Повторный клик «Заполнить по ИНН» при активном запросе игнорируется.
        app._dadata_flow.busy = True
        app._dadata_flow.fill_from_inn()
        check(len(_guard_threads) == _threads_before and app._dadata_flow.busy,
              "dadata: повторный клик при активном запросе не порождает воркера")
        app._dadata_flow.busy = False
    finally:
        m.threading.Thread = _guard_thread_orig
        ALERTS.clear()
        app.root.update()


    # Контрагент появляется в списке страницы договора.
    app._doc_context_flow.refresh_party_combo_values()
    app.root.update()
    app.contract_party_var.set(storage.display_with_inn(party_rec))
    app.root.update()
    check("ООО Ромашка" in app.contract_party_var.get(), "выбор контрагента в списке")

    app.contract_number_var.set("5")
    app.contract_day_var.set("12")
    app.contract_month_var.set("января")
    app.root.update()
    app._doc_page_flow.render_contract_preview()
    check("Ромашка" in app.contract_preview.get("1.0", "end"),
          "предпросмотр договора: данные контрагента видны")
    data = dict(party_rec["fields"])
    data.update({m.F.NUMBER: "5", m.F.DAY: "12", m.F.MONTH: "января"})
    doc = app._doc_save_flow.render_contract_docx(data, is_ip=False)
    contract_text = "\n".join(p.text for p in doc.paragraphs)
    check("ООО Ромашка" in contract_text and "Иванов" in contract_text,
          "печатная форма договора: данные контрагента подставлены")
    check("Ромашка" in contract_text and "5" in contract_text,
          "печатная форма договора: номер подставлен")
    check(len(doc.tables) == 1, "печатная форма договора: таблица реквизитов на месте")

    # 5.5: шаблон разбирается один раз и кэшируется, рендер получает копию.
    check(bool(m._TEMPLATE_CACHE), "5.5: шаблоны печатных форм кэшируются")
    _cached = m._TEMPLATE_CACHE.get(m.resource_path("dogovor_template.docx"))
    check(_cached is not None and _cached[0] > 0,
          "5.5: в кэше шаблона хранятся mtime и разобранный Document")
    _doc_again = app._doc_save_flow.render_contract_docx(data, is_ip=False)
    check("{{" not in "\n".join(p.text for p in _doc_again.paragraphs),
          "5.5: повторный рендер из кэша даёт чистый документ без {{...}}")

    app._doc_save_flow.save_contract()
    check(len(app.store.documents) == 1 and app.store.documents[0]["kind"] == storage.KIND_CONTRACT,
          "договор сохранён в реестре")

    _saved_doc = app.store.documents[0]
    check(app._import_flow.document_is_duplicate(party_rec["id"], {
        "kind": _saved_doc["kind"],
        "number": _saved_doc["number"],
        "date": _saved_doc["date"],
    }), "импорт: повторный документ распознан как дубль")
    check(not app._import_flow.document_is_duplicate(party_rec["id"], {
        "kind": _saved_doc["kind"],
        "number": "999",
        "date": _saved_doc["date"],
    }), "импорт: другой номер — не дубль")
    check(not app._import_flow.document_is_duplicate(party_rec["id"], {
        "kind": storage.KIND_SPEC, "number": "", "date": "",
    }), "импорт: без номера и даты дубль не определяется")

    # Страница «Доп. соглашение»: договор-основание подтягивается из реестра.
    app._select_page("additional")
    app.root.update()
    check(app.save_button.cget("text") == "Сохранить доп. соглашение",
          "кнопка футера: доп. соглашение")
    app.addl_party_var.set(storage.display_with_inn(party_rec))
    app.root.update()
    check("№ 5" in app.addl_contract_var.get(), f"список договоров: {app.addl_contract_var.get()!r}")
    doc2 = app._doc_save_flow.render_additional_docx(data, is_ip=False, number="1",
                                       limit_display="1 500 000", term="45")
    additional_text = "\n".join(p.text for p in doc2.paragraphs)
    check("Ромашка" in additional_text and "№ 1" in additional_text,
          "печатная форма доп. соглашения сформирована")
    check(len(doc2.tables) == 1, "печатная форма доп. соглашения: таблица подписей на месте")

    app._limit_var.set("1500000")
    check(app._limit_var.get() == "1 500 000", f"лимит форматирован: {app._limit_var.get()!r}")
    check("пятьсот тысяч" in app.limit_words_label.cget("text"), "лимит прописью")
    app.term_var.set("45")
    check("сорок пять" in app.words_label.cget("text"), "срок прописью")
    app._doc_page_flow.render_additional_preview()
    check("Ромашка" in app.addl_preview.get("1.0", "end"),
          "предпросмотр доп. соглашения: данные контрагента видны")
    app._doc_save_flow.save_additional()
    check(sum(1 for d in app.store.documents if d["kind"] == storage.KIND_ADDITIONAL) == 1,
          "доп. соглашение сохранено в реестре")

    # Страница «Спецификация».
    app._select_page("spec")
    app.root.update()
    check(app.save_button.cget("text") == "Сохранить спецификацию",
          "кнопка футера: спецификация")
    app.spec_party_var.set(storage.display_with_inn(party_rec))
    app.root.update()
    app.spec_number_var.set("1")
    app.spec_day_var.set("20")
    row = app._doc_save_flow.spec_rows[0]
    row["name"].set("Товар")
    row["qty"].set("10")
    row["unit"].set("шт.")
    row["price"].set("100")
    app.root.update()
    # 5.6: предпросмотр спецификации отложен — дособираем его синхронно.
    app._doc_save_flow.flush_spec_preview()
    spec_text = app.spec_text.get("1.0", "end")
    check("СПЕЦИФИКАЦИЯ № 1" in spec_text, "заголовок спецификации")
    check("1 000" in spec_text, "итог спецификации посчитан")
    check("одна тысяча" in spec_text.lower(), "итог прописью")
    check("шт." in spec_text, "ед. изм. видна в предпросмотре")

    # 5.6: пересборка откладывается, а не идёт на каждый символ.
    app._spec_preview_after = None
    app.spec_delivery_term_var.set("45")
    check(getattr(app, "_spec_preview_after", None) is not None,
          "5.6: изменение поля планирует отложенную пересборку предпросмотра")
    app._doc_save_flow.flush_spec_preview()
    check("отсрочка" in app.spec_text.get("1.0", "end").lower()
          or "45" in app.spec_text.get("1.0", "end"),
          "5.6: flush дособирает предпросмотр актуальным текстом")
    app.spec_delivery_term_var.set("")
    app._doc_save_flow.flush_spec_preview()
    app._doc_save_flow.save_spec()
    check(sum(1 for d in app.store.documents if d["kind"] == storage.KIND_SPEC) == 1,
          "спецификация сохранена в реестре")

    # Приёмка п. 1.1: у ИП печатается ОГРНИП, нет КПП и пустых подписей.
    def _doc_full_text(document):
        _chunks = [p.text for p in document.paragraphs]
        for _t in document.tables:
            for _row in _t.rows:
                for _cell in _row.cells:
                    _chunks.append(_cell.text)
        for _section in document.sections:
            for _part in (_section.header, _section.footer):
                _chunks.extend(p.text for p in _part.paragraphs)
        return "\n".join(_chunks)

    _ip_party = app.store.add_counterparty("ip", {
        m.F.IP_NAME: "Иванов Иван Иванович",
        m.F.INN: "253612345678",
        m.F.OGRNIP: "304253600012345",
        m.F.LEGAL_ADDRESS: "690074, г. Владивосток, ул. Ленина, д. 5",
        m.F.FACT_ADDRESS: "690074, г. Владивосток, ул. Ленина, д. 5",
        m.F.EMAIL: "ip@mail.ru",
        m.F.PHONE: "+7 (912) 345-67-89",
        m.F.ACCOUNT: "40802810500000012345",
        m.F.CORR_ACCOUNT: "30101810800000000770",
        m.F.BANK: "ФИЛИАЛ «ХАБАРОВСКИЙ» АО «АЛЬФА-БАНК»",
        m.F.BIK: "040813770",
    }, save=False)
    _ip_spec_rec = {
        "kind": storage.KIND_SPEC,
        "party_id": _ip_party["id"],
        "number": "1",
        "data": dict(_ip_party["fields"]),
        "extra": {
            "items": [{"name": "Товар", "qty": 10.0, "unit": "шт.", "price": 100.0}],
            "total": 1000.0,
            "spec_day": "20",
            "spec_month": "ноября",
            "is_ip": True,
        },
        "text": "",
    }
    _ip_doc = app._doc_save_flow.render_spec_docx(_ip_spec_rec)
    check("{{" not in _doc_full_text(_ip_doc),
          "спецификация ИП: в документе нет незаполненных {{...}}")
    _ip_buyer = next(
        (cell.text
         for table in _ip_doc.tables
         for row in table.rows
         for cell in row.cells if "Покупателя" in cell.text),
        "",
    )
    check(bool(_ip_buyer), "спецификация ИП: блок реквизитов покупателя найден")
    check("КПП" not in _ip_buyer, "спецификация ИП: в реквизитах покупателя нет КПП")
    check(re.search(r"\bОГРН\b", _ip_buyer) is None,
          "спецификация ИП: подпись «ОГРН» заменена на ОГРНИП")
    check("ОГРНИП 304253600012345" in _ip_buyer,
          "спецификация ИП: печатается значение ОГРНИП")

    # Печатная форма ООО не пострадала (КПП и ОГРН на месте).
    _ooo_spec_rec = next(d for d in app.store.documents if d["kind"] == storage.KIND_SPEC)
    _ooo_text = _doc_full_text(app._doc_save_flow.render_spec_docx(_ooo_spec_rec))
    check("{{" not in _ooo_text, "спецификация ООО: нет незаполненных {{...}}")
    check("770101001" in _ooo_text and "1037739010221" in _ooo_text,
          "спецификация ООО: КПП и ОГРН подставлены как раньше")

    # ИП-карточка нужна была только для проверки печатной формы — убираем её,
    # чтобы не мешать проверкам корзины ниже.
    app.store.counterparties.remove(_ip_party)

    # Порядок «файл → запись» (п. 1.6): сбой реестра откатывает и запись, и файл.
    # Теперь проверяется настоящая транзакция сервисного слоя (п. 9.2),
    # а не отдельный хелпер отката.
    _rb_dir = tempfile.mkdtemp(prefix="smoke_rollback_")
    _orig_save = app.store.save

    def _rb_failing_save():  # имитация сбоя записи реестра на диске
        app.store.last_error = "имитация: диск недоступен"
        return False

    app.store.save = _rb_failing_save
    try:
        _rb_result = m.services.save_document(app.store, {
            "kind": storage.KIND_CONTRACT,
            "party_id": party_rec["id"],
            "number": "RB-1",
            "date": "12 января",
            "data": dict(party_rec.get("fields", {})),
            "extra": {"day": "12", "month": "января", "is_ip": False},
            "text": "",
        }, _rb_dir)
    finally:
        app.store.save = _orig_save
        app.store.last_error = None
    check(not _rb_result["ok"] and _rb_result["record"] is None
          and not os.listdir(_rb_dir),
          "файл→запись: при сбое реестра документ откатывается целиком")
    shutil.rmtree(_rb_dir, ignore_errors=True)

    # Регистрация импорта: документ без ИНН привязывается к подтверждённому
    # контрагенту, а не создаёт новую карточку (п. 3.3).
    _tmp_copied = os.path.join(tempfile.gettempdir(), "smoke_import_copy.docx")
    with open(_tmp_copied, "wb") as _fh:
        _fh.write(b"stub")
    _orig_report = app._import_flow.show_report
    _orig_ask = app._import_flow.ask_party
    app._import_flow.show_report = lambda entries, summary: None
    app._import_flow.ask_party = lambda result: (party_rec, "attach")
    try:
        _parties_before = len(app.store.counterparties)
        app._import_flow.finish([{
            "name": "Импорт без ИНН.docx",
            "kind": storage.KIND_ADDITIONAL,
            "number": "77",
            "date": "01.02.2026",
            "party": {
                "recognized": False, "type": "ooo",
                "display": "ООО «ЛИДЕР ДВ»",
                "fields": {"Сокращенное название компании": "ООО «ЛИДЕР ДВ»"},
                "missing": "ИНН",
            },
            "text": "", "ok": True, "recognized": False,
            "copied": _tmp_copied, "error": "", "warning": "не найден ИНН",
        }])
        _imported = [
            doc for doc in app.store.documents
            if (doc.get("extra") or {}).get("source_file") == _tmp_copied
        ]
        check(bool(_imported) and _imported[0]["party_id"] == party_rec["id"],
              "импорт: документ без ИНН привязан к подтверждённому контрагенту")
        check(len(app.store.counterparties) == _parties_before,
              "импорт: карточка-дубль не создана")
    finally:
        app._import_flow.show_report = _orig_report
        app._import_flow.ask_party = _orig_ask
        try:
            os.remove(_tmp_copied)
        except OSError:
            pass

    # Таблица позиций: вставка блока из буфера (эмуляция копирования из Excel).
    check(app.spec_table._map_row(["Болт М8", "5", "шт.", "200"]) ==
          ["Болт М8", "5", "шт.", "200"], "таблица: 4 столбца раскладываются по колонкам")
    check(app.spec_table._map_row(["Гайка", "12", "50"]) == ["Гайка", "12", "", "50"],
          "таблица: 3 числовых столбца -> наименование, кол-во, цена")
    check(app.spec_table._normalize_number("1 000,50") == "1000.50",
          "таблица: нормализация чисел с пробелами и запятой (копейки сохраняются)")
    app.root.clipboard_clear()
    app.root.clipboard_append("Болт М8\t5\tшт.\t200\nГайка М10\t12\tшт.\t50\n\n")
    check(app.spec_table._paste_block((0, 0)), "таблица: блок из буфера вставлен")
    app.root.update()
    items = app._doc_save_flow.spec_items()
    check(len(items) == 2 and items[0]["name"] == "Болт М8" and items[1]["qty"] == 12,
          "таблица: строки из буфера попали в позиции")
    app._doc_save_flow.flush_spec_preview()
    check("1 600" in app.spec_text.get("1.0", "end"), "таблица: итог пересчитан после вставки")
    check(len(app._doc_save_flow.spec_rows) == 10, "таблица: стартовое количество строк — 10")

    # Задание точного количества строк.
    app.spec_table.set_row_count(20)
    app.root.update()
    check(len(app._doc_save_flow.spec_rows) == 20, "таблица: количество строк задано полем «Строк»")
    app.spec_table.set_row_count(4)
    check(len(app._doc_save_flow.spec_rows) == 4, "таблица: лишние строки удаляются")
    check(app._doc_save_flow.spec_rows[-1]["num"].cget("text") == "4", "таблица: нумерация строк актуальна")

    # Аудит 2.0.1: валидаторы ячеек включены флагом validate="key" — раньше
    # Tk создавал validatecommand, но никогда его не вызывал, и в «Кол-во»/
    # «Цена» можно было ввести буквы.
    _qty_cell = app.spec_table.rows[0]["cells"]["qty"]
    _price_cell = app.spec_table.rows[0]["cells"]["price"]
    check(str(_qty_cell.cget("validate")) == "key"
          and str(_price_cell.cget("validate")) == "key",
          "таблица: валидаторы «Кол-во»/«Цена» включены (validate=key)")
    check(m.SpecTable._sanitize_value("qty", "abc") == ""
          and m.SpecTable._sanitize_value("qty", "12,5") == "12.5"
          and m.SpecTable._sanitize_value("price", "1 000,50") == "1000.50",
          "таблица: вставка мимо валидации (var.set) проходит санитизацию")

    # Применение количества строк по Enter в поле «Строк:».
    app.spec_table._count_var.set("6")
    # клавиша приходит во внутренний tk.Entry, который должен быть в фокусе
    inner = app.spec_table._count_entry.entry._entry
    inner.focus_force()
    app.root.update()
    inner.event_generate("<Return>")
    app.root.update()
    check(len(app._doc_save_flow.spec_rows) == 6, "таблица: количество строк применяется по Enter")

    # 5.4: массовое задание строк больше не O(n²) — 500 строк создаются быстро.
    _t0 = time.perf_counter()
    app.spec_table.set_row_count(500)
    _elapsed = time.perf_counter() - _t0
    check(len(app._doc_save_flow.spec_rows) == 500, "5.4: таблица задаёт 500 строк")
    check(_elapsed < 10.0,
          f"5.4: 500 строк создаются без O(n²) за {_elapsed:.1f} с (<10 с)")
    check(str(app._doc_save_flow.spec_rows[0]["cells"]["name"].cget("readonlybackground")) != "",
          "5.4: после массового задания строк все строки стилизованы")

    # 5.4: комбобокс ограничивает число пунктов, подсветка движется точечно.
    _combo = m.ui.ComboField(app.root, values=["— выберите —"])
    _many = [f"Вариант {_i:03d}" for _i in range(300)]
    _popup = m.ui._ComboPopup(_combo, _many, _combo._pick)
    check(len(_popup._rows) == m.ui._ComboPopup.MAX_ITEMS,
          "5.4: в выпадающем списке не больше 100 пунктов")
    check(_popup._truncated, "5.4: длинный список помечен как обрезанный")
    _popup._set_index(0)
    _hover_bg = _popup._rows[0].cget("bg")
    _normal_bg = _popup._rows[7].cget("bg")
    check(_hover_bg != _normal_bg, "5.4: активная строка подсвечивается")
    _popup._set_index(5)
    check(_popup._rows[0].cget("bg") == _normal_bg
          and _popup._rows[5].cget("bg") == _hover_bg,
          "5.4: подсветка переставляется точечно, не перескрашивая список")
    _popup._confirm(5)
    check(_combo.get() == _many[5], "5.4: выбор из обрезанного списка работает")
    try:
        _popup.close()
    except Exception:
        pass
    _combo.destroy()

    # Выделение строки и удаление по Delete.
    app.spec_table.set_row_count(4)
    app.spec_table._selected_row = 1
    app.spec_table._delete_row(1)
    check(len(app._doc_save_flow.spec_rows) == 3 and app._doc_save_flow.spec_rows[0]["num"].cget("text") == "1"
          and app._doc_save_flow.spec_rows[1]["num"].cget("text") == "2",
          "таблица: Delete удаляет выделенную строку, нумерация сдвигается")
    check(str(app._doc_save_flow.spec_rows[0]["cells"]["name"].cget("state")) == "readonly",
          "таблица: ячейка доступна для правки только по двойному клику")

    # ---------- P1: подтверждения, диалоги, защита данных (этапы 2–4) ----------

    # 2.5: подтверждение удаления документа с описанием последствий.
    _orig_confirm_stub = m.ui.confirm
    _del_caps: list = []
    m.ui.confirm = lambda *args, **kwargs: _del_caps.append(args) or False
    try:
        app._lifecycle_flow.delete_document(app.store.documents[0])
    finally:
        m.ui.confirm = _orig_confirm_stub
    _del_msg = str(_del_caps[-1]) if _del_caps else ""
    check(bool(_del_caps), "2.5: удаление документа требует подтверждения")
    check("Переместить документ в корзину" in _del_msg
          and "Файл на диске останется" in _del_msg
          and "№ 5" in _del_msg and "Ромашка" in _del_msg,
          "2.5: в подтверждении тип, номер, контрагент и судьба файла")
    check(not app.store.trashed_documents(),
          "2.5: после «Отмена» документ не удаляется")

    # 2.5: подтверждение удаления контрагента с перечнем документов.
    _cp_caps: list = []
    m.ui.confirm = lambda *args, **kwargs: _cp_caps.append(args) or False
    try:
        app._lifecycle_flow.delete_counterparty(party_rec)
    finally:
        m.ui.confirm = _orig_confirm_stub
    _cp_msg = str(_cp_caps[-1]) if _cp_caps else ""
    check(bool(_cp_caps) and "Удалить контрагента" in _cp_msg,
          "2.5: удаление контрагента требует подтверждения")
    check("Договор поставки" in _cp_msg and "Файлы на диске останутся" in _cp_msg,
          "2.5: в подтверждении перечень документов и судьба файлов")
    check(not app.store.trashed_documents(),
          "2.5: после «Отмена» ничего не удалено")

    # 2.8: недоступная папка документов — явный выбор, а не молчаливая подмена.
    _choice_caps: list = []
    _orig_choice = app._choice_dialog
    app._choice_dialog = lambda *args, **kwargs: _choice_caps.append(args) or "retry"
    try:
        app._settings_flow.resolve_unavailable_documents_dir(
            os.path.join(tempfile.gettempdir(), "missing_docs_dir"),
            OSError(2, "Нет такого каталога"),
        )
    finally:
        app._choice_dialog = _orig_choice
        app._docs_dir_retry_count = 0
    check(bool(_choice_caps), "2.8: недоступная папка поднимает диалог")
    _choice_text = str(_choice_caps[0]) if _choice_caps else ""
    check("недоступна" in _choice_text.lower()
          and "Повторить" in _choice_text and "Папка по умолчанию" in _choice_text,
          "2.8: в диалоге варианты повторить / другая / по умолчанию")

    # 2.9: утерянный оригинал импортированного документа не подменяется молча.
    _lost_rec = {
        "id": 555001, "kind": storage.KIND_CONTRACT, "party_id": party_rec["id"],
        "number": "777", "date": "10.01.2026",
        "data": dict(party_rec["fields"]),
        "extra": {
            "imported": True,
            "source_file": os.path.join(tempfile.gettempdir(), "smoke_no_original.docx"),
        },
    }
    _lost_caps: list = []
    app._choice_dialog = lambda *args, **kwargs: _lost_caps.append(args) or "cancel"
    _orig_show_folder = m.show_in_folder
    m.show_in_folder = lambda path: OPENED.append(f"folder:{path}")
    try:
        app._lifecycle_flow.export_document(_lost_rec)
    finally:
        app._choice_dialog = _orig_choice
        m.show_in_folder = _orig_show_folder
    check(bool(_lost_caps) and "Оригинал не найден" in str(_lost_caps[0][0]),
          "2.9: утеря оригинала поднимает диалог")
    check("Пересобрать из шаблона" in str(_lost_caps[0])
          and "отличаться" in str(_lost_caps[0]),
          "2.9: диалог предупреждает о неточности пересборки")
    check(not OPENED,
          "2.9: без выбора пользователя суррогат не создаётся")
    OPENED.clear()

    # 3.3: повторный импорт того же файла не плодит карточки.
    _dup_copy = os.path.join(tempfile.gettempdir(), "smoke_p1_dup.docx")
    with open(_dup_copy, "wb") as _fh:
        _fh.write(b"PK\x03\x04 stub")
    _dup_result = {
        "name": "Договор ООО.docx", "kind": storage.KIND_CONTRACT,
        "number": "888", "date": "10.01.2026",
        "party": {
            "recognized": True, "type": "ooo",
            "display": "ООО «Ромашка-Импорт»",
            "fields": {"ИНН": "2540123456",
                       "Сокращенное название компании": "ООО «Ромашка-Импорт»"},
            "missing": "",
        },
        "text": "", "ok": True, "recognized": True,
        "copied": _dup_copy, "error": "", "warning": "",
    }
    _orig_report = app._import_flow.show_report
    app._import_flow.show_report = lambda entries, summary: None
    try:
        app._import_flow.finish([dict(_dup_result)])
        _dup_card = app.store.find_by_inn("2540123456")
        check(_dup_card is not None, "3.3: карточка по ИНН создана автоматически")
        _parties_after_first = len(app.store.counterparties)
        app._import_flow.finish([dict(_dup_result)])
        check(len(app.store.counterparties) == _parties_after_first,
              "3.3: повторный импорт того же файла не плодит карточек")
        check(app.store.find_by_inn("2540123456")["id"] == _dup_card["id"],
              "3.3: документ привязан к существующей карточке")
        app.store.purge_counterparty(_dup_card["id"])
    finally:
        app._import_flow.show_report = _orig_report
        try:
            os.remove(_dup_copy)
        except OSError:
            pass

    # 5.2/5.3: пакет импорта пишется на диск ОДНИМ save(), разбор отменяется.
    _batch_dir = tempfile.mkdtemp(prefix="smoke_batch_")
    _batch_results = []
    _batch_fields = {"ИНН": "2540999887", "Сокращенное название компании": "ООО «Батч»"}
    for _i in range(1, 31):
        _fpath = os.path.join(_batch_dir, f"Док {_i}.docx")
        _docx = m.Document()
        _docx.add_paragraph(f"Спецификация № {_i} от 01.01.2026 ООО «Батч» ИНН 2540999887")
        _docx.save(_fpath)
        _batch_results.append({
            "name": f"Док {_i}.docx", "kind": storage.KIND_SPEC,
            "number": str(_i), "date": "01.01.2026",
            "party": {"recognized": True, "type": "ooo", "display": "ООО «Батч»",
                      "fields": dict(_batch_fields), "missing": ""},
            "text": "", "ok": True, "recognized": True,
            "copied": _fpath, "error": "", "warning": "",
        })
    _save_calls = [0]
    _orig_store_save = app.store.save

    def _counting_save():
        _save_calls[0] += 1
        return bool(_orig_store_save())

    app.store.save = _counting_save
    _orig_report2 = app._import_flow.show_report
    app._import_flow.show_report = lambda entries, summary: None
    try:
        _docs_before = len(app.store.documents)
        app._import_flow.finish(_batch_results)
        check(_save_calls == [1],
              f"5.2: пакет из 30 файлов записан одним save() (раньше — 30 записей): {_save_calls}")
        check(len(app.store.documents) == _docs_before + 30,
              "5.2: все 30 документов пакета добавлены в реестр")
        check(app.store.find_by_inn("2540999887") is not None,
              "5.2: карточка контрагента из пакета создана")
        check("def cancel" in _app_src and "self.cancelled = True" in _app_src,
              "5.3: у импорта есть отмена «Отменить» (кнопка в окне загрузки)")
        app._import_flow.running = True
        app._import_flow.cancel()
        check(app._import_flow.cancelled, "5.3: после «Отменить» разбор остатка пакета прекращается")
        app._import_flow.running = False
        app._import_flow.cancelled = False
    finally:
        app.store.save = _orig_store_save
        app._import_flow.show_report = _orig_report2
        _batch_party_rec = app.store.find_by_inn("2540999887")
        if _batch_party_rec is not None:
            app.store.purge_counterparty(_batch_party_rec["id"])
        shutil.rmtree(_batch_dir, ignore_errors=True)

    # 4.1: дубликаты контрагентов по ИНН и номеров документов.
    # Создание карточки с чужим ИНН gated диалогом save_card (find_by_inn).
    _dup_party = app.store.add_counterparty("ooo", {
        "ИНН": "7701234567",
        "Сокращенное название компании": "ООО Ромашка",
    }, save=False)
    check(any("ООО Ромашка" in base for base in app._doc_context_flow.party_duplicate_bases()),
          "4.1: одинаковые карточки помечаются дубликатами в списках")
    app.store.counterparties.remove(_dup_party)
    # save_card с шага 6 декомпозиции живёт в counterparty_controller —
    # исходник читаем оттуда (def save_card).
    with open("counterparty_controller.py", encoding="utf-8") as _cpf:
        _cp_ctrl_src = _cpf.read()
    _cp_src_start = _cp_ctrl_src.index("def save_card")
    _cp_src = _cp_ctrl_src[_cp_src_start:_cp_src_start + 3000]
    check("find_by_inn" in _cp_src and "Открыть существующего" in _cp_src,
          "4.1: при создании карточки ИНН проверяется на дубликат")
    check(app.store.find_by_inn("7701234567")["id"] == party_rec["id"],
          "4.1: ИНН из формы найдёт существующую карточку")
    _dupnum_caps: list = []
    m.ui.confirm = lambda *args, **kwargs: _dupnum_caps.append(args) or False
    try:
        _dup_allowed = app._doc_save_flow.confirm_document_duplicate(
            party_rec["id"], storage.KIND_CONTRACT, "5", "12.01.2026",
        )
    finally:
        m.ui.confirm = _orig_confirm_stub
    check(not _dup_allowed and bool(_dupnum_caps),
          "4.1: повтор номера договора требует подтверждения")
    check("уже есть" in str(_dupnum_caps[0]) and "№ 5" in str(_dupnum_caps[0]),
          "4.1: в предупреждении назван существующий документ")
    check(app._doc_save_flow.confirm_document_duplicate(
        party_rec["id"], storage.KIND_CONTRACT, "555", "12.01.2026",
    ), "4.1: новый номер не требует подтверждения")

    # 4.2: единый «текущий контрагент» на всех страницах документов.
    app._doc_context_flow.set_current_party(party_rec)
    _expected_party = app._doc_context_flow.party_display(
        party_rec, app._doc_context_flow.party_duplicate_bases())
    check(app._current_party_id == party_rec["id"],
          "4.2: текущий контрагент запомнен")
    for _attr in ("contract_party_var", "addl_party_var", "spec_party_var"):
        check(getattr(app, _attr).get() == _expected_party,
              f"4.2: выбор перенесён на {_attr}")
    app.contract_party_var.set("Несуществующий контрагент XYZ")
    app._doc_context_flow.refresh_party_combo_values()
    check(app.contract_party_var.get() == "",
          "4.2: устаревшее имя в поле выбора сбрасывается")

    # 4.3: предпросмотр только для чтения — UI не обещает правок.
    check(str(app.contract_preview.cget("state")) == "disabled"
          and str(app.addl_preview.cget("state")) == "disabled"
          and str(app.spec_text.cget("state")) == "disabled",
          "4.3: предпросмотр только для чтения")

    # 4.4: защита несохранённых данных.
    check("contract" not in app._doc_context_flow.dirty_pages(),
          "4.4: сохранённая форма не считается изменённой")
    app.contract_number_var.set("999")
    check("contract" in app._doc_context_flow.dirty_pages(),
          "4.4: изменение формы фиксируется как несохранённое")
    _close_caps: list = []
    m.ui.confirm = lambda *args, **kwargs: _close_caps.append(args) or False
    try:
        app.close()
    finally:
        m.ui.confirm = _orig_confirm_stub
    check(bool(_close_caps)
          and "несохранённые изменения" in str(_close_caps[0]).lower(),
          "4.4: при закрытии с изменениями запрашивается подтверждение")
    check(app.root.winfo_exists(), "4.4: после «Остаться» окно живо")
    app.contract_number_var.set("5")
    check("contract" not in app._doc_context_flow.dirty_pages(),
          "4.4: возврат значений возвращает форме сохранённое состояние")

    # 4.5: быстрые UX-правки.
    check("<Control-s>" in _main_src, "4.5: Ctrl+S привязан к сохранению")
    check(hasattr(app, "toast"), "4.5: немодальное уведомление об успехе есть")
    _no_contacts = {key: value for key, value in party_rec["fields"].items()
                    if key not in (m.F.EMAIL, m.F.PHONE)}
    _contact_errors = m.validate_party_fields(_no_contacts, is_ip=False)
    check("Email" not in _contact_errors and "Телефон" not in _contact_errors,
          "4.5: email и телефон необязательны")

    # 6.1/6.5: установщик — директивы, версия, вопрос деинсталлятора.
    with open("installer.iss", "r", encoding="utf-8") as _iss_f:
        _iss_src = _iss_f.read()
    for _directive in ("AppMutex=Global\\EWDocHub_InstallerMutex;Local\\EWDocHub_InstallerMutex",
                       "MinVersion=10.0",
                       "ArchitecturesAllowed=x64compatible", "CloseApplications=yes",
                       "PrivilegesRequiredOverridesAllowed=dialog",
                       "LicenseFile=license.txt",
                       "EW.DocHub-{#MyAppVersion}-setup"):
        check(_directive in _iss_src, f"6.1: в установщике есть {_directive}")
    check("DelTree" in _iss_src and "DogovorGenerator" in _iss_src,
          "6.5: деинсталлятор спрашивает про данные и ключ DaData")
    # Аудит 2.0.1: версии не должны «разъезжаться» — installer.iss обязан
    # совпадать с APP_VERSION, spec обязан читать версию из main.py.
    _iss_version = re.search(r'#define MyAppVersion "([^"]+)"', _iss_src)
    check(_iss_version is not None and _iss_version.group(1) == m.APP_VERSION,
          "6.1: версия установщика совпадает с APP_VERSION из main.py")
    with open("EWDochub.spec", "r", encoding="utf-8") as _spec_f:
        _spec_src = _spec_f.read()
    check("VSVersionInfo" in _spec_src and "version=exe_version" in _spec_src,
          "6.1: у exe есть VersionInfo в сборке")
    check("filevers=FILEVERS" in _spec_src and "prodvers=FILEVERS" in _spec_src,
          "6.1: бинарная версия exe вычисляется из APP_VERSION, а не захардкожена")
    check(re.search(r'APP_VERSION\s*=\s*"\d+\.\d+\.\d+"', _spec_src) is None,
          "6.1: в spec нет собственной копии версии (читается из main.py)")
    check(os.path.exists("license.txt"), "6.1: файл лицензии и уведомлений существует")

    # 6.2: сообщение соответствует действительности — открытие проверяется.
    app._doc_context_flow.set_current_party(party_rec)
    m.open_file = lambda path: False
    try:
        app.contract_number_var.set("9")
        app._doc_save_flow.save_contract()
    finally:
        m.open_file = lambda path: True  # возвращаем стаб успешного открытия
    app.root.update()
    _toast_text = str(app.toast_label.cget("text"))
    check(_toast_text.rstrip(".").endswith("сохранён") and "и открыт" not in _toast_text,
          f"6.2: при неудачном открытии не обещается «открыт»: {_toast_text!r}")
    try:
        app.contract_number_var.set("10")
        app._doc_save_flow.save_contract()
    finally:
        m.open_file = _real_open_file
    app.root.update()
    check(str(app.toast_label.cget("text")).endswith("сохранён и открыт."),
          f"6.2: при успешном открытии тост говорит «сохранён и открыт»: "
          f"{app.toast_label.cget('text')!r}")

    # 6.2: папка документов по умолчанию — видимые «Документы», а не скрытый %APPDATA%.
    _def_dir = app._settings_flow.default_documents_dir(interactive=False)
    check("Documents" in _def_dir and os.path.basename(_def_dir) == "EW.DocHub",
          f"6.2: папка документов по умолчанию в «Документах»: {_def_dir}")
    check(m.open_file(os.path.join(tempfile.gettempdir(), "smoke_missing_6_2.docx")) is False,
          "6.2: open_file возвращает False при неудаче открытия (файла нет на диске)")

    # 2.15: проверка обновлений — страница «Настройки» и логика updater
    # (без сетевых вызовов: только разбор версий/ответа и виджеты).
    app._select_page("settings")
    app.root.update()
    check(hasattr(app, "_update_check_button") and hasattr(app, "_update_status_var"),
          "2.15: на «Настройках» есть кнопка проверки обновлений и статус версии")
    check(str(m.APP_VERSION) in app._update_status_var.get(),
          f"2.15: статус показывает установленную версию {m.APP_VERSION}")
    check(app._update_auto_var.get() is True,
          "2.15: авто-проверка обновлений включена по умолчанию")
    check(app._update_flow.busy is False and app._update_flow.result_queue is None,
          "2.15: глушилка EWDOCHUB_NO_UPDATE_CHECK отключила автопроверку при старте")
    check(updater.parse_version("v2.16.0") == (2, 16, 0)
          and updater.is_newer("2.15", "2.15.0") is False
          and updater.is_newer("2.16.0", "2.15.0") is True,
          "2.15: сравнение версий — v-префикс и выравнивание длины")
    check(updater.parse_release_payload({"tag_name": "v1.0", "draft": True}) is None
          and updater.parse_release_payload({"tag_name": ""}) is None,
          "2.15: draft-релиз и ответ без тега не считаются релизом")
    _upd_rel = updater.parse_release_payload({
        "tag_name": "v2.16.0",
        "html_url": "https://example.com/release/v2.16.0",
        "assets": [{"name": "EW.DocHub-2.16.0-setup.exe",
                    "browser_download_url": "https://example.com/EW.DocHub-2.16.0-setup.exe"}],
    })
    check(_upd_rel is not None and _upd_rel["version_str"] == "2.16.0"
          and _upd_rel["installer_url"].endswith("-setup.exe"),
          "2.15: из ответа GitHub извлекаются версия и ссылка на установщик")
    _upd_settings: dict = {}
    updater.mark_checked(_upd_settings)
    check(not updater.auto_check_due(_upd_settings),
          "2.15: суточный троттлинг — сразу после проверки автопроверка не нужна")
    check(updater.check_enabled() is False,
          "2.15: check_enabled() видит env-глушку smoke-прогона")

    # Реестр.
    app._select_page("registry")
    app.root.update()
    check(app._registry_rendered is True,
          "п.9: реестр отрисован при первом открытии страницы")
    check(app.current_page == "registry" and not app.save_button.winfo_manager(),
          "кнопка сохранения скрыта на реестре")
    texts = []

    def collect(widget):
        for child in widget.winfo_children():
            try:
                value = child.cget("text")
            except Exception:
                value = None
            if isinstance(value, str):
                texts.append(value)
            collect(child)

    collect(app.registry_list)
    registry_text = "\n".join(texts)
    check("ООО Ромашка" in registry_text, "реестр: контрагент виден")

    # Документы в реестре появляются после выбора контрагента.
    app._registry_flow.select_party(party_rec)
    app.root.update()
    texts = []
    collect(app.registry_list)
    registry_text = "\n".join(texts)
    check("Договор поставки № 5" in registry_text, "реестр: договор виден")
    check("Доп. соглашение" in registry_text, "реестр: доп. соглашение видно")

    # Поиск по ИНН в реестре.
    app._registry_flow.select_party(party_rec)  # снять выбор
    app.registry_search_var.set("7701234567")
    app._registry_flow.apply_filter()
    app.root.update()
    texts = []
    collect(app.registry_list)
    check("ООО Ромашка" in "\n".join(texts), "реестр: поиск по ИНН находит контрагента")
    app.registry_search_var.set("")
    app._registry_flow.apply_filter()

    # 5.1: клик по контрагенту не пересоздаёт строки других контрагентов.
    _party_2 = app.store.add_counterparty("ooo", {
        "ИНН": "7709990000",
        "Сокращенное название компании": "ООО Стрелка",
        "Юридический адрес": "г. Владивосток, ул. Светланская, д. 1",
    }, save=False)
    app._registry_flow.refresh()
    _frames_before = {pid: entry["frame"] for pid, entry in app._registry_rows.items()}
    check(len(_frames_before) == 2, "5.1: в реестре два контрагента для проверки")
    app._registry_flow.select_party(party_rec)
    check(all(app._registry_rows[pid]["frame"] is _frames_before[pid]
              for pid in _frames_before if pid in app._registry_rows),
          "5.1: выбор контрагента не пересоздаёт строки реестра")
    check(app._registry_selected_id == party_rec["id"], "5.1: выбор контрагента запомнен")
    _docs_texts = []
    for _child in app._registry_docs_frame.winfo_children():
        try:
            _docs_texts.append(_child.cget("text"))
        except Exception:
            pass
    check(any("Документы" in str(_t) for _t in _docs_texts),
          "5.1: секция документов выбранного видна после частичного обновления")
    app._registry_flow.select_party(party_rec)  # снять выбор
    check(app._registry_selected_id is None, "5.1: повторный клик снимает выбор")
    app.store.counterparties.remove(_party_2)
    app._registry_flow.refresh()

    check(app.store.delete_document(app.store.documents[0]["id"]),
          "хранилище: окончательное удаление записи документа")

    # Удаление контрагента: теперь это перемещение в корзину (п. 2.6).
    app._lifecycle_flow.delete_counterparty(party_rec)
    app.root.update()
    check(not app.store.active_counterparties() and not app.store.active_documents(),
          "корзина: контрагент и документы перемещены в корзину")
    check(bool(app.store.trashed_counterparties()) and bool(app.store.trashed_documents()),
          "корзина: записи доступны для восстановления")

    app._lifecycle_flow.restore_counterparty(party_rec)
    app.root.update()
    check(app.store.get_counterparty(party_rec["id"]) is not None,
          "корзина: контрагент восстановлен")
    check(bool(app.store.documents_for(party_rec["id"])),
          "корзина: документы контрагента восстановлены")

    # Повторное удаление и полная очистка корзины.
    _party_again = app.store.get_counterparty(party_rec["id"])
    app._lifecycle_flow.delete_counterparty(_party_again)
    app.root.update()
    check(app.store.trash_count() > 0, "корзина: записи снова в корзине")
    app.store.purge_all_trash()
    check(app.store.trash_count() == 0, "корзина: полная очистка удаляет записи")

    app.store.save()
finally:
    try:
        m.open_file = _real_open_file
    except Exception:
        pass
    try:
        app.root.destroy()
    except Exception:
        pass

print("SMOKE OK" if not FAILURES else "SMOKE FAILED")

# Аудит 2.0.1: сводка всех неудач в конце — в CI видно всё сразу.
if FAILURES:
    print(f"\n=== СВОДКА: {len(FAILURES)} из {_TOTAL_CHECKS} проверок не прошли ===")
    for _message in FAILURES:
        print(f"FAIL: {_message}")
    sys.exit(min(len(FAILURES), 100))
print(f"SMOKE OK: {_TOTAL_CHECKS} проверок")
