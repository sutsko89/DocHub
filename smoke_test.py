"""Быстрая проверка интерфейса и логики без ручного запуска."""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile

import dadata_client as dadata

# Настройки приложения и база данных (контрагенты/документы) читаются из APPDATA:
# изолируем тест от реальных данных пользователя, чтобы результат был детерминированным.
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "dogovor_smoke_appdata")

import main as m
import reconcile
import storage
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
m.open_file = lambda path: OPENED.append(path)


def check(condition, message):
    if not condition:
        print(f"FAIL: {message}")
        sys.exit(1)
    print(f"ok: {message}")


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

errors = m.validate_party_fields({}, is_ip=False)
check("ИНН" in errors and "Расчётный счёт" in errors, "валидация контрагента: пустые поля")
check(m.validate_party_fields({
    "ИНН": "7701234567", "КПП": "770101001", "ОГРН": "1037739010221",
    "Юридический адрес": "Москва", "Email": "a@b.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012", "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "Банк", "БИК": "044525225",
    "Полное название компании (Покупь)": "ООО", "Сокращенное название компании": "ООО",
    "ФИО директора": "Иванов И.И.",
}, is_ip=False) == {}, "валидация контрагента: валидные данные")

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
check(m.ContractApp._is_inside_directory(_path_base, _path_inside),
      "пути: вложенная папка разрешена")
check(not m.ContractApp._is_inside_directory(_path_base, _path_outside),
      "пути: выход за папку документов запрещён")

# ---------- Хранилище ----------
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

# ---------- Надёжность записи реестра (п. 2.1) ----------
check(store.save(), "хранилище: повторная запись реестра")
check(os.path.exists(storage.db_path() + ".bak"), "хранилище: создана резервная копия .bak")
with open(storage.db_path(), "r", encoding="utf-8") as _fh:
    _raw_db = json.load(_fh)
check(_raw_db.get("schema_version") == storage.SCHEMA_VERSION,
      "хранилище: schema_version записан в файл")
check(storage.is_deleted({"deleted_at": "2026-01-01T00:00:00"}),
      "корзина: запись с deleted_at считается удалённой")
check(not storage.is_deleted({}), "корзина: запись без deleted_at активна")

# ---------- Защита чтения реестра (п. 1.4) ----------
_appdata_main = os.environ["APPDATA"]
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "smoke_storage_bad")
# Песочница могла остаться от прошлого прогона (в т.ч. .bad/.tmp файлы) — сбрасываем.
shutil.rmtree(storage.data_dir(), ignore_errors=True)
os.makedirs(storage.data_dir(), exist_ok=True)
with open(storage.db_path(), "w", encoding="utf-8") as _fh:
    _fh.write("{ это не JSON !!!")
_broken = storage.DataStore()
check(bool(_broken.load_warning), "реестр: при битом JSON пользователь предупреждён")
check(not _broken.counterparties and not _broken.documents,
      "реестр: битый файл — реестр пуст, запуск не падает")
_bad_copies = [n for n in os.listdir(storage.data_dir())
               if n.startswith("database.json.bad-")]
check(len(_bad_copies) == 1,
      f"реестр: повреждённый файл сохранён как .bad ({_bad_copies})")

# Восстановление из временного файла записи.
with open(storage.db_path(), "w", encoding="utf-8") as _fh:
    _fh.write("{ тоже не JSON")
_rescue_tmp = storage.db_path() + ".999.tmp"
with open(_rescue_tmp, "w", encoding="utf-8") as _fh:
    _fh.write('{"counterparties": [{"id": 3, "type": "ooo", "fields": {"ИНН": "7701234567"}}], '
              '"documents": [], "seq": 4}')
_rescued = storage.DataStore()
check([rec["id"] for rec in _rescued.counterparties] == [3],
      "реестр: данные восстановлены из временного файла записи")
check("восстановлены" in _rescued.load_warning,
      "реестр: о восстановлении сообщено пользователю")

# null вместо списков не роняет запуск (п. 1.4).
with open(storage.db_path(), "w", encoding="utf-8") as _fh:
    _fh.write('{"counterparties": null, "documents": null}')
_null_store = storage.DataStore()
check(bool(_null_store.load_warning) and not _null_store.counterparties,
      "реестр: null вместо разделов — предупреждение без падения")

# ---------- Сбой записи реестра (п. 1.5) ----------
os.remove(storage.db_path())
_ro = storage.DataStore()
_ro.add_counterparty("ooo", {"ИНН": "7709999999"}, save=False)
os.makedirs(storage.db_path())  # каталог вместо файла: os.replace падает
check(_ro.save() is False and _ro.last_error,
      "save(): сбой записи возвращает False и last_error, а не ложный успех")
_upd_id = _ro.counterparties[0]["id"]
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
for key, value in (
    (di.K_INN, m.F.INN), (di.K_KPP, m.F.KPP), (di.K_OGRN, m.F.OGRN),
    (di.K_OGRNIP, m.F.OGRNIP), (di.K_FULL_NAME, m.F.FULL_NAME),
    (di.K_SHORT_NAME, m.F.SHORT_NAME), (di.K_DIRECTOR, m.F.DIRECTOR),
    (di.K_IP_NAME, m.F.IP_NAME), (di.K_LEGAL_ADDRESS, m.F.LEGAL_ADDRESS),
    (di.K_EMAIL, m.F.EMAIL), (di.K_PHONE, m.F.PHONE),
    (di.K_ACCOUNT, m.F.ACCOUNT), (di.K_CORR_ACCOUNT, m.F.CORR_ACCOUNT),
    (di.K_BANK, m.F.BANK), (di.K_BIK, m.F.BIK),
):
    check(key == value, f"импорт: ключ «{key}» совпадает с F")

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
check(_party["fields"][di.K_ACCOUNT] == "40702810123456789012",
      f"счёт: «Р/счет №» извлечён ({_party['fields'][di.K_ACCOUNT]!r})")
check(_party["fields"][di.K_CORR_ACCOUNT] == "30101810100000000225",
      "счёт: корр. счёт извлечён")
check(_party["fields"][di.K_ACCOUNT] != _party["fields"][di.K_CORR_ACCOUNT],
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
    check(_variant["fields"][di.K_ACCOUNT] == "40702810123456789012",
          f"счёт: вариант «{_label}» извлечён")

_cell_account = di.extract_party(
    "ООО «Тест», ИНН 7701234567\nБИК 044525225\n40702810999999999991"
)
check(_cell_account["fields"][di.K_ACCOUNT] == "40702810999999999991",
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

check(store.find_by_name("ООО «Ромашка»") is not None,
      "хранилище: поиск контрагента по названию без учёта кавычек")
check(store.find_by_name("нет такого") is None,
      "хранилище: поиск по названию не находит лишнего")

# ---------- Интерфейс ----------
app = m.ContractApp()
try:
    app.root.update_idletasks()

    check(hasattr(app, "nav_holder") and len(app.nav_items) == 6,
          "кнопки разделов в шапке, включая «Настройки»")
    check(set(app.pages) == {"contract", "additional", "spec", "registry", "counterparty", "settings"},
          f"страницы: {sorted(app.pages)}")
    check(hasattr(app, "contract_preview") and hasattr(app, "addl_preview")
          and hasattr(app, "spec_text"),
          "панели предпросмотра: договор и доп. соглашение — есть (спецификация — есть)")
    check(hasattr(app, "status_label"), "строка статуса в футере")
    check(app.save_button.cget("text") == "Сохранить договор", "кнопка футера: договор")

    # Контрагент появляется в списке страницы договора.
    app._refresh_party_combo_values()
    app.root.update()
    app.contract_party_var.set(storage.display_with_inn(party_rec))
    app.root.update()
    check("ООО Ромашка" in app.contract_party_var.get(), "выбор контрагента в списке")

    app.contract_number_var.set("5")
    app.contract_day_var.set("12")
    app.contract_month_var.set("января")
    app.root.update()
    app._render_contract_preview()
    check("Ромашка" in app.contract_preview.get("1.0", "end"),
          "предпросмотр договора: данные контрагента видны")
    data = dict(party_rec["fields"])
    data.update({m.F.NUMBER: "5", m.F.DAY: "12", m.F.MONTH: "января"})
    doc = app._render_contract_docx(data, is_ip=False)
    contract_text = "\n".join(p.text for p in doc.paragraphs)
    check("ООО Ромашка" in contract_text and "Иванов" in contract_text,
          "печатная форма договора: данные контрагента подставлены")
    check("Ромашка" in contract_text and "5" in contract_text,
          "печатная форма договора: номер подставлен")
    check(len(doc.tables) == 1, "печатная форма договора: таблица реквизитов на месте")

    app.save_contract()
    check(len(app.store.documents) == 1 and app.store.documents[0]["kind"] == storage.KIND_CONTRACT,
          "договор сохранён в реестре")

    _saved_doc = app.store.documents[0]
    check(app._import_document_is_duplicate(party_rec["id"], {
        "kind": _saved_doc["kind"],
        "number": _saved_doc["number"],
        "date": _saved_doc["date"],
    }), "импорт: повторный документ распознан как дубль")
    check(not app._import_document_is_duplicate(party_rec["id"], {
        "kind": _saved_doc["kind"],
        "number": "999",
        "date": _saved_doc["date"],
    }), "импорт: другой номер — не дубль")
    check(not app._import_document_is_duplicate(party_rec["id"], {
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
    doc2 = app._render_additional_docx(data, is_ip=False, number="1",
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
    app._render_additional_preview()
    check("Ромашка" in app.addl_preview.get("1.0", "end"),
          "предпросмотр доп. соглашения: данные контрагента видны")
    app.save_additional()
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
    row = app._spec_rows[0]
    row["name"].set("Товар")
    row["qty"].set("10")
    row["unit"].set("шт.")
    row["price"].set("100")
    app.root.update()
    spec_text = app.spec_text.get("1.0", "end")
    check("СПЕЦИФИКАЦИЯ № 1" in spec_text, "заголовок спецификации")
    check("1 000" in spec_text, "итог спецификации посчитан")
    check("одна тысяча" in spec_text.lower(), "итог прописью")
    check("шт." in spec_text, "ед. изм. видна в предпросмотре")
    app.save_spec()
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
    _ip_doc = app._render_spec_docx(_ip_spec_rec)
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
    _ooo_text = _doc_full_text(app._render_spec_docx(_ooo_spec_rec))
    check("{{" not in _ooo_text, "спецификация ООО: нет незаполненных {{...}}")
    check("770101001" in _ooo_text and "1037739010221" in _ooo_text,
          "спецификация ООО: КПП и ОГРН подставлены как раньше")

    # ИП-карточка нужна была только для проверки печатной формы — убираем её,
    # чтобы не мешать проверкам корзины ниже.
    app.store.counterparties.remove(_ip_party)

    # Порядок «файл → запись» (п. 1.6): сбой реестра откатывает и запись, и файл.
    _rb_path = os.path.join(tempfile.gettempdir(), "smoke_rollback.docx")
    with open(_rb_path, "wb") as _fh:
        _fh.write(b"stub")
    _rb_rec = {"id": 987654321, "kind": storage.KIND_CONTRACT, "party_id": party_rec["id"]}
    app.store.documents.append(_rb_rec)
    app._rollback_document(_rb_rec, _rb_path)
    check(_rb_rec not in app.store.documents and not os.path.exists(_rb_path),
          "файл→запись: при сбое реестра документ откатывается целиком")

    # Регистрация импорта: документ без ИНН привязывается к подтверждённому
    # контрагенту, а не создаёт новую карточку (п. 3.3).
    _tmp_copied = os.path.join(tempfile.gettempdir(), "smoke_import_copy.docx")
    with open(_tmp_copied, "wb") as _fh:
        _fh.write(b"stub")
    _orig_report = app._show_import_report
    _orig_ask = app._ask_import_party
    app._show_import_report = lambda entries, summary: None
    app._ask_import_party = lambda result: (party_rec, "attach")
    try:
        _parties_before = len(app.store.counterparties)
        app._finish_documents_import([{
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
        app._show_import_report = _orig_report
        app._ask_import_party = _orig_ask
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
    items = app._spec_items()
    check(len(items) == 2 and items[0]["name"] == "Болт М8" and items[1]["qty"] == 12,
          "таблица: строки из буфера попали в позиции")
    app.root.update()
    check("1 600" in app.spec_text.get("1.0", "end"), "таблица: итог пересчитан после вставки")
    check(len(app._spec_rows) == 10, "таблица: стартовое количество строк — 10")

    # Задание точного количества строк.
    app.spec_table.set_row_count(20)
    app.root.update()
    check(len(app._spec_rows) == 20, "таблица: количество строк задано полем «Строк»")
    app.spec_table.set_row_count(4)
    check(len(app._spec_rows) == 4, "таблица: лишние строки удаляются")
    check(app._spec_rows[-1]["num"].cget("text") == "4", "таблица: нумерация строк актуальна")

    # Применение количества строк по Enter в поле «Строк:».
    app.spec_table._count_var.set("6")
    # клавиша приходит во внутренний tk.Entry, который должен быть в фокусе
    inner = app.spec_table._count_entry.entry._entry
    inner.focus_force()
    app.root.update()
    inner.event_generate("<Return>")
    app.root.update()
    check(len(app._spec_rows) == 6, "таблица: количество строк применяется по Enter")

    # Выделение строки и удаление по Delete.
    app.spec_table.set_row_count(4)
    app.spec_table._selected_row = 1
    app.spec_table._delete_row(1)
    check(len(app._spec_rows) == 3 and app._spec_rows[0]["num"].cget("text") == "1"
          and app._spec_rows[1]["num"].cget("text") == "2",
          "таблица: Delete удаляет выделенную строку, нумерация сдвигается")
    check(str(app._spec_rows[0]["cells"]["name"].cget("state")) == "readonly",
          "таблица: ячейка доступна для правки только по двойному клику")

    # Реестр.
    app._select_page("registry")
    app.root.update()
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
    app._select_registry_party(party_rec)
    app.root.update()
    texts = []
    collect(app.registry_list)
    registry_text = "\n".join(texts)
    check("Договор поставки № 5" in registry_text, "реестр: договор виден")
    check("Доп. соглашение" in registry_text, "реестр: доп. соглашение видно")

    # Поиск по ИНН в реестре.
    app._select_registry_party(party_rec)  # снять выбор
    app.registry_search_var.set("7701234567")
    app._registry_apply_filter()
    app.root.update()
    texts = []
    collect(app.registry_list)
    check("ООО Ромашка" in "\n".join(texts), "реестр: поиск по ИНН находит контрагента")
    app.registry_search_var.set("")
    app._registry_apply_filter()

    check(app.store.delete_document(app.store.documents[0]["id"]),
          "хранилище: окончательное удаление записи документа")

    # Удаление контрагента: теперь это перемещение в корзину (п. 2.6).
    app._delete_counterparty(party_rec)
    app.root.update()
    check(not app.store.active_counterparties() and not app.store.active_documents(),
          "корзина: контрагент и документы перемещены в корзину")
    check(bool(app.store.trashed_counterparties()) and bool(app.store.trashed_documents()),
          "корзина: записи доступны для восстановления")

    app._restore_counterparty(party_rec)
    app.root.update()
    check(app.store.get_counterparty(party_rec["id"]) is not None,
          "корзина: контрагент восстановлен")
    check(bool(app.store.documents_for(party_rec["id"])),
          "корзина: документы контрагента восстановлены")

    # Повторное удаление и полная очистка корзины.
    _party_again = app.store.get_counterparty(party_rec["id"])
    app._delete_counterparty(_party_again)
    app.root.update()
    check(app.store.trash_count() > 0, "корзина: записи снова в корзине")
    app.store.purge_all_trash()
    check(app.store.trash_count() == 0, "корзина: полная очистка удаляет записи")

    app.store.save()
finally:
    try:
        app.root.destroy()
    except Exception:
        pass

print("SMOKE OK")
