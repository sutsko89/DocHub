"""Генератор договора поставки — интерфейс в стиле macOS.

Приложение построено вокруг реестра: контрагенты сохраняются в программе,
а договоры, доп. соглашения и спецификации создаются к ним и тоже остаются
в реестре. Все разделы доступны кнопками в верхней панели окна.

Логика формирования документов не менялась: те же шаблоны .docx, те же
плейсхолдеры и те же проверки полей. При сохранении договор и доп. соглашение
попадают в реестр и сразу открываются в Word — печатная форма собирается
напрямую из шаблона .docx, поэтому форматирование не «плывёт».
"""

from __future__ import annotations

import calendar
import copy
import ctypes
import hashlib
import json
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import threading
import zipfile
from datetime import datetime

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import customtkinter as ctk
import tkinter as tk
from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker
from docx import Document
from tkinter import filedialog, font as tkfont

import applog
import dadata_client as dadata
import document_importer as importer
import reconcile
import single_instance
import storage
import ui_macos as ui

# Печатные формы вынесены в document_builder (п. 9.1 плана); имена
# реэкспортируются, чтобы main.py, smoke_test.py и тесты работали как раньше.
from document_builder import (
    F, PH, SPEC_APPENDIX_NUMBER, SPEC_VAT_RATE, MONTHS_GENITIVE,
    spec_vat_amount, spec_vat_amount_from_items, digits_only,
    format_director_name, format_name_ip2, format_date, parse_decimal,
    format_amount, format_quantity, format_phone, format_number_grouped,
    number_to_words_ru, amount_to_words_ru, amount_in_words_ru,
    prices_valid_until_default, replace_in_paragraph, replace_in_doc,
    spec_items_table, fill_spec_items, mark_spec_final_page,
    docx_to_text, docx_from_text, build_replacements,
)


APP_VERSION = "1.5"
# Оставляем запас под полный путь Windows: лимит 260 символов, минус имя файла
# и расширение, минус служебные суффиксы « (2).docx».
MAX_PATH_LENGTH = 240
MAX_FILENAME_LENGTH = 80

# Масштабирование берём на себя: иначе customtkinter масштабирует повторно.
ScalingTracker.deactivate_automatic_dpi_awareness = True

if sys.platform == "win32":
    try:  # чёткий текст на экранах с масштабированием (125 %, 150 % и т. д.)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ---------- Кэш шаблонов печатных форм (п. 5.5 этапа 5) ----------
# Шаблон .docx разбирается один раз; на каждый рендер отдаётся deepcopy.
# Раньше предпросмотр при каждом вводе заново читал и разбирал ZIP шаблона.
# Кэш инвалидируется по mtime — правка шаблона подхватывается без перезапуска.
_TEMPLATE_CACHE: dict = {}


def load_template_cached(template_path: str):
    """Возвращает копию разобранного шаблона .docx (кэш по пути + mtime)."""
    try:
        mtime = os.path.getmtime(template_path)
    except OSError:
        mtime = 0.0
    cached = _TEMPLATE_CACHE.get(template_path)
    if cached is None or cached[0] != mtime:
        _TEMPLATE_CACHE[template_path] = (mtime, Document(template_path))
    return copy.deepcopy(_TEMPLATE_CACHE[template_path][1])


# ---------- Сохранение настроек между запусками ----------
SETTINGS_DIR_NAME = "DogovorGenerator"
MIN_WINDOW_WIDTH = 1120  # логические пиксели
MIN_WINDOW_HEIGHT = 620


def _user_documents_dir() -> str:
    """Папка «Документы» пользователя (п. 6.2): видимое место вместо скрытого %APPDATA%."""
    if sys.platform == "win32":
        try:
            import uuid as _uuid

            class _GUID(ctypes.Structure):
                _fields_ = [
                    ("Data1", ctypes.c_ulong),
                    ("Data2", ctypes.c_ushort),
                    ("Data3", ctypes.c_ushort),
                    ("Data4", ctypes.c_ubyte * 8),
                ]

            guid = _GUID()
            known = _uuid.UUID("FDD39AD0-238F-46AF-ADB4-6C85480369C7")  # FOLDERID_Documents
            guid.Data1, guid.Data2, guid.Data3 = (
                known.time_low, known.time_mid, known.time_hi_version,
            )
            guid.Data4[:] = known.bytes[8:]
            out = ctypes.c_void_p()
            hr = ctypes.windll.shell32.SHGetKnownFolderPath(
                ctypes.byref(guid), 0, None, ctypes.byref(out)
            )
            if hr == 0 and out.value:
                documents = ctypes.wstring_at(out.value)
                ctypes.windll.ole32.CoTaskMemFree(out)
                return documents
            if out.value:
                ctypes.windll.ole32.CoTaskMemFree(out)
        except Exception:
            pass
        profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        return os.path.join(profile, "Documents")
    return os.path.join(os.path.expanduser("~"), "Documents")


def settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, SETTINGS_DIR_NAME, "settings.json")


def load_settings() -> dict:
    try:
        with open(settings_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def write_settings(settings: dict) -> bool:
    """Сохраняет настройки атомарно (tmp + os.replace).

    Прямая перезапись при аварийном завершении оставляла обрезанный файл,
    из-за чего тема, геометрия окна и папка документов молча сбрасывались.
    """
    path = settings_path()
    tmp = path + ".tmp"
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(settings, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return True
    except (OSError, TypeError, ValueError) as exc:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        print(f"Не удалось сохранить настройки ({path}): {exc}", file=sys.stderr)
        return False


# ---------- Функции валидации ----------
def validate_ogrn(ogrn):
    """Проверка ОГРН: только цифры, 13 символов"""
    if not ogrn:
        return False
    return ogrn.isdigit() and len(ogrn) == 13


def validate_ogrnip(ogrnip):
    """Проверка ОГРНИП: только цифры, 15 символов"""
    if not ogrnip:
        return False
    return ogrnip.isdigit() and len(ogrnip) == 15


def validate_inn(inn, is_ip=False):
    """Проверка ИНН: для ООО - 10 цифр, для ИП - 12 цифр"""
    if not inn:
        return False
    if is_ip:
        return inn.isdigit() and len(inn) == 12
    return inn.isdigit() and len(inn) == 10


def validate_kpp(kpp):
    """Проверка КПП: только цифры, 9 символов"""
    if not kpp:
        return False
    return kpp.isdigit() and len(kpp) == 9


def validate_email(email):
    """Проверка email по регулярному выражению"""
    if not email:
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


# Регулярка одна на оба места: проверку сохранённых данных и фильтр ввода на лету.
_PHONE_ALLOWED_RE = re.compile(r"^[\d\+\-\(\)\s\/\.]*$")


def validate_phone(phone):
    """Проверка телефона: только цифры и знаки + - ( ) / . и пробелы."""
    return bool(phone) and _PHONE_ALLOWED_RE.match(phone) is not None


def validate_account(account):
    """Проверка расчетного счета: 20 цифр"""
    if not account:
        return False
    return account.isdigit() and len(account) == 20


def validate_correspondent_account(account):
    """Проверка корреспондентского счета: 20 цифр"""
    if not account:
        return False
    return account.isdigit() and len(account) == 20


def validate_bik(bik):
    """Проверка БИК: 9 цифр"""
    if not bik:
        return False
    return bik.isdigit() and len(bik) == 9


# ---------- Ключи полей и плейсхолдеров (единый источник имён) ----------


# Email и телефон — контактные, а не юридически обязательные поля (п. 4.5).
REQUIRED_COMMON = (
    F.NUMBER, F.DAY, F.MONTH, F.INN,
    F.LEGAL_ADDRESS,
    F.ACCOUNT, F.CORR_ACCOUNT, F.BANK, F.BIK,
)
REQUIRED_OOO = (
    F.FULL_NAME, F.SHORT_NAME,
    F.DIRECTOR, F.KPP, F.OGRN,
)
REQUIRED_IP = (F.IP_NAME, F.OGRNIP)

# Константы печатной формы спецификации (spec_template.docx).

# Импорт принимает только документы, которые программа умеет разбирать.
# Всё остальное (в т. ч. .exe/.bat/.lnk) не копируется и не открывается.
IMPORT_ALLOWED_EXTENSIONS = {".docx", ".xlsx"}

# Лимиты пакетного импорта (п. 3.4): файл, количество, суммарный объём,
# свободное место на диске. Проверяются до запуска фонового разбора.
IMPORT_MAX_FILE_SIZE = importer.MAX_FILE_SIZE
IMPORT_MAX_FILES = importer.MAX_FILES_PER_BATCH
IMPORT_MAX_TOTAL_SIZE = importer.MAX_TOTAL_SIZE
IMPORT_MIN_FREE_SPACE = importer.MIN_FREE_SPACE


# Поля карточки контрагента (без реквизитов договора).
# Email и телефон необязательны: без них карточку можно сохранить (п. 4.5).
REQUIRED_PARTY_COMMON = (
    F.INN, F.LEGAL_ADDRESS,
    F.ACCOUNT, F.CORR_ACCOUNT, F.BANK, F.BIK,
)


def validate_contract_fields(data: dict, is_ip: bool = False) -> dict:
    """Возвращает словарь «поле -> текст ошибки» для inline-подсветки в интерфейсе."""
    errors: dict[str, str] = {}

    for key in REQUIRED_COMMON + (REQUIRED_IP if is_ip else REQUIRED_OOO):
        if not str(data.get(key, "")).strip():
            errors[key] = "Заполните поле"

    ogrn_key = F.OGRNIP if is_ip else F.OGRN
    ogrn = str(data.get(ogrn_key, "")).strip()
    if ogrn and not (validate_ogrnip(ogrn) if is_ip else validate_ogrn(ogrn)):
        errors.setdefault(ogrn_key, f"{ogrn_key} должен содержать {15 if is_ip else 13} цифр")

    inn = str(data.get(F.INN, "")).strip()
    expected = 12 if is_ip else 10
    if inn and not validate_inn(inn, is_ip=is_ip):
        errors.setdefault(F.INN, f"ИНН должен содержать {expected} цифр")

    if not is_ip:
        kpp = str(data.get(F.KPP, "")).strip()
        if kpp and not validate_kpp(kpp):
            errors.setdefault(F.KPP, "КПП должен содержать 9 цифр")

    email = str(data.get(F.EMAIL, "")).strip()
    if email and not validate_email(email):
        errors.setdefault(F.EMAIL, "Email имеет неверный формат")

    phone = str(data.get(F.PHONE, "")).strip()
    if phone and not validate_phone(phone):
        errors.setdefault(F.PHONE, "Телефон может содержать только цифры и знаки: + - ( ) / . и пробелы")

    for key, checker, message in (
        (F.ACCOUNT, validate_account, "Расчётный счёт должен содержать 20 цифр"),
        (F.CORR_ACCOUNT, validate_correspondent_account, "Корреспондентский счёт должен содержать 20 цифр"),
        (F.BIK, validate_bik, "БИК должен содержать 9 цифр"),
    ):
        value = str(data.get(key, "")).strip()
        if value and not checker(value):
            errors.setdefault(key, message)

    return errors


def validate_party_fields(data: dict, is_ip: bool = False) -> dict:
    """Проверка карточки контрагента (без номера и даты договора)."""
    errors: dict[str, str] = {}

    for key in REQUIRED_PARTY_COMMON + (REQUIRED_IP if is_ip else REQUIRED_OOO):
        if not str(data.get(key, "")).strip():
            errors[key] = "Заполните поле"

    ogrn_key = F.OGRNIP if is_ip else F.OGRN
    ogrn = str(data.get(ogrn_key, "")).strip()
    if ogrn and not (validate_ogrnip(ogrn) if is_ip else validate_ogrn(ogrn)):
        errors.setdefault(ogrn_key, f"{ogrn_key} должен содержать {15 if is_ip else 13} цифр")

    inn = str(data.get(F.INN, "")).strip()
    expected = 12 if is_ip else 10
    if inn and not validate_inn(inn, is_ip=is_ip):
        errors.setdefault(F.INN, f"ИНН должен содержать {expected} цифр")

    if not is_ip:
        kpp = str(data.get(F.KPP, "")).strip()
        if kpp and not validate_kpp(kpp):
            errors.setdefault(F.KPP, "КПП должен содержать 9 цифр")

    email = str(data.get(F.EMAIL, "")).strip()
    if email and not validate_email(email):
        errors.setdefault(F.EMAIL, "Email имеет неверный формат")

    phone = str(data.get(F.PHONE, "")).strip()
    if phone and not validate_phone(phone):
        errors.setdefault(F.PHONE, "Телефон может содержать только цифры и знаки: + - ( ) / . и пробелы")

    for key, checker, message in (
        (F.ACCOUNT, validate_account, "Расчётный счёт должен содержать 20 цифр"),
        (F.CORR_ACCOUNT, validate_correspondent_account, "Корреспондентский счёт должен содержать 20 цифр"),
        (F.BIK, validate_bik, "БИК должен содержать 9 цифр"),
    ):
        value = str(data.get(key, "")).strip()
        if value and not checker(value):
            errors.setdefault(key, message)

    return errors


def validate_all_fields(data, is_ip=False):
    """Совместимость со старыми вызовами: список текстов ошибок по всем полям."""
    return list(validate_contract_fields(data, is_ip).values())


# ---------- Функции преобразования ФИО ----------


# ---------- Функция форматирования даты ----------


# ---------- Месяцы и маски ввода ----------


def restore_caret_after_reformat(entry, caret, old_len, new_len):
    """Ставит курсор с поправкой на изменение длины строки при переформатировании.

    При var.set() Tk сбрасывает курсор в непредсказуемую позицию — без этой
    поправки ввод телефона «рассыпается» уже со второй цифры. Установка через
    after_idle: внутренний callback Entry обновляет текст и курсор ПОСЛЕ
    нашего trace-обработчика, поэтому немедленный icursor затирается.
    """
    if entry is None or caret is None:
        return

    def _apply():
        try:
            entry.icursor(min(max(caret + new_len - old_len, 0), new_len))
        except Exception:
            pass

    try:
        entry.after_idle(_apply)
    except Exception:
        _apply()


def make_phone_formatter(var: "tk.StringVar", entry=None):
    """Trace-функция: переформатирует ввод телефона в маску +7 (___) ___-__-__."""

    def _format(*_args):
        raw = var.get()
        digits = digits_only(raw)
        if digits.startswith("8"):
            digits = "7" + digits[1:]
        if digits and not digits.startswith("7"):
            digits = "7" + digits
        formatted = format_phone(digits)
        if raw == formatted:
            return
        caret = None
        if entry is not None:
            try:
                caret = entry.index(tk.INSERT)
            except Exception:
                caret = None
        var.set(formatted)
        restore_caret_after_reformat(entry, caret, len(raw), len(formatted))

    return _format


# ---------- Числа прописью (RU) ----------


# ---------- Открытие файла в зависимости от ОС ----------
def open_file(file_path) -> bool:
    """Открывает файл системным приложением и возвращает успех (п. 6.2).

    Раньше результат не возвращался: тост «сохранён и открыт» показывался
    даже когда открытие не удалось. Теперь вызывающий код честно сообщает
    «сохранён» без ложного «открыт». Отсутствующий файл — сразу отказ
    (xdg-open возвращает 0 даже для несуществующего пути).
    """
    if not file_path or not os.path.exists(file_path):
        applog.get_logger("files").warning(
            "Открытие отменено: файла нет на диске (%s)",
            os.path.splitext(str(file_path or ""))[1].lower(),
        )
        ui.alert(
            ROOT_HOLDER.get(), "Файл не найден",
            f"Файл не найден на диске:\n{file_path or '—'}", "error",
        )
        return False
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)  # noqa: S606
        elif platform.system() == "Darwin":
            if subprocess.run(["open", file_path], check=False).returncode != 0:
                return False
        else:
            if subprocess.run(["xdg-open", file_path], check=False).returncode != 0:
                return False
        return True
    except Exception as e:
        applog.get_logger("files").warning(
            "Не удалось открыть файл (%s): %s",
            os.path.splitext(str(file_path))[1].lower(), type(e).__name__,
        )
        ui.alert(ROOT_HOLDER.get(), "Не удалось открыть файл", str(e), "error")
        return False


def show_in_folder(file_path: str) -> None:
    """Показывает файл в проводнике, НЕ запуская его (безопасно для любых типов)."""
    if not file_path or not os.path.exists(file_path):
        ui.alert(
            ROOT_HOLDER.get(), "Файл не найден",
            file_path or "Путь к файлу не указан.", "error",
        )
        return
    try:
        if platform.system() == "Windows":
            subprocess.run(
                ["explorer", "/select,", os.path.normpath(file_path)], check=False,
            )
        elif platform.system() == "Darwin":
            subprocess.run(["open", "-R", file_path], check=False)
        else:
            subprocess.run(["xdg-open", os.path.dirname(file_path)], check=False)
    except Exception as exc:
        ui.alert(ROOT_HOLDER.get(), "Не удалось открыть папку", str(exc), "error")


def save_docx_atomic(doc, path: str) -> None:
    """Пишет .docx во временный файл, проверяет ZIP и только затем заменяет итоговый.

    Так аварийное завершение не оставляет обрезанный файл под «нормальным»
    именем — при сбое остаётся только *.docx.tmp, который удаляется.
    """
    tmp = path + ".tmp"
    try:
        doc.save(tmp)
        with zipfile.ZipFile(tmp) as archive:
            if archive.testzip() is not None:
                raise OSError("файл документа повреждён при записи")
        os.replace(tmp, path)
    except Exception:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise


def copy_with_zone_identifier(src: str, dest: str) -> None:
    """Копирует файл, сохраняя метку «скачано из интернета» (Mark-of-the-Web).

    shutil.copy2 не переносит альтернативные потоки NTFS, поэтому
    импортированный документ открывался без Protected View.
    """
    shutil.copy2(src, dest)
    if sys.platform != "win32":
        return
    zone_stream = ":Zone.Identifier"
    try:
        with open(src + zone_stream, "rb") as fh:
            zone = fh.read()
    except OSError:
        return
    try:
        with open(dest + zone_stream, "wb") as fh:
            fh.write(zone)
    except OSError:
        pass


# ---------- DOCX: замена плейсхолдеров (в т.ч. если разбиты на runs) ----------


class _RootHolder:
    """Позволяет функциям вне класса показывать алерты, не таща за собой ссылку на окно."""

    def __init__(self):
        self._root = None

    def set(self, root):
        self._root = root

    def get(self):
        return self._root


ROOT_HOLDER = _RootHolder()

# ---------- Параметры разметки ----------
ENTRY_WIDTH = 300
SHORT_WIDTH = 150
DAY_WIDTH = 56
MONTH_WIDTH = 146
LEFT_PANEL_WIDTH = 400
SIDEBAR_WIDTH = 210
BASE_WIDTH = 1280
BASE_HEIGHT = 740

NAV_PAGES = (
    ("counterparty", "Добавить контрагента", "person"),
    ("contract", "Договор поставки", "doc"),
    ("additional", "Доп. соглашение", "pencil"),
    ("spec", "Спецификация", "bank"),
    ("registry", "Реестр", "list"),
    ("settings", "Настройки", "gear"),
)

KIND_TO_PAGE = {
    storage.KIND_CONTRACT: "contract",
    storage.KIND_ADDITIONAL: "additional",
    storage.KIND_SPEC: "spec",
}

DIGITS_FIELDS = {F.INN, F.KPP, F.OGRN, F.OGRNIP, F.ACCOUNT, F.CORR_ACCOUNT, F.BIK}


def validate_digits_input(new_value):
    return new_value == "" or new_value.isdigit()


def validate_phone_input(new_value):
    """Ввод на лету: разрешены только символы маски телефона."""
    return _PHONE_ALLOWED_RE.match(new_value or "") is not None


def validate_day_input(new_value):
    """День: 1–31, не более двух цифр."""
    return new_value == "" or (new_value.isdigit() and len(new_value) <= 2 and int(new_value) <= 31)


def validate_limit_input(new_value):
    """Лимит: цифры и пробелы-разделители (до 15 цифр)."""
    if not new_value:
        return True
    if any(not (ch.isdigit() or ch == " ") for ch in new_value):
        return False
    return len(digits_only(new_value)) <= 15


def validate_term_input(new_value):
    """Отсрочка: целое число дней, не более четырёх цифр."""
    return new_value == "" or (new_value.isdigit() and len(new_value) <= 4)


def validate_price_input(new_value):
    """Цена: цифры, разряды и запятая/точка (копейки), до 12 цифр."""
    if not new_value:
        return True
    if any(not (ch.isdigit() or ch in " .,") for ch in new_value):
        return False
    return len(digits_only(new_value)) <= 14


# Имена, которые Windows не разрешает как имя файла или папки.
_RESERVED_FILENAMES = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{i}" for i in range(1, 10)}
    | {f"LPT{i}" for i in range(1, 10)}
)
_FORBIDDEN_FILENAME_CHARS = frozenset('<>:"/\\|?*')


def safe_filename(name: str, fallback: str = "") -> str:
    """Готовит безопасное имя файла/папки для Windows.

    Защищает от выхода за пределы папки документов и «неудобных» имён:
    убирает запрещённые и управляющие символы, точки/пробелы на конце,
    зарезервированные имена (CON, PRN, NUL, COM1…) и ограничивает длину,
    добавляя короткий хэш для уникальности.
    """
    if not name:
        return fallback
    text = "".join(
        char for char in str(name)
        if char not in _FORBIDDEN_FILENAME_CHARS and ord(char) >= 32
    )
    text = re.sub(r"\s+", " ", text).strip().strip(".")
    if not text or text in {".", ".."}:
        return fallback
    if text.split(".")[0].upper() in _RESERVED_FILENAMES:
        text = "_" + text
    if len(text) > MAX_FILENAME_LENGTH:
        digest = hashlib.sha1(text.encode("utf-8")).hexdigest()[:8]
        text = text[: MAX_FILENAME_LENGTH - 9].rstrip(" .") + "~" + digest
    return text or fallback


# --------------------------------------------------------------------------- #
#  Таблица позиций спецификации (excel-подобная, вставка из буфера обмена)
# --------------------------------------------------------------------------- #
class SpecTable(ctk.CTkFrame):
    """Excel-подобная таблица позиций с поддержкой вставки из буфера.

    Столбцы: Наименование, Кол-во, Ед. изм., Цена за ед. с НДС (руб.).
    Ctrl+V в любой ячейке вставляет прямоугольный блок данных (из Excel —
    столбцы, разделённые табуляцией); при нехватке строки добавляются сами.
    Одинарный клик выделяет строку, Delete удаляет выделенную строку,
    двойной клик по ячейке включает правку.
    """

    HEADERS = ("№", "Наименование", "Кол-во", "Ед. изм.", "Цена за ед. с НДС (руб.)")
    KEYS = ("name", "qty", "unit", "price")
    MIN_ROWS = 10
    MAX_ROWS = 500
    # Ширина фиксированных колонок в логических пикселях (№, Кол-во, Ед. изм., Цена).
    COL_WIDTHS = {0: 34, 2: 64, 3: 84, 4: 150}

    def __init__(self, master, scale: float = 1.0, on_change=None, **kwargs):
        super().__init__(master, fg_color="transparent", corner_radius=0, **kwargs)
        self.scale = scale
        self.on_change = on_change
        self._rows: list[dict] = []
        self._batch = False  # подавление on_change при массовой вставке
        self._selected_row: int | None = None
        self._digits_vcmd = self.register(self._validate_digits)
        self._price_vcmd = self.register(self._validate_price)
        self._entry_font = tkfont.Font(font=ui.font(12))
        # Отмена разрушительных действий с таблицей (п. 4.4).
        self._undo_stack: list[list[dict]] = []
        self._undo_limit = 30

        # -- Панель инструментов ---------------------------------------- #
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, int(4 * self.scale)))
        ui.IconButton(
            toolbar, "plus", self.add_row, size=26, tooltip="Добавить строку",
        ).pack(side="left")
        ui.IconButton(
            toolbar, "minus", self.remove_last, size=26, tooltip="Удалить последнюю строку",
        ).pack(side="left", padx=(int(4 * self.scale), 0))
        ui.IconButton(
            toolbar, "trash", self.clear, size=26, tooltip="Очистить таблицу",
        ).pack(side="left", padx=(int(4 * self.scale), 0))

        # Кол-во строк: задаётся в поле, применяется по Enter.
        self._count_var = tk.StringVar(value=str(self.MIN_ROWS))
        ctk.CTkLabel(
            toolbar, text="Строк:", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"],
        ).pack(side="right")
        count_entry = ui.entry(
            toolbar, width=int(56 * self.scale),
            textvariable=self._count_var,
            validatecommand=(self._digits_vcmd, "%P"),
        )
        count_entry.pack(side="right", padx=(int(8 * self.scale), 0))
        self._count_entry = count_entry
        try:  # фокус внутри внутреннего tk.Entry — биндим именно его
            count_entry.bind_entry("<Return>", lambda _e: self._apply_row_count())
            count_entry.bind_entry("<KP_Enter>", lambda _e: self._apply_row_count())
            count_entry.bind_entry("<Control-z>", self._on_undo)
            count_entry.bind_entry("<Control-Z>", self._on_undo)
        except Exception:
            pass

        # -- Шапка таблицы: ячейки фиксированной ширины, подписи по центру -- #
        header = tk.Frame(self)
        header.pack(fill="x", padx=1)
        header.grid_columnconfigure(1, weight=1)
        for col in (0, 2, 3, 4):
            header.grid_columnconfigure(col, weight=0)
        self._header_frame = header
        self._header_cells: list[tk.Frame] = []
        self._header_labels: list[tk.Label] = []
        for col, title in enumerate(self.HEADERS):
            cell = tk.Frame(header, height=int(30 * self.scale))
            if col in self.COL_WIDTHS:
                cell.configure(width=int(self.COL_WIDTHS[col] * self.scale))
            cell.grid(row=0, column=col, sticky="ew",
                      padx=(2 if col == 0 else 3, 6 if col == 0 else 3))
            cell.pack_propagate(False)
            cell.grid_propagate(False)
            label = tk.Label(
                cell, text=title, font=ui.font(11, "bold"),
                anchor="center", justify="center",
                wraplength=int(150 * self.scale),
            )
            label.pack(fill="both", expand=True)
            self._header_cells.append(cell)
            self._header_labels.append(label)

        # -- Тело таблицы ------------------------------------------------ #
        self.body = ctk.CTkScrollableFrame(
            self, height=int(216 * self.scale),
            fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        self.body.pack(fill="x")
        self.body.grid_columnconfigure(0, weight=0)
        self._schedule_height_update()
        self.body.grid_columnconfigure(1, weight=1)
        for col in (2, 3, 4):
            self.body.grid_columnconfigure(col, weight=0)

        ui.register_theme_listener(self._apply_theme)
        self._apply_theme()

        for _ in range(self.MIN_ROWS):
            self._add_row()

    # -- Строки ----------------------------------------------------------- #
    @property
    def rows(self) -> list[dict]:
        """Строки таблицы: словари со StringVar «name», «qty», «unit», «price»."""
        return self._rows

    def add_row(self, values: dict | None = None) -> dict:
        return self._add_row(values)

    def _entry_width_chars(self, px: int) -> int:
        """Ширина ячейки в символах, соответствующая ширине шапки в пикселях."""
        digit = max(1, self._entry_font.measure("0"))
        return max(3, int((px * self.scale - 10) // digit))

    def _add_row(self, values: dict | None = None) -> dict:
        index = len(self._rows)
        values = values or {}

        num = tk.Label(self.body, text=str(index + 1), font=ui.font(11), anchor="center")
        num.grid(row=index, column=0, sticky="ew",
                 padx=(2, 6), ipady=int(3 * self.scale))
        num.bind("<Button-1>", self._on_cell_click)

        row: dict = {"num": num, "cells": {}}
        for col, key in enumerate(self.KEYS, start=1):
            var = tk.StringVar(value=str(values.get(key, "") or ""))
            entry = tk.Entry(
                self.body, textvariable=var, font=ui.font(12),
                relief="flat", justify="left" if key == "name" else "center",
                highlightthickness=1, state="readonly",  # правка — по двойному клику
                validatecommand=(self._digits_vcmd, "%P") if key == "qty"
                else (self._price_vcmd, "%P") if key == "price" else None,
            )
            if col in self.COL_WIDTHS:  # ширина синхронизирована со шапкой
                entry.configure(width=self._entry_width_chars(self.COL_WIDTHS[col]))
            entry.grid(row=index, column=col, sticky="ew",
                       padx=(0 if col == 1 else 3, 3), ipady=int(3 * self.scale))

            entry.bind("<Button-1>", self._on_cell_click)
            entry.bind("<Double-Button-1>", self._begin_edit)
            entry.bind("<FocusOut>", self._end_edit)
            entry.bind("<Delete>", self._on_delete_key)
            entry.bind("<Control-z>", self._on_undo)
            entry.bind("<Control-Z>", self._on_undo)
            entry.bind("<Return>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<KP_Enter>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<Up>", lambda e: self._move_row(e.widget, -1))
            entry.bind("<Down>", lambda e: self._move_row(e.widget, +1))
            entry.bind("<Tab>", lambda e: self._move_cell(e.widget, +1))
            entry.bind("<ISO_Left_Tab>", lambda e: self._move_cell(e.widget, -1))
            entry.bind("<<Paste>>", self._on_paste)
            try:
                ui.fix_clipboard_shortcuts(entry)
            except Exception:
                pass

            var.trace_add("write", lambda *_a: self._cell_changed())
            row[key] = var
            row["cells"][key] = entry

        self._rows.append(row)
        # П. 5.4: стилизуем только новую строку. Раньше _apply_theme()
        # перестилизовывала ВСЕ строки на каждое добавление — задание
        # 500 строк было O(n²) и блокировало интерфейс.
        if not self._batch:
            self._style_row(index)
            self._schedule_height_update()
        return row

    # -- Высота тела таблицы ------------------------------------------------ #
    def _schedule_height_update(self):
        """Пересчитать высоту таблицы, когда виджеты будут отображены."""
        try:
            self.after_idle(self._update_body_height)
        except Exception:
            pass

    def _update_body_height(self):
        """Высота тела — ровно под строки, без пустого места снизу.

        Если строк больше, чем помещается на экране, высота ограничивается
        максимумом и включается прокрутка.
        """
        try:
            if not self._rows:
                return
            row_h = self._rows[0]["cells"]["name"].winfo_reqheight()
            if row_h <= 1:  # виджеты ещё не отображены — оценка
                row_h = int(24 * self.scale)
            row_h += int(6 * self.scale)  # ipady вокруг ячейки
            total = row_h * len(self._rows) + int(4 * self.scale)
            max_h = int(360 * self.scale)
            self.body.configure(height=max(int(140 * self.scale), min(total, max_h)))
        except Exception:
            pass

    # -- Выделение строки и правка ячеек ------------------------------------- #
    def _row_of_widget(self, widget) -> int | None:
        for i, row in enumerate(self._rows):
            if widget is row["num"] or widget in row["cells"].values():
                return i
        return None

    def _on_cell_click(self, event=None):
        """Одинарный клик: выделить строку (правка — только по двойному клику)."""
        if event is not None:
            index = self._row_of_widget(event.widget)
            if index is not None and index != self._selected_row:
                previous = self._selected_row
                self._selected_row = index
                # Стилизуем только прежнюю и новую строку, а не всю таблицу (п. 5.4).
                if previous is not None:
                    self._style_row(previous)
                self._style_row(index)
        return None

    def _begin_edit(self, event):
        """Двойной клик: разрешить правку ячейки."""
        entry = event.widget
        try:
            entry.configure(state="normal")
            entry.focus_set()
            entry.icursor("end")
        except Exception:
            pass
        return None

    def _end_edit(self, event=None):
        """После выхода из ячейки возвращаем режим «только чтение»."""
        if event is None:
            return None
        entry = event.widget
        try:
            if entry.winfo_exists() and str(entry.cget("state")) == "normal":
                entry.configure(state="readonly")
        except Exception:
            pass
        return None

    def _on_delete_key(self, event):
        """Delete: если ячейка не редактируется — удалить выделенную строку."""
        entry = event.widget
        if str(entry.cget("state")) == "normal":
            return None  # идёт правка — обычное удаление символа
        index = (self._selected_row if self._selected_row is not None
                 else self._row_of_widget(entry))
        if index is not None and 0 <= index < len(self._rows):
            self._delete_row(index)
            return "break"
        return None

    def _delete_row(self, index: int):
        if len(self._rows) <= 1:
            return
        row = self._rows[index]
        if any(row[key].get().strip() for key in self.KEYS):
            if not self._confirm_destructive("Удалить строку с данными?"):
                return
        self._push_undo()
        self._destroy_row(index)
        self._selected_row = None
        self._renumber()
        self._notify()

    def set_row_count(self, count: int):
        """Задаёт точное количество строк (добавляет или удаляет лишние)."""
        count = max(1, min(int(count), self.MAX_ROWS))
        if count == len(self._rows):
            return
        if count < len(self._rows):
            lost = [
                index for index in range(count, len(self._rows))
                if any(self._rows[index][key].get().strip() for key in self.KEYS)
            ]
            if lost and not self._confirm_destructive(
                f"Уменьшение числа строк удалит заполненные строки: {len(lost)}.\n\n"
                "Продолжить?"
            ):
                self._count_var.set(str(len(self._rows)))
                return
            self._push_undo()
        self._batch = True
        try:
            while len(self._rows) < count:
                self._add_row()
            while len(self._rows) > count:
                self._destroy_row(len(self._rows) - 1)
        finally:
            self._batch = False
        self._apply_theme()
        self._schedule_height_update()
        self._selected_row = None
        self._renumber()
        self._count_var.set(str(count))
        self._notify()

    def _apply_row_count(self):
        value = digits_only(self._count_var.get())
        self.set_row_count(int(value or self.MIN_ROWS))

    def _renumber(self):
        for i, row in enumerate(self._rows):
            try:
                row["num"].configure(text=str(i + 1))
            except Exception:
                pass

    def _destroy_row(self, index: int):
        row = self._rows.pop(index)
        for widget in (row["num"], *row["cells"].values()):
            try:
                widget.destroy()
            except Exception:
                pass
        self._schedule_height_update()

    # -- Отмена и подтверждения (п. 4.4) --------------------------------- #
    def _snapshot(self) -> list[dict]:
        return [{key: row[key].get() for key in self.KEYS} for row in self._rows]

    def _push_undo(self) -> None:
        try:
            self._undo_stack.append(self._snapshot())
            if len(self._undo_stack) > self._undo_limit:
                self._undo_stack.pop(0)
        except Exception:
            pass

    def _confirm_destructive(self, message: str) -> bool:
        try:
            return ui.confirm(
                self.winfo_toplevel(), "Таблица позиций", message,
                ok_text="Продолжить", cancel_text="Отмена",
            )
        except Exception:
            return True

    def _on_undo(self, _event=None):
        self.undo()
        return "break"

    def undo(self) -> bool:
        """Возвращает таблицу к состоянию перед последним разрушительным действием."""
        if not self._undo_stack:
            return False
        snapshot = self._undo_stack.pop()
        self._batch = True
        try:
            self._reset_rows()
            while len(self._rows) < len(snapshot):
                self._add_row()
            while len(self._rows) > len(snapshot) and len(self._rows) > self.MIN_ROWS:
                self._destroy_row(len(self._rows) - 1)
            for row, values in zip(self._rows, snapshot):
                for key in self.KEYS:
                    row[key].set(values.get(key, ""))
            self._count_var.set(str(len(self._rows)))
        finally:
            self._batch = False
        self._selected_row = None
        self._renumber()
        self._apply_theme()
        self._schedule_height_update()
        self._notify()
        return True

    def remove_last(self):
        if len(self._rows) <= 1:
            return
        last = self._rows[-1]
        if any(last[key].get().strip() for key in self.KEYS):
            if not self._confirm_destructive("Удалить последнюю строку с данными?"):
                return
        self._push_undo()
        self._destroy_row(len(self._rows) - 1)
        self._selected_row = None
        self._notify()

    def clear(self):
        filled = sum(
            1 for row in self._rows
            if any(row[key].get().strip() for key in self.KEYS)
        )
        if filled and not self._confirm_destructive(
            f"Очистить таблицу? Заполненных строк: {filled}."
        ):
            return
        if filled:
            self._push_undo()
        self._reset_rows()
        self._selected_row = None
        self._notify()

    def _reset_rows(self):
        """Очистка без уведомления (используется перед массовой вставкой)."""
        while self._rows:
            self._destroy_row(len(self._rows) - 1)
        for _ in range(self.MIN_ROWS):
            self._add_row()
        self._apply_theme()
        self._schedule_height_update()

    # -- Данные ------------------------------------------------------------ #
    def get_items(self) -> list[dict]:
        """Позиции с заполненным наименованием; суммы считаются на лету.

        Кол-во и цена — Decimal: цена может содержать копейки,
        кол-во может быть дробным (например, 2,5 м).
        """
        items = []
        for row in self._rows:
            name = row["name"].get().strip()
            if not name:
                continue
            items.append({
                "name": name,
                "qty": parse_decimal(row["qty"].get()),
                "unit": row["unit"].get().strip() or "шт.",
                "price": parse_decimal(row["price"].get()),
            })
        return items

    def total(self) -> Decimal:
        return sum(
            (item["qty"] * item["price"] for item in self.get_items()),
            Decimal("0"),
        )

    # -- Вставка из буфера -------------------------------------------------- #
    def paste_clipboard(self):
        """Кнопка «Вставить из Excel»: заменить содержимое таблицы буфером."""
        if any(
            row[key].get().strip() for row in self._rows for key in self.KEYS
        ):
            self._push_undo()
        self._reset_rows()
        self._paste_block((0, 0))

    def _on_paste(self, event):
        cell = self._cell_of(event.widget)
        if cell is None:
            return None
        if any(
            row[key].get().strip() for row in self._rows for key in self.KEYS
        ):
            self._push_undo()
        self._paste_block(cell)
        return "break"

    def _paste_block(self, start_cell: tuple[int, int]) -> bool:
        """Вставляет блок текста (TSV) начиная с ячейки start_cell."""
        try:
            raw = self.clipboard_get()
        except Exception:
            return False
        raw = raw.replace("\r\n", "\n").replace("\r", "\n")
        matrix = [line.split("\t") for line in raw.split("\n")]
        while matrix and all(not cell.strip() for cell in matrix[-1]):
            matrix.pop()
        if not matrix:
            return False

        r0, c0 = start_cell
        need = r0 + len(matrix) - len(self._rows)
        for _ in range(max(0, need)):
            self._add_row()

        self._batch = True
        try:
            for dr, values in enumerate(matrix):
                mapped = self._map_row(values)
                for dc, value in enumerate(mapped):
                    col = c0 + dc
                    if col >= len(self.KEYS):
                        break
                    self._rows[r0 + dr][self.KEYS[col]].set(value)
        finally:
            self._batch = False
        self._notify()
        return True

    def _map_row(self, cells: list[str]) -> list[str]:
        """Эвристика раскладки вставленных столбцов по колонкам таблицы."""
        cells = [cell.strip() for cell in cells]
        while cells and not cells[-1]:
            cells.pop()
        n = len(cells)
        if n == 0:
            return ["", "", "", ""]
        if n == 1:
            return [cells[0], "", "", ""]
        if n == 2:  # Наименование, Кол-во
            return [cells[0], self._normalize_number(cells[1]), "", ""]
        if n == 3:
            # «Наименование, Кол-во, Цена» либо «Наименование, Кол-во, Ед. изм.»
            if self._normalize_number(cells[2]):
                return [cells[0], self._normalize_number(cells[1]), "",
                        self._normalize_number(cells[2])]
            return [cells[0], self._normalize_number(cells[1]), cells[2], ""]
        return [cells[0], self._normalize_number(cells[1]), cells[2],
                self._normalize_number(cells[3])]

    @staticmethod
    def _normalize_number(text: str) -> str:
        """«1 000,50» -> «1000.50»: убирает разрядные разделители,
        сохраняет дробную часть (копейки).

        Правила: если встречаются «,» и «.» — правый из них десятичный;
        одиночная запятая считается десятичным разделителем; повторные
        запятые/точки считаются разрядными и удаляются.
        """
        text = str(text).replace("\xa0", "").replace(" ", "")
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "")   # «2.052,59» -> «2052,59»
            else:
                text = text.replace(",", "")   # «2,052.59» -> «2052.59»
        if text.count(",") > 1:
            text = text.replace(",", "")       # «1,000,000» — разрядные
        if text.count(".") > 1:
            text = text.replace(".", "")       # «1.000.000» — разрядные
        text = text.replace(",", ".")
        cleaned = "".join(ch for ch in text if ch.isdigit() or ch == ".")
        return cleaned.rstrip(".")

    # -- Навигация ----------------------------------------------------------- #
    def _cell_of(self, entry) -> tuple[int, int] | None:
        for r, row in enumerate(self._rows):
            for key, widget in row["cells"].items():
                if widget is entry:
                    return (r, self.KEYS.index(key))
        return None

    def _move_row(self, entry, delta: int):
        cell = self._cell_of(entry)
        if cell is None:
            return None
        r, c = cell
        r = min(max(r + delta, 0), len(self._rows) - 1)
        return self._focus_cell((r, c))

    def _move_cell(self, entry, delta: int):
        cell = self._cell_of(entry)
        if cell is None:
            return None
        r, c = cell
        c += delta
        if c >= len(self.KEYS):
            if r + 1 < len(self._rows):
                return self._focus_cell((r + 1, 0))
            self._add_row()
            return self._focus_cell((r + 1, 0))
        if c < 0:
            if r > 0:
                return self._focus_cell((r - 1, len(self.KEYS) - 1))
            c = 0
        return self._focus_cell((r, c))

    def _focus_cell(self, cell: tuple[int, int]):
        r, c = cell
        widget = self._rows[r]["cells"][self.KEYS[c]]
        widget.focus_set()
        widget.icursor("end")
        return "break"

    # -- Служебное ------------------------------------------------------------ #
    def _validate_digits(self, new_value: str) -> bool:
        """Кол-во: цифры и запятая/точка (может быть дробным)."""
        return new_value == "" or all(ch.isdigit() or ch in ",." for ch in new_value)

    def _validate_price(self, new_value: str) -> bool:
        """Цена: цифры, разряды и запятая/точка (копейки)."""
        return new_value == "" or all(
            ch.isdigit() or ch in " .," for ch in new_value
        )

    def _cell_changed(self):
        self._notify()

    def _notify(self):
        if self._batch:
            return
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass

    def _style_row(self, index: int) -> None:
        """Стилизует одну строку (фон, рамки, цвета) с учётом выделения."""
        border = ui.resolve(ui.COLORS["entry_border"])
        accent = ui.resolve(ui.COLORS["accent"])
        base_bg = ui.resolve(ui.COLORS["entry_bg"])
        selected_bg = ui.resolve(ui.COLORS["sidebar_active"])
        bg = selected_bg if index == self._selected_row else base_bg
        row = self._rows[index]
        try:
            row["num"].configure(
                bg=bg,
                fg=ui.resolve(ui.COLORS["text_tertiary"]),
            )
        except Exception:
            pass
        for entry in row["cells"].values():
            try:
                entry.configure(
                    bg=bg,
                    readonlybackground=bg,
                    fg=ui.resolve(ui.COLORS["text"]),
                    insertbackground=ui.resolve(ui.COLORS["text"]),
                    highlightbackground=border,
                    highlightcolor=accent,
                )
            except Exception:
                pass

    def _apply_theme(self):
        header_bg = ui.resolve(ui.COLORS["chrome"])
        for cell in self._header_cells:
            try:
                cell.configure(bg=header_bg)
            except Exception:
                pass
        for label in self._header_labels:
            try:
                label.configure(bg=header_bg, fg=ui.resolve(ui.COLORS["text_secondary"]))
            except Exception:
                pass
        for index in range(len(self._rows)):
            self._style_row(index)


class ContractApp:
    def __init__(self):
        ui.init_fonts()
        self._settings = load_settings()
        ctk.set_appearance_mode(str(self._settings.get("theme", "light")).lower())

        self.root = ctk.CTk()
        self.root.title("EW.DocHub — контрагенты и документы")
        self.root.configure(fg_color=ui.COLORS["window"])
        try:  # иконка приложения (в сборке и в запуске из исходников)
            self.root.iconbitmap(resource_path("app_icon.ico"))
        except Exception:
            pass
        ROOT_HOLDER.set(self.root)

        self.root.update_idletasks()
        self.scale = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        ui.set_scale(self.scale)

        self.store = storage.DataStore()
        if self.store.load_warning:
            # Сообщаем после построения окна, чтобы алерт был поверх интерфейса.
            self.root.after(300, lambda: ui.alert(
                self.root, "Реестр загружен не полностью",
                self.store.load_warning, "error",
            ))

        # Служебное состояние этапа 2: сверка с папкой, корзина, недоступная папка.
        self._orphans: list[str] = []
        self._docs_dir_fallback_active = False
        self._docs_dir_retry_count = 0
        self._docs_dir_dialog_active = False
        self._log = applog.get_logger("main")

        self.pages: dict[str, ctk.CTkFrame] = {}
        self.nav_items: dict[str, ui.NavItem] = {}
        self.scrollers: list[ctk.CTkScrollableFrame] = []
        self.current_page = "contract"
        self._zoomed = False
        self._fullscreen = False
        self._fullscreen_restore = ""
        self._restore_geometry = ""
        self._min_width = int(MIN_WINDOW_WIDTH * self.scale)
        self._min_height = int(MIN_WINDOW_HEIGHT * self.scale)
        self.root.minsize(self._min_width, self._min_height)
        self._resize_origin = None
        self._resize_start = None

        self._digits_vcmd = self.root.register(validate_digits_input)
        self._phone_vcmd = self.root.register(validate_phone_input)
        self._day_vcmd = self.root.register(validate_day_input)
        self._limit_vcmd = self.root.register(validate_limit_input)
        self._term_vcmd = self.root.register(validate_term_input)
        self._price_vcmd = self.root.register(validate_price_input)

        # Inline-валидация: переменные подсказок об ошибках под полями страниц.
        self._error_vars: dict[tuple[str, str], tk.StringVar] = {}
        self._page_widgets: dict[tuple[str, str], tk.Misc] = {}

        # Карточка «Добавить контрагента» — отдельная страница основного окна.
        self._cp_entries: dict[str, dict[str, tk.Misc]] = {"ooo": {}, "ip": {}}
        self._cp_error_vars: dict[str, dict[str, tk.StringVar]] = {"ooo": {}, "ip": {}}
        self._cp_date_error_vars = {"ooo": tk.StringVar(value=""), "ip": tk.StringVar(value="")}
        self._cp_status_var = tk.StringVar(value="")
        self._cp_busy = False
        self._cp_busy_after = None
        self._cp_busy_phase = 0

        # Единый «текущий контрагент» и защита несохранённых данных (п. 4.2, 4.4).
        self._current_party_id = None
        self._current_party_vars: dict[str, tk.StringVar] = {}
        self._syncing_party = False
        self._form_baseline: dict[str, tuple] = {}
        self._registry_search_after = None
        self._toast_after = None
        self._toast_path: str | None = None

        self._place_window()
        ui.make_borderless(self.root)

        self._build_titlebar()
        self._build_body()
        self._build_pages()
        self._reset_all_form_baselines()
        self._install_resize_handles()
        self._select_page("contract")
        # Первый запуск: сразу открываем добавление контрагента (п. 4.5).
        if not self.store.active_counterparties():
            self.root.after(250, lambda: self._select_page("counterparty"))
        self.root.after(250, self._refresh_scrollbars)
        self.root.after(600, self._cleanup_temp_documents)
        # Автоочистка старой корзины и сверка «реестр ↔ папка документов».
        self.root.after(900, self._cleanup_expired_trash)
        self.root.after(1500, self._scan_documents_consistency)

        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<Escape>", self._on_escape)
        self.root.bind("<F11>", lambda _event: self.toggle_fullscreen())
        # Ctrl+S и Enter в формах сохраняют текущий документ (п. 4.5).
        self.root.bind("<Control-s>", self._on_save_shortcut)
        self.root.bind("<Control-S>", self._on_save_shortcut)
        self.root.bind("<Control-Return>", self._on_save_shortcut)
        self.root.bind("<Return>", self._on_enter_save)
        # У borderless-окна Alt+F4 может не доходить до WM_DELETE_WINDOW —
        # страхуемся явной привязкой.
        self.root.bind("<Alt-KeyPress-F4>", lambda _event: self.close())
        self.root.bind("<Control-n>", lambda _event: self._select_page("counterparty"))
        self.root.bind("<Control-N>", lambda _event: self._select_page("counterparty"))
        ui.refresh_taskbar(self.root)
        self._enable_drop_files()

    # ------------------------------------------------------------------ #
    #  Окно
    # ------------------------------------------------------------------ #
    def _place_window(self):
        self.root.update_idletasks()
        area_x, area_y, area_w, area_h = ui.workarea(self.root)
        # Минимальный размер не должен превышать рабочую область: иначе
        # на экране 1366×768 с масштабом 150 % окно нельзя сузить (п. 6.3).
        self._min_width = min(self._min_width, max(320, area_w))
        self._min_height = min(self._min_height, max(240, area_h))
        try:
            self.root.minsize(self._min_width, self._min_height)
        except Exception:
            pass
        saved = self._settings.get("window") or {}
        try:
            width = int(saved.get("width", 0))
            height = int(saved.get("height", 0))
            x = int(saved.get("x", 0))
            y = int(saved.get("y", 0))
        except (TypeError, ValueError):
            width = height = x = y = 0
        if width <= 0 or height <= 0:
            width = int(BASE_WIDTH * self.scale)
            height = int(BASE_HEIGHT * self.scale)
            x = area_x + (area_w - width) // 2
            y = area_y + max(int(20 * self.scale), (area_h - height) // 2 - int(30 * self.scale))
        width = max(self._min_width, min(width, area_w))
        height = max(self._min_height, min(height, area_h))
        # Окно должно оставаться хотя бы частично видимым на экране.
        x = min(max(x, -width + int(120 * self.scale)), area_x + area_w - int(120 * self.scale))
        y = min(max(y, area_y - int(8 * self.scale)), area_y + area_h - int(60 * self.scale))
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self._restore_geometry = self.root.geometry()
        if saved.get("maximized"):
            self.root.after(150, self.toggle_zoom)

    def _build_titlebar(self):
        self.titlebar = tk.Frame(self.root, bg=ui.resolve(ui.COLORS["chrome"]), height=int(50 * self.scale))
        self.titlebar.pack(fill="x", side="top")
        self.titlebar.pack_propagate(False)

        # Кнопки разделов в шапке окна (вместо бокового меню и меню «Файл»).
        self.nav_holder = tk.Frame(self.titlebar, bg=ui.resolve(ui.COLORS["chrome"]))
        self.nav_holder.pack(side="left", padx=(int(12 * self.scale), 0), pady=(int(10 * self.scale), 0))
        for key, title, icon in NAV_PAGES:
            item = ui.TopNavItem(
                self.nav_holder, title, icon,
                command=lambda page_key=key: self._select_page(page_key),
                height=30,
            )
            item.pack(side="left", padx=(int(3 * self.scale), 0))
            self.nav_items[key] = item

        self.lights = ui.WindowControls(
            self.titlebar, on_close=self.close, on_minimize=self.minimize, on_zoom=self.toggle_zoom,
            height=int(50 * self.scale),
        )
        self.lights.pack(side="right")

        self.drag = ui.WindowDrag(self.root, on_start=self._start_move)
        self.drag.attach(self.titlebar)
        self.titlebar.bind("<Double-Button-1>", lambda _event: self.toggle_zoom())

        self.chrome_line = tk.Frame(self.root, bg=ui.resolve(ui.COLORS["chrome_line"]), height=1)
        self.chrome_line.pack(fill="x", side="top")

        ui.register_theme_listener(self._sync_chrome)

    def _sync_chrome(self):
        try:
            self.titlebar.configure(bg=ui.resolve(ui.COLORS["chrome"]))
            self.nav_holder.configure(bg=ui.resolve(ui.COLORS["chrome"]))
            self.chrome_line.configure(bg=ui.resolve(ui.COLORS["chrome_line"]))
        except Exception:
            pass

    def _start_move(self, event=None):
        """Перетаскивание развёрнутого окна: восстанавливаем размер
        и ставим окно под курсор (пропорциональный перенос)."""
        if not self._zoomed:
            return
        self._zoomed = False
        self.lights.set_zoomed(False)
        ui.restore_window(self.root)
        geometry = self._restore_geometry or f"{int(BASE_WIDTH * self.scale)}x{int(BASE_HEIGHT * self.scale)}"
        try:
            width, height, _, _ = self._parse_geometry(geometry)
        except (ValueError, AttributeError):
            width = int(BASE_WIDTH * self.scale)
            height = int(BASE_HEIGHT * self.scale)
        if event is not None:
            try:
                rel = (event.x_root - self.root.winfo_rootx()) / max(self.root.winfo_width(), 1)
            except Exception:
                rel = 0.5
            rel = min(max(rel, 0.15), 0.85)
            x = int(event.x_root - width * rel)
            y = int(event.y_root - int(24 * self.scale))
            self.root.geometry(f"{int(width)}x{int(height)}+{x}+{y}")
        else:
            self.root.geometry(f"{int(width)}x{int(height)}")
        self.root.update_idletasks()

    def toggle_zoom(self):
        """Настоящая максимизация окна (Win32 SW_MAXIMIZE / SW_RESTORE)."""
        if self._fullscreen:
            self.toggle_fullscreen()
        if not self._zoomed:
            self._restore_geometry = self.root.geometry()
            ui.maximize_window(self.root)
            self._zoomed = True
        else:
            # Геометрию возвращаем вручную: максимизация задаётся без SW_MAXIMIZE,
            # поэтому возврат тоже делаем явно (панель задач не перекрывается).
            ui.restore_window(self.root, self._restore_geometry)
            self.root.update_idletasks()
            self._zoomed = False
        self.lights.set_zoomed(self._zoomed)

    def toggle_fullscreen(self):
        """Полноэкранный режим по F11 (Esc — выход)."""
        if not self._fullscreen:
            try:
                self._fullscreen_restore = self.root.geometry()
                self.root.geometry(
                    f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}+0+0"
                )
                self._fullscreen = True
            except Exception:
                self._fullscreen = False
        else:
            try:
                self.root.geometry(self._fullscreen_restore)
            except Exception:
                pass
            self._fullscreen = False
            self.root.update_idletasks()

    def _on_escape(self, _event=None):
        if self._fullscreen:
            self.toggle_fullscreen()
        return None

    def minimize(self):
        ui.minimize_window(self.root)

    @staticmethod
    def _parse_geometry(geometry: str):
        """'ШxВ+X+Y' -> (ш, в, x, y); бросает ValueError при битой строке."""
        size, _, pos = geometry.partition("+")
        width, _, height = size.partition("x")
        x, _, y = pos.partition("+")
        return int(width), int(height), int(x), int(y)

    def _install_resize_handles(self):
        """Невидимые полосы у краёв и углов: окно тянется мышью со всех сторон."""
        t = int(6 * self.scale)    # толщина края
        c = int(16 * self.scale)   # размер угла
        # Хэндлы верхней кромки делаем детьми шапки: они должны быть под
        # кнопками окна, иначе «n»/«nw»/«ne» перекрывают красный ховер
        # кнопки «закрыть» и перехватывают клики по ней.
        specs = (
            ("n",  "top_side",            "chrome",  dict(relx=0.0, rely=0.0, relwidth=1.0, height=t), self.titlebar),
            ("nw", "top_left_corner",     "chrome",  dict(width=c, height=c), self.titlebar),
            ("ne", "top_right_corner",    "chrome",  dict(relx=1.0, x=-c, width=c, height=c), self.titlebar),
            ("s",  "bottom_side",         "window",  dict(relx=0.0, rely=1.0, y=-t, relwidth=1.0, height=t), self.root),
            ("w",  "left_side",           "sidebar", dict(rely=0.0, relheight=1.0, width=t), self.root),
            ("e",  "right_side",          "content", dict(relx=1.0, x=-t, rely=0.0, relheight=1.0, width=t), self.root),
            ("sw", "bottom_left_corner",  "sidebar", dict(rely=1.0, y=-c, width=c, height=c), self.root),
            ("se", "bottom_right_corner", "window",  dict(relx=1.0, rely=1.0, x=-c, y=-c, width=c, height=c), self.root),
        )
        self._resize_handles = []
        for edge, cursor, color_key, place_kwargs, parent in specs:
            try:
                handle = tk.Frame(parent, bg=ui.resolve(ui.COLORS[color_key]), cursor=cursor)
            except tk.TclError:
                # Минимальные X-серверы могут не иметь системных курсоров —
                # окно остаётся работоспособным без курсоров растяжения.
                handle = tk.Frame(parent, bg=ui.resolve(ui.COLORS[color_key]))
            handle.place(**place_kwargs)
            handle.bind("<ButtonPress-1>", lambda event, e=edge: self._resize_press(event, e))
            handle.bind("<B1-Motion>", lambda event, e=edge: self._resize_drag(event, e))
            handle.bind("<Double-Button-1>", lambda _event, e=edge: self.toggle_zoom())
            ui.register_theme_listener(
                lambda h=handle, key=color_key: self._sync_handle_theme(h, key)
            )
            self._resize_handles.append(handle)

        # Шапка — над корневыми хэндлами (иначе «e» перекрывал правый край
        # кнопок окна), а сами кнопки — над хэндлами шапки.
        self.titlebar.tkraise()
        tk.Misc.tkraise(self.lights)

    @staticmethod
    def _sync_handle_theme(handle, color_key):
        try:
            handle.configure(bg=ui.resolve(ui.COLORS[color_key]))
        except Exception:
            pass

    def _resize_press(self, event, edge):
        if self._zoomed:
            self._zoomed = False
            self.lights.set_zoomed(False)
            ui.restore_window(self.root)
        self.root.update_idletasks()
        self._resize_origin = (
            self.root.winfo_x(), self.root.winfo_y(),
            self.root.winfo_width(), self.root.winfo_height(),
        )
        self._resize_start = (event.x_root, event.y_root)

    def _resize_drag(self, event, edge):
        if self._resize_origin is None or self._resize_start is None:
            return
        dx = event.x_root - self._resize_start[0]
        dy = event.y_root - self._resize_start[1]
        x, y, width, height = self._resize_origin
        _area_x, _area_y, area_w, area_h = ui.workarea()
        max_w = max(area_w, self._min_width)
        max_h = max(area_h, self._min_height)

        if "e" in edge:
            width = min(max(width + dx, self._min_width), max_w)
        if "s" in edge:
            height = min(max(height + dy, self._min_height), max_h)
        if "w" in edge:
            right = x + width
            x = min(max(x + dx, right - max_w), right - self._min_width)
            width = right - x
        if "n" in edge:
            bottom = y + height
            y = min(max(y + dy, bottom - max_h), bottom - self._min_height)
            height = bottom - y
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    # ------------------------------------------------------------------ #
    #  Настройки между запусками
    # ------------------------------------------------------------------ #
    def _window_settings(self) -> dict:
        try:
            geometry = self._restore_geometry if self._zoomed else self.root.geometry()
            width, height, x, y = self._parse_geometry(geometry)
        except (ValueError, AttributeError):
            width = int(BASE_WIDTH * self.scale)
            height = int(BASE_HEIGHT * self.scale)
            x = y = 0
        return {"width": width, "height": height, "x": x, "y": y, "maximized": bool(self._zoomed)}

    def close(self):
        dirty = [self._form_label(key) for key in self._dirty_pages()]
        if dirty and not ui.confirm(
            self.root, "Несохранённые данные",
            "Есть несохранённые изменения:\n"
            + "\n".join(f"• {label}" for label in dirty)
            + "\n\nЗакрыть программу и потерять их?",
            ok_text="Закрыть без сохранения", cancel_text="Остаться",
        ):
            return
        saved = write_settings({
            **self._settings,  # сохраняем и остальные настройки (папка документов и т. п.)
            "theme": str(ctk.get_appearance_mode()).lower(),
            "window": self._window_settings(),
        })
        if not saved:
            ui.alert(
                self.root, "Настройки не сохранены",
                f"Не удалось записать файл:\n{settings_path()}", "error",
            )
        self._log.info("Приложение закрыто пользователем")
        try:
            self.root.destroy()
        except Exception:
            pass

    def _cleanup_temp_documents(self) -> None:
        """Удаляет *.docx.tmp, оставшиеся после аварийного завершения записи."""
        try:
            base = self._documents_dir(interactive=False)
            directories = [base]
            with os.scandir(base) as entries:
                directories.extend(
                    entry.path for entry in entries if entry.is_dir()
                )
        except OSError:
            return
        removed = 0
        for directory in directories:
            try:
                with os.scandir(directory) as entries:
                    for entry in entries:
                        if entry.is_file() and entry.name.endswith(".docx.tmp"):
                            try:
                                os.remove(entry.path)
                                removed += 1
                            except OSError:
                                pass
            except OSError:
                continue
        if removed:
            self._log.info("Удалено незавершённых .docx.tmp: %d", removed)

    def _cleanup_expired_trash(self) -> None:
        """Физически удаляет записи, пролежавшие в корзине дольше срока хранения."""
        try:
            removed = self.store.cleanup_trash()
        except Exception as exc:  # noqa: BLE001 — фоновая задача не должна падать
            self._log.warning("Автоочистка корзины не удалась: %s", type(exc).__name__)
            return
        if removed:
            self.refresh_registry()
            self._update_status()

    # ------------------------------------------------------------------ #
    #  Диалоги с произвольным набором кнопок
    # ------------------------------------------------------------------ #
    def _choice_dialog(self, title: str, message: str,
                       options: list[tuple[str, str, bool]]) -> str | None:
        """Модальный диалог с N кнопками.

        ``options``: список ``(ключ, подпись, главная_кнопка)`` слева направо.
        Возвращает ключ выбранного варианта или ``None`` при закрытии окна.
        """
        result: dict[str, str | None] = {"key": None}
        win = ctk.CTkToplevel(self.root)
        win.title(title)
        win.configure(fg_color=ui.COLORS["window"])
        win.transient(self.root)
        win.resizable(False, False)

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=int(24 * self.scale),
                  pady=int(20 * self.scale))
        ctk.CTkLabel(
            body, text=title, font=ui.font(16, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", justify="left",
            wraplength=int(460 * self.scale),
        ).pack(anchor="w")
        ctk.CTkLabel(
            body, text=message, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left",
            wraplength=int(460 * self.scale),
        ).pack(anchor="w", pady=(int(10 * self.scale), int(16 * self.scale)))

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill="x")

        def choose(key: str | None) -> None:
            result["key"] = key
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                win.destroy()
            except Exception:
                pass

        for key, label, primary in options:
            if primary:
                button = ui.primary_button(
                    buttons, label, lambda k=key: choose(k),
                    width=int(190 * self.scale), height=int(32 * self.scale),
                )
            else:
                button = ui.ghost_button(
                    buttons, label, lambda k=key: choose(k),
                    width=int(160 * self.scale), height=int(32 * self.scale),
                )
            button.pack(side="left", padx=(0, int(8 * self.scale)))

        win.protocol("WM_DELETE_WINDOW", lambda: choose(None))
        win.update_idletasks()
        try:
            x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
            y = self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_height()) // 3
            win.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass
        try:
            win.grab_set()
        except Exception:
            pass
        win.wait_window()
        return result["key"]

    def _persist_settings(self) -> bool:
        """Сразу сохраняет настройки (тема, папка документов) на диск."""
        saved = write_settings({
            **self._settings,
            "theme": str(ctk.get_appearance_mode()).lower(),
        })
        if not saved:
            ui.alert(
                self.root, "Настройки не сохранены",
                f"Не удалось записать файл:\n{settings_path()}", "error",
            )
        return saved

    def _on_theme_change(self, mode: str):
        """Переключение темы со страницы «Настройки» (mode: "light"/"dark")."""
        ui.set_appearance_mode(mode)
        self._settings["theme"] = mode
        self._persist_settings()

    # ------------------------------------------------------------------ #
    #  Каркас: сайдбар, контент, футер
    # ------------------------------------------------------------------ #
    def _build_body(self):
        self.body = ctk.CTkFrame(self.root, fg_color=ui.COLORS["content"], corner_radius=0)
        self.body.pack(fill="both", expand=True)

        self.content = ctk.CTkFrame(self.body, fg_color=ui.COLORS["content"], corner_radius=0)
        self.content.pack(fill="both", expand=True)
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.footer = ctk.CTkFrame(self.body, fg_color=ui.COLORS["window"], corner_radius=0, height=int(62 * self.scale))
        self.footer.pack(fill="x", side="bottom")
        self.footer.pack_propagate(False)

        self._build_footer()
        # Немодальное уведомление об успехе (п. 4.5): не блокирует работу и
        # предлагает открыть файл или показать его в папке.
        self.toast = ctk.CTkFrame(
            self.body, fg_color=ui.COLORS["card"], corner_radius=10,
            border_width=1, border_color=ui.COLORS["card_border"],
        )
        self.toast_label = ctk.CTkLabel(
            self.toast, text="", font=ui.font(12),
            text_color=ui.COLORS["text"], anchor="w", justify="left",
            wraplength=int(420 * self.scale),
        )
        self.toast_label.pack(side="left", padx=int(14 * self.scale), pady=int(10 * self.scale))
        self.toast_open = ui.ghost_button(
            self.toast, "Открыть файл", self._toast_open,
            width=int(130 * self.scale), height=int(26 * self.scale),
        )
        self.toast_folder = ui.ghost_button(
            self.toast, "Показать в папке", self._toast_show_folder,
            width=int(160 * self.scale), height=int(26 * self.scale),
        )

    # -- Немодальные уведомления --------------------------------------- #
    def _notify(self, message: str, path: str | None = None) -> None:
        if not hasattr(self, "toast"):
            return
        self._toast_path = path or None
        try:
            self.toast_label.configure(text=message)
        except Exception:
            return
        if self._toast_path:
            self.toast_open.pack(side="right", padx=(int(8 * self.scale), int(10 * self.scale)),
                                 pady=int(8 * self.scale))
            self.toast_folder.pack(side="right", padx=(0, int(2 * self.scale)),
                                   pady=int(8 * self.scale))
        else:
            self.toast_open.pack_forget()
            self.toast_folder.pack_forget()
        try:
            self.toast.place(
                relx=1.0, rely=1.0, anchor="se",
                x=-int(20 * self.scale), y=-int(76 * self.scale),
            )
            self.toast.lift()
        except Exception:
            return
        if self._toast_after:
            try:
                self.root.after_cancel(self._toast_after)
            except Exception:
                pass
        self._toast_after = self.root.after(9000, self._hide_toast)

    def _hide_toast(self) -> None:
        self._toast_after = None
        try:
            self.toast.place_forget()
        except Exception:
            pass

    def _toast_open(self) -> None:
        if self._toast_path:
            open_file(self._toast_path)

    def _toast_show_folder(self) -> None:
        if self._toast_path:
            show_in_folder(self._toast_path)

    def _build_footer(self):
        ctk.CTkFrame(self.footer, fg_color=ui.COLORS["hairline"], corner_radius=0, height=1).pack(
            fill="x", side="top"
        )

        inner = ctk.CTkFrame(self.footer, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=int(20 * self.scale))

        self.save_button = ui.primary_button(
            inner, "Сохранить договор", self.save_contract,
            width=int(186 * self.scale), height=int(32 * self.scale),
        )
        self.save_button.pack(side="right", pady=int(15 * self.scale))

        self.status_label = ctk.CTkLabel(
            inner, text="", font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w",
        )
        self.status_label.pack(side="left", pady=int(16 * self.scale))
        self._update_status()

    def _update_footer(self):
        """Кнопка футера и её подпись зависят от открытой страницы."""
        commands = {
            "counterparty": ("Сохранить контрагента", self._cp_save),
            "contract": ("Сохранить договор", self.save_contract),
            "additional": ("Сохранить доп. соглашение", self.save_additional),
            "spec": ("Сохранить спецификацию", self.save_spec),
        }
        if self.current_page in commands:
            text, command = commands[self.current_page]
            self.save_button.configure(text=text, command=command)
            self.save_button.pack(side="right", pady=int(15 * self.scale))
        else:
            self.save_button.pack_forget()
        self._update_status()

    # ------------------------------------------------------------------ #
    #  Страницы: общий каркас «параметры слева — документ справа»
    # ------------------------------------------------------------------ #
    def _build_pages(self):
        self._build_contract_page()
        self._build_additional_page()
        self._build_spec_page()
        self._build_counterparty_page()
        self._build_registry_page()
        self._build_settings_page()
        # Обновляем списки только после того, как все страницы созданы:
        # иначе в «Доп. соглашении» и «Спецификации» не будет контрагентов.
        self._refresh_party_combo_values()

    def _make_page(self, key: str, title: str, subtitle: str, with_text: bool = False,
                   text_title: str = "Текст документа (только чтение)"):
        page = ctk.CTkFrame(self.content, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        self.pages[key] = page

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * self.scale), pady=(int(18 * self.scale), int(6 * self.scale)))
        ctk.CTkLabel(
            header, text=title, font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text=subtitle, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * self.scale), 0))

        # Общая панель «Документ для: …» — выбор контрагента переносится
        # между страницами договора/доп. соглашения/спецификации (п. 4.2).
        if with_text:
            party_bar = ctk.CTkFrame(
                page, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8,
            )
            party_bar.pack(
                fill="x", padx=int(24 * self.scale),
                pady=(int(4 * self.scale), int(4 * self.scale)),
            )
            party_var = tk.StringVar(value="Документ для: контрагент не выбран")
            if hasattr(self, "_current_party_vars"):
                self._current_party_vars[key] = party_var
            ctk.CTkLabel(
                party_bar, textvariable=party_var, font=ui.font(12, "bold"),
                text_color=ui.COLORS["text"], anchor="w", justify="left",
            ).pack(side="left", padx=int(12 * self.scale), pady=int(6 * self.scale))
            ui.ghost_button(
                party_bar, "Выбрать в реестре", lambda k=key: self._select_page("registry"),
                width=int(150 * self.scale), height=int(26 * self.scale),
            ).pack(side="right", padx=int(10 * self.scale), pady=int(4 * self.scale))

        content = ctk.CTkFrame(page, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=int(24 * self.scale), pady=(0, int(6 * self.scale)))
        content.grid_columnconfigure(0, weight=0 if with_text else 1)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        left = ctk.CTkScrollableFrame(
            content, width=int(LEFT_PANEL_WIDTH * self.scale),
            fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        left.grid(row=0, column=0, columnspan=1 if with_text else 2, sticky="nsew",
                  padx=(0, 0 if with_text else int(2 * self.scale)))
        left.grid_columnconfigure(0, weight=1)
        self.scrollers.append(left)

        textbox = None
        if with_text:
            right = ctk.CTkFrame(content, fg_color=ui.COLORS["card"], corner_radius=12,
                                 border_width=1, border_color=ui.COLORS["card_border"])
            right.grid(row=0, column=1, sticky="nsew")
            right.grid_rowconfigure(1, weight=1)
            right.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                right, text=text_title,
                font=ui.font(12, "bold"), text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew", padx=int(16 * self.scale), pady=(int(10 * self.scale), 0))
            textbox = ctk.CTkTextbox(
                right, wrap="word", font=ui.font(12),
                fg_color=ui.COLORS["card"], text_color=ui.COLORS["text"],
            )
            textbox.grid(row=1, column=0, sticky="nsew", padx=int(12 * self.scale), pady=(int(4 * self.scale), int(12 * self.scale)))
            # Предпросмотр — только для чтения: правки не влияют на .docx,
            # поэтому интерфейс не должен обещать обратное (п. 4.3).
            try:
                textbox.configure(state="disabled")
            except Exception:
                pass
        else:
            right = None

        if textbox is not None:
            # Ctrl+C/V/X в предпросмотре должны работать на русской раскладке.
            try:
                ui.fix_clipboard_shortcuts(textbox._textbox)
            except Exception:
                pass

        return page, left, right, textbox

    # ------------------------------------------------------------------ #
    #  Страница «Договор поставки»
    # ------------------------------------------------------------------ #
    def _build_contract_page(self):
        _page, left, _right, textbox = self._make_page(
            "contract", "Договор поставки",
            "Выберите контрагента, укажите номер и дату. «Сохранить договор» запишет его "
            "в реестр и сразу откроет готовый документ в Word.",
            with_text=True, text_title="Предпросмотр договора (только чтение)",
        )
        self.contract_preview = textbox

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Контрагент")

        self.contract_party_var = tk.StringVar(value="")
        self.contract_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(ENTRY_WIDTH * self.scale), variable=self.contract_party_var,
        )
        card.wide(self.contract_party_combo, pady=(int(2 * self.scale), int(10 * self.scale)))

        search_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        self.contract_search_var = tk.StringVar()
        search_entry = ui.entry(
            search_holder, width=int(180 * self.scale),
            textvariable=self.contract_search_var,
            validatecommand=self._digits_vcmd_for(12),
        )
        ui.enable_sanitized_paste(search_entry.entry, digits_only, max_len=12)
        search_entry.pack(side="left")
        try:  # Enter в поиске ищет, а не сохраняет договор
            search_entry.bind_entry("<Return>", lambda _e: (self._find_party_by_inn(), "break")[1])
            search_entry.bind_entry("<KP_Enter>", lambda _e: (self._find_party_by_inn(), "break")[1])
        except Exception:
            pass
        ui.ghost_button(
            search_holder, "Найти", self._find_party_by_inn,
            width=int(92 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(10 * self.scale), 0))
        card.row(
            "Поиск по ИНН", search_holder, stretch=True,
            hint="Найдёт контрагента в реестре программы.",
        )

        # -- Договор ----------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Договор")

        self.contract_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(SHORT_WIDTH * self.scale),
            textvariable=self.contract_number_var,
        )
        self._page_widgets[("contract", F.NUMBER)] = number
        card.row(F.NUMBER, number, required=True,
                 error_var=self._error_var_for("contract", F.NUMBER))
        self.contract_number_var.trace_add("write", lambda *_: self._on_doc_field_changed("contract", F.NUMBER))

        self.contract_day_var = tk.StringVar()
        self.contract_month_var = tk.StringVar(value=MONTHS_GENITIVE[datetime.now().month - 1])
        self._add_date_row(card, "contract", self.contract_day_var, self.contract_month_var)
        self.contract_day_var.trace_add("write", lambda *_: self._on_doc_field_changed("contract", F.DAY))
        self.contract_month_var.trace_add("write", lambda *_: self._on_doc_field_changed("contract", F.MONTH))

        self._refresh_party_combo_values()

        # Предпросмотр обновляется при любом изменении полей договора.
        self.contract_party_var.trace_add("write", self._preview_contract)
        self.contract_party_var.trace_add("write", lambda *_: self._on_party_selected("contract"))
        self.contract_number_var.trace_add("write", self._preview_contract)
        self.contract_day_var.trace_add("write", self._preview_contract)
        self.contract_month_var.trace_add("write", self._preview_contract)
        self._preview_contract()

    def _preview_contract(self, *_args):
        """Откладывает перерисовку предпросмотра, чтобы не собирать .docx на каждый ввод."""
        if getattr(self, "_preview_contract_after", None):
            try:
                self.root.after_cancel(self._preview_contract_after)
            except Exception:
                pass
        self._preview_contract_after = self.root.after(120, self._render_contract_preview)

    def _render_contract_preview(self):
        self._preview_contract_after = None
        rec = self._party_from_var(self.contract_party_var)
        if rec is None:
            self._set_text(
                self.contract_preview,
                "Выберите контрагента и заполните поля слева — здесь появится текст договора.",
            )
            return
        data = dict(rec.get("fields", {}))
        data[F.NUMBER] = self.contract_number_var.get().strip()
        data[F.DAY] = self.contract_day_var.get().strip()
        data[F.MONTH] = self.contract_month_var.get().strip()
        is_ip = rec.get("type") == "ip"
        try:
            doc = self._render_contract_docx(data, is_ip)
            self._set_text(self.contract_preview, docx_to_text(doc))
        except Exception as exc:
            self._set_text(self.contract_preview, f"Не удалось собрать предпросмотр: {exc}")

    def _find_party_by_inn(self):
        """Ищет контрагента по ИНН в реестре и подставляет в выпадающий список."""
        inn = digits_only(self.contract_search_var.get())
        rec = self.store.find_by_inn(inn)
        if rec is None:
            ui.alert(
                self.root, "Контрагент не найден",
                f"В реестре программы нет контрагента с ИНН {inn or '—'}.\n"
                "Добавьте его через кнопку «Добавить контрагента».",
                "error",
            )
            return
        self._set_current_party(rec)
        self.contract_search_var.set("")

    # ------------------------------------------------------------------ #
    #  Страница «Доп. соглашение»
    # ------------------------------------------------------------------ #
    def _build_additional_page(self):
        _page, left, _right, textbox = self._make_page(
            "additional", "Доп. соглашение",
            "Выберите контрагента и договор-основание. «Сохранить» запишет документ "
            "в реестр и сразу откроет его в Word.",
            with_text=True, text_title="Предпросмотр доп. соглашения (только чтение)",
        )
        self.addl_preview = textbox

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Контрагент")

        self.addl_party_var = tk.StringVar(value="")
        self.addl_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(ENTRY_WIDTH * self.scale), variable=self.addl_party_var,
        )
        card.wide(self.addl_party_combo, pady=(int(2 * self.scale), int(10 * self.scale)))
        self.addl_party_var.trace_add("write", self._on_addl_party_changed)
        self.addl_party_var.trace_add("write", lambda *_: self._on_party_selected("additional"))

        # -- Договор-основание -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Договор-основание")

        self.addl_contract_var = tk.StringVar(value="")
        self.addl_contract_combo = ui.option_menu(
            card.body, ["— нет сохранённых договоров —"],
            width=int(ENTRY_WIDTH * self.scale), variable=self.addl_contract_var,
        )
        card.wide(self.addl_contract_combo, pady=(int(2 * self.scale), int(4 * self.scale)))
        self.addl_contract_info_var = tk.StringVar(value="Договор-основание не выбран.")
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=self.addl_contract_info_var,
                font=ui.font(11), text_color=ui.COLORS["text_secondary"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 40) * self.scale),
            ),
            pady=(int(2 * self.scale), int(2 * self.scale)),
        )
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Договоры берутся из реестра: сначала создайте договор поставки.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left", wraplength=int((LEFT_PANEL_WIDTH - 40) * self.scale),
            ),
            pady=(int(4 * self.scale), int(6 * self.scale)),
        )

        # -- Условия ------------------------------------------------------ #
        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Условия")

        self.addl_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(SHORT_WIDTH * self.scale),
            textvariable=self.addl_number_var,
        )
        card.row("Номер доп. соглашения", number,
                 hint="Можно оставить пустым — сохранится как «б/н».")

        self._limit_var = tk.StringVar()
        self._limit_var.trace_add("write", self._format_limit)
        limit_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        self.limit_entry = ui.entry(
            limit_holder, width=int(220 * self.scale), textvariable=self._limit_var,
            validatecommand=(self._limit_vcmd, "%P"),
        )
        self.limit_entry.pack(side="left")
        ctk.CTkLabel(
            limit_holder, text="руб.", font=ui.font(12), text_color=ui.COLORS["text_tertiary"],
        ).pack(side="left", padx=(int(8 * self.scale), 0))
        self._page_widgets[("additional", F.LIMIT)] = self.limit_entry
        card.row(
            F.LIMIT, limit_holder, stretch=True, required=True,
            hint="Сумма по доп. соглашению к договору поставки.",
            error_var=self._error_var_for("additional", F.LIMIT),
        )

        self.limit_words_label = ctk.CTkLabel(
            card.body, text="Сумма прописью: —", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left", height=18,
        )
        card.wide(self.limit_words_label, pady=(int(4 * self.scale), int(8 * self.scale)))

        term_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        self.term_var = tk.StringVar()
        self.term_entry = ui.entry(
            term_holder, width=int(120 * self.scale), textvariable=self.term_var,
            validatecommand=(self._term_vcmd, "%P"),
        )
        self.term_entry.pack(side="left")
        ctk.CTkLabel(
            term_holder, text="дней", font=ui.font(12), text_color=ui.COLORS["text_tertiary"],
        ).pack(side="left", padx=(int(8 * self.scale), 0))
        self._page_widgets[("additional", F.TERM)] = self.term_entry
        card.row(
            F.TERM, term_holder, required=True,
            hint="Отсрочка платежа по доп. соглашению.",
            error_var=self._error_var_for("additional", F.TERM),
        )

        self.words_label = ctk.CTkLabel(
            card.body, text="Срок прописью: —", font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left", height=18,
        )
        card.wide(self.words_label, pady=(int(2 * self.scale), int(2 * self.scale)))

        self.term_var.trace_add("write", self._update_words)
        self._update_words()
        self._update_limit_words()

        # Предпросмотр обновляется при любом изменении данных доп. соглашения.
        self.addl_party_var.trace_add("write", self._preview_additional)
        self.addl_contract_var.trace_add("write", self._preview_additional)
        self.addl_contract_var.trace_add("write", lambda *_: self._update_contract_info("additional"))
        self.addl_number_var.trace_add("write", self._preview_additional)
        self._limit_var.trace_add("write", self._preview_additional)
        self.term_var.trace_add("write", self._preview_additional)
        self._preview_additional()

    def _preview_additional(self, *_args):
        """Откладывает перерисовку предпросмотра, чтобы не собирать .docx на каждый ввод."""
        if getattr(self, "_preview_additional_after", None):
            try:
                self.root.after_cancel(self._preview_additional_after)
            except Exception:
                pass
        self._preview_additional_after = self.root.after(120, self._render_additional_preview)

    def _render_additional_preview(self):
        self._preview_additional_after = None
        party = self._party_from_var(self.addl_party_var)
        contract = self._contract_from_var(self.addl_contract_var, party["id"]) if party else None
        if party is None or contract is None:
            self._set_text(
                self.addl_preview,
                "Выберите контрагента и договор-основание слева — здесь появится текст документа.",
            )
            return
        data = dict(contract.get("data") or {})
        is_ip = party.get("type") == "ip"
        try:
            doc = self._render_additional_docx(
                data, is_ip,
                self.addl_number_var.get().strip(),
                self._limit_var.get().strip(),
                self.term_var.get().strip(),
            )
            self._set_text(self.addl_preview, docx_to_text(doc))
        except Exception as exc:
            self._set_text(self.addl_preview, f"Не удалось собрать предпросмотр: {exc}")

    def _on_addl_party_changed(self, *_args):
        """Обновляет список договоров выбранного контрагента."""
        rec = self._party_from_var(self.addl_party_var)
        contracts = self.store.contracts_for(rec["id"]) if rec else []
        if contracts:
            values = [self._contract_display(c) for c in contracts]
        else:
            values = ["— нет сохранённых договоров —"]
        self.addl_contract_combo.set_values(values)
        self.addl_contract_var.set(values[0] if contracts else "")
        # Явно сообщаем, что договор подставлен автоматически (п. 4.2).
        if contracts:
            self.addl_contract_info_var.set(
                f"Договор-основание: {values[0]} — подставлен автоматически, проверьте."
            )
        else:
            self.addl_contract_info_var.set(
                "Договор-основание не выбран: сохраните договор поставки."
            )

    # ------------------------------------------------------------------ #
    #  Страница «Спецификация»
    # ------------------------------------------------------------------ #
    def _build_spec_page(self):
        _page, left, right, textbox = self._make_page(
            "spec", "Спецификация",
            "Позиции поставки к договору; справа — таблица позиций и текст документа.",
            with_text=True, text_title="Предпросмотр спецификации (только чтение)",
        )
        self.spec_text = textbox

        # -- Правая область: таблица позиций + предпросмотр ---------------- #
        # Заголовок правой области убираем — в табличной части только строки.
        top_label = right.grid_slaves(row=0, column=0)
        if top_label:
            top_label[0].grid_remove()
        textbox.grid_forget()
        right.grid_rowconfigure(1, weight=0)
        right.grid_rowconfigure(2, weight=1)

        self.spec_table = SpecTable(right, scale=self.scale)
        self.spec_table.grid(row=1, column=0, sticky="ew",
                             padx=int(12 * self.scale), pady=(int(4 * self.scale), 0))

        textbox.grid(row=2, column=0, sticky="nsew",
                     padx=int(12 * self.scale), pady=(int(4 * self.scale), int(12 * self.scale)))
        self.spec_table.on_change = self._on_spec_item_changed

        # -- Контрагент -------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Контрагент")

        self.spec_party_var = tk.StringVar(value="")
        self.spec_party_combo = ui.option_menu(
            card.body, ["— выберите контрагента —"],
            width=int(ENTRY_WIDTH * self.scale), variable=self.spec_party_var,
        )
        card.wide(self.spec_party_combo, pady=(int(2 * self.scale), int(10 * self.scale)))
        self.spec_party_var.trace_add("write", self._on_spec_party_changed)
        self.spec_party_var.trace_add("write", lambda *_: self._on_party_selected("spec"))

        # -- Договор-основание -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Договор-основание")

        self.spec_contract_var = tk.StringVar(value="")
        self.spec_contract_combo = ui.option_menu(
            card.body, ["— нет сохранённых договоров —"],
            width=int(ENTRY_WIDTH * self.scale), variable=self.spec_contract_var,
        )
        card.wide(self.spec_contract_combo, pady=(int(2 * self.scale), int(4 * self.scale)))
        self.spec_contract_info_var = tk.StringVar(value="Договор-основание не выбран.")
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=self.spec_contract_info_var,
                font=ui.font(11), text_color=ui.COLORS["text_secondary"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 40) * self.scale),
            ),
            pady=(int(2 * self.scale), int(2 * self.scale)),
        )
        self.spec_contract_var.trace_add("write", self._regenerate_spec)
        self.spec_contract_var.trace_add("write", lambda *_: self._update_contract_info("spec"))
        self.spec_contract_var.trace_add("write", lambda *_: self._on_doc_field_changed("spec", F.NUMBER))

        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Спецификация")

        self.spec_number_var = tk.StringVar()
        number = ui.entry(
            card.body, width=int(SHORT_WIDTH * self.scale),
            textvariable=self.spec_number_var,
        )
        self._page_widgets[("spec", F.NUMBER)] = number
        card.row("Номер спецификации", number, required=True,
                 error_var=self._error_var_for("spec", F.NUMBER))
        self.spec_number_var.trace_add("write", lambda *_: self._on_doc_field_changed("spec", F.NUMBER))

        self.spec_day_var = tk.StringVar()
        self.spec_month_var = tk.StringVar(value=MONTHS_GENITIVE[datetime.now().month - 1])
        self._add_date_row(card, "spec", self.spec_day_var, self.spec_month_var)
        self.spec_day_var.trace_add("write", self._regenerate_spec)
        self.spec_month_var.trace_add("write", self._regenerate_spec)

        # -- Условия поставки -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=3, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Условия поставки")

        self.spec_delivery_term_var = tk.StringVar()
        delivery_term = ui.entry(
            card.body, width=int(ENTRY_WIDTH * self.scale),
            textvariable=self.spec_delivery_term_var,
        )
        self._page_widgets[("spec", F.DELIVERY_TERM)] = delivery_term
        card.row(F.DELIVERY_TERM, delivery_term,
                 hint="Например: 30 дней с момента подписания договора.")

        self.spec_payment_terms_var = tk.StringVar()
        payment_terms = ui.entry(
            card.body, width=int(ENTRY_WIDTH * self.scale),
            textvariable=self.spec_payment_terms_var,
        )
        self._page_widgets[("spec", F.PAYMENT_TERMS)] = payment_terms
        card.row(F.PAYMENT_TERMS, payment_terms,
                 hint="Например: оплата в течение 10 банковских дней после отгрузки.")

        self.spec_delivery_basis_var = tk.StringVar()
        delivery_basis = ui.entry(
            card.body, width=int(ENTRY_WIDTH * self.scale),
            textvariable=self.spec_delivery_basis_var,
        )
        self._page_widgets[("spec", F.DELIVERY_BASIS)] = delivery_basis
        card.row(F.DELIVERY_BASIS, delivery_basis,
                 hint="Например: самовывоз со склада поставщика.")

        self.spec_prices_until_var = tk.StringVar(value=prices_valid_until_default())
        prices_until = ui.entry(
            card.body, width=int(ENTRY_WIDTH * self.scale),
            textvariable=self.spec_prices_until_var,
        )
        card.row("Цены действительны до", prices_until,
                 hint="Автоматически — последний день текущего месяца")

        self.spec_delivery_term_var.trace_add("write", self._regenerate_spec)
        self.spec_payment_terms_var.trace_add("write", self._regenerate_spec)
        self.spec_delivery_basis_var.trace_add("write", self._regenerate_spec)

        # -- Позиции (подсказка) ------------------------------------------- #

    @property
    def _spec_rows(self) -> list[dict]:
        """Строки позиций (StringVar «name», «qty», «unit», «price»)."""
        return self.spec_table.rows

    def _add_spec_row(self, name="", qty="", price="", unit=""):
        """Совместимость: добавление строки в таблицу позиций."""
        return self.spec_table.add_row(
            {"name": name, "qty": qty, "unit": unit, "price": price}
        )

    def _remove_spec_row(self):
        self.spec_table.remove_last()

    def _spec_items(self) -> list[dict]:
        """Позиции с заполненным наименованием; суммы считаются на лету."""
        return self.spec_table.get_items()

    def _spec_total(self) -> Decimal:
        return sum(
            (item["qty"] * item["price"] for item in self._spec_items()),
            Decimal("0"),
        )

    def _on_spec_item_changed(self, *_args):
        self._regenerate_spec()

    def _on_spec_party_changed(self, *_args):
        rec = self._party_from_var(self.spec_party_var)
        contracts = self.store.contracts_for(rec["id"]) if rec else []
        if contracts:
            values = [self._contract_display(c) for c in contracts]
        else:
            values = ["— нет сохранённых договоров —"]
        self.spec_contract_combo.set_values(values)
        self.spec_contract_var.set(values[0] if contracts else "")
        if contracts:
            self.spec_contract_info_var.set(
                f"Договор-основание: {values[0]} — подставлен автоматически, проверьте."
            )
        else:
            self.spec_contract_info_var.set(
                "Договор-основание не выбран: сохраните договор поставки."
            )
        self._regenerate_spec()

    # ------------------------------------------------------------------ #
    #  Страница «Реестр»
    # ------------------------------------------------------------------ #
    def _build_registry_page(self):
        page = ctk.CTkFrame(self.content, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        self.pages["registry"] = page

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * self.scale), pady=(int(18 * self.scale), int(6 * self.scale)))
        title_row = ctk.CTkFrame(header, fg_color="transparent")
        title_row.pack(fill="x")
        ctk.CTkLabel(
            title_row, text="Реестр", font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        ).pack(side="left")
        ui.ghost_button(
            title_row, "Обновить", self.refresh_registry,
            width=int(110 * self.scale), height=int(28 * self.scale),
        ).pack(side="right")
        ui.ghost_button(
            title_row, "Добавить документы…", self._import_documents_dialog,
            width=int(180 * self.scale), height=int(28 * self.scale),
        ).pack(side="right", padx=(0, int(8 * self.scale)))
        ctk.CTkLabel(
            header, text="Список контрагентов. Выберите контрагента — ниже появятся его документы.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * self.scale), 0))

        search_bar = ctk.CTkFrame(page, fg_color="transparent")
        search_bar.pack(fill="x", padx=int(24 * self.scale), pady=(0, int(8 * self.scale)))
        self.registry_search_var = tk.StringVar()
        # Поиск идёт и по названию, и по ИНН, с фильтрацией по мере ввода (п. 4.5).
        self.registry_search_var.trace_add("write", self._on_registry_search_changed)
        search_entry = ui.entry(
            search_bar, width=int(280 * self.scale),
            textvariable=self.registry_search_var,
        )
        search_entry.pack(side="left")
        ui.ghost_button(
            search_bar, "Найти", self._registry_apply_filter,
            width=int(92 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(10 * self.scale), 0))
        ui.ghost_button(
            search_bar, "Показать всех", self._registry_reset_filter,
            width=int(130 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(8 * self.scale), 0))

        self.registry_list = ctk.CTkScrollableFrame(
            page, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        self.registry_list.pack(fill="both", expand=True, padx=int(24 * self.scale), pady=(0, int(10 * self.scale)))
        self.registry_list.grid_columnconfigure(0, weight=1)
        self.scrollers.append(self.registry_list)

        self._registry_selected_id = None
        self._registry_rows: dict = {}
        self._registry_docs_row = 1
        self._registry_docs_frame = None
        # Ленивый рендер (п. 9 этапа 5): при старте список не рисуется —
        # реестр отрисуется при первом открытии страницы или при изменениях.
        self._registry_rendered = False

    # -- Реестр: список контрагентов и документы выбранного -------------- #
    def _registry_parties(self) -> list[dict]:
        """Активные контрагенты с учётом поиска по названию и ИНН (подстрока)."""
        query = self.registry_search_var.get().strip().lower()
        parties = self.store.active_counterparties()
        if not query:
            return parties
        digits = digits_only(query)

        def matches(rec: dict) -> bool:
            fields = rec.get("fields") or {}
            inn = str(fields.get(F.INN, "")).strip()
            if digits and digits in inn:
                return True
            name_parts = (
                storage.display_name(rec),
                str(fields.get(F.FULL_NAME, "")),
                str(fields.get(F.IP_NAME, "")),
                str(fields.get(F.SHORT_NAME, "")),
            )
            haystack = " ".join(name_parts).lower()
            return bool(query) and query in haystack

        return [rec for rec in parties if matches(rec)]

    def _on_registry_search_changed(self, *_args):
        """Фильтрует реестр по мере ввода с небольшой задержкой."""
        pending = getattr(self, "_registry_search_after", None)
        if pending:
            try:
                self.root.after_cancel(pending)
            except Exception:
                pass
        self._registry_search_after = self.root.after(200, self.refresh_registry)

    def _registry_apply_filter(self):
        pending = getattr(self, "_registry_search_after", None)
        if pending:
            try:
                self.root.after_cancel(pending)
            except Exception:
                pass
            self._registry_search_after = None
        self.refresh_registry()

    def _registry_reset_filter(self):
        self.registry_search_var.set("")
        self._registry_apply_filter()

    def _select_registry_party(self, rec: dict):
        """Клик по контрагенту: показать/скрыть его документы.

        Частичное обновление (п. 5.1): строки контрагентов не пересоздаются,
        перерисовываются только подсветка выбора и секция документов.
        """
        if self._registry_selected_id == rec.get("id"):
            self._registry_selected_id = None
        else:
            self._registry_selected_id = rec.get("id")
            # Выбор из реестра переносится на страницы документов (п. 4.2).
            self._set_current_party(rec)
        if not self._registry_rows:
            # Страница ещё не отрисована (ленивый рендер) — рисуем целиком.
            self.refresh_registry()
            return
        self._refresh_registry_selection()

    @staticmethod
    def _party_row_text(rec: dict, selected: bool) -> str:
        """Текст строки контрагента в реестре (название, тип, ИНН, ID, адрес)."""
        kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
        inn = str(rec.get("fields", {}).get(F.INN, "")).strip() or "—"
        address = str(rec.get("fields", {}).get(F.LEGAL_ADDRESS, "")).strip()
        if len(address) > 48:
            address = address[:47].rstrip() + "…"
        text = f"{storage.display_name(rec)}   ·   {kind}   ·   ИНН {inn}   ·   ID {rec.get('id')}"
        if address:
            text += f"   ·   {address}"
        if selected:
            text = "●  " + text
        return text

    def _refresh_registry_selection(self):
        """Частичное обновление реестра при смене выбора (п. 5.1).

        Строки контрагентов остаются теми же виджетами (меняется только
        подсветка), секция документов выбранного пересобирается на своём
        месте. Полная перерисовка нужна только при изменении состава списка
        (поиск, добавление, удаление).
        """
        for party_id, entry in self._registry_rows.items():
            selected = self._registry_selected_id == party_id
            btn = entry["btn"]
            try:
                btn.configure(
                    text=self._party_row_text(entry["rec"], selected),
                    font=ui.font(12, "bold" if selected else "normal"),
                    text_color=(ui.COLORS["text"] if selected
                                else ui.COLORS["text_secondary"]),
                )
            except Exception:
                pass
        if self._registry_docs_frame is not None:
            try:
                self._registry_docs_frame.destroy()
            except Exception:
                pass
            self._registry_docs_frame = None
        self._render_registry_documents(self._registry_docs_row)
        self.root.after(60, self._refresh_scrollbars)

    def refresh_registry(self):
        """Перерисовывает реестр целиком (поиск, изменения состава).

        Длительные перерисовки (более 500 мс на больших объёмах) попадают
        в журнал (п. 10 этапа 5); клик по контрагенту перерисовывает только
        подсветку и секцию документов (_refresh_registry_selection, п. 5.1).
        """
        with applog.timed("Перерисовка реестра", 500, "ui"):
            self._refresh_registry_full()
        self._registry_rendered = True

    def _refresh_registry_full(self):
        for child in self.registry_list.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        self._registry_docs_frame = None

        parties = self._registry_parties()
        query = self.registry_search_var.get().strip()
        caption = (
            f"Найдено по запросу «{query}»: {len(parties)}"
            if query else f"Контрагенты ({len(parties)})"
        )
        ctk.CTkLabel(
            self.registry_list, text=caption,
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(6 * self.scale)))

        row_index = 1
        self._registry_rows: dict = {}
        if self._orphans:
            banner = ctk.CTkFrame(
                self.registry_list, fg_color=ui.COLORS["sidebar_hover"], corner_radius=8,
            )
            banner.grid(row=row_index, column=0, sticky="ew", pady=(0, int(8 * self.scale)))
            banner.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                banner,
                text=f"В папке документов найдено файлов вне реестра: {len(self._orphans)}.\n"
                     "Их можно восстановить в реестр или убрать в служебную папку _trash.",
                font=ui.font(12), text_color=ui.COLORS["text"], anchor="w", justify="left",
                wraplength=int(500 * self.scale),
            ).grid(row=0, column=0, sticky="ew", padx=int(10 * self.scale),
                   pady=int(8 * self.scale))
            ui.ghost_button(
                banner, "Разобрать", self._open_reconcile_dialog,
                width=int(110 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(0, int(10 * self.scale)))
            row_index += 1

        if not parties:
            empty_text = (
                "Ничего не найдено — измените запрос (поиск идёт по названию и ИНН)."
                if query else
                "Пока нет ни одного контрагента.\n"
                "Добавьте первого через кнопку «Добавить контрагента»."
            )
            ctk.CTkLabel(
                self.registry_list, text=empty_text,
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=row_index, column=0, sticky="ew", pady=(0, int(12 * self.scale)))
            row_index += 1

        for rec in parties:
            selected = self._registry_selected_id == rec.get("id")
            text = self._party_row_text(rec, selected)
            row = ctk.CTkFrame(self.registry_list, fg_color="transparent")
            row.grid(row=row_index, column=0, sticky="ew", pady=1)
            row.grid_columnconfigure(0, weight=1)
            btn = ctk.CTkButton(
                row,
                text=text,
                anchor="w",
                font=ui.font(12, "bold" if selected else "normal"),
                fg_color="transparent",
                hover_color=ui.COLORS["sidebar_hover"],
                text_color=ui.COLORS["text"] if selected else ui.COLORS["text_secondary"],
                corner_radius=8,
                height=int(30 * self.scale),
                command=lambda r=rec: self._select_registry_party(r),
            )
            btn.grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Редактировать", lambda r=rec: self._edit_counterparty(r),
                width=int(110 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(int(6 * self.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda r=rec: self._delete_counterparty(r),
                width=int(92 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=2, padx=(int(6 * self.scale), 0))
            self._registry_rows[rec.get("id")] = {"rec": rec, "frame": row, "btn": btn}
            row_index += 1

        self._registry_docs_row = row_index
        row_index = self._render_registry_documents(row_index)
        row_index = self._render_registry_unknown(row_index)
        self._render_registry_trash(row_index)
        self.root.after(60, self._refresh_scrollbars)

    def _render_registry_documents(self, row_index: int) -> int:
        """Документы выбранного контрагента под списком. Возвращает следующий row."""
        holder = ctk.CTkFrame(self.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(10 * self.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        self._registry_docs_frame = holder

        party = self.store.get_counterparty(self._registry_selected_id) \
            if self._registry_selected_id is not None else None
        if party is None:
            ctk.CTkLabel(
                holder, text="Выберите контрагента из списка — ниже появятся его документы.",
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=0, column=0, sticky="ew")
            return row_index + 1

        docs = self.store.documents_for(party["id"])
        ctk.CTkLabel(
            holder, text=f"Документы — {storage.display_name(party)} ({len(docs)})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(6 * self.scale)))

        if not docs:
            ctk.CTkLabel(
                holder, text="У контрагента пока нет документов.",
                font=ui.font(12), text_color=ui.COLORS["text_tertiary"], justify="left",
            ).grid(row=1, column=0, sticky="ew")
            return row_index + 1

        for offset, rec in enumerate(docs, start=1):
            kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
            number = rec.get("number") or "б/н"
            date_text = rec.get("date") or ""
            text = f"{kind_label} № {number}"
            if date_text:
                text += f" от {date_text}"

            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=text, font=ui.font(12), text_color=ui.COLORS["text"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Открыть Word", lambda d=rec: self.export_document(d),
                width=int(120 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(int(8 * self.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda d=rec: self._delete_document(d),
                width=int(92 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=2, padx=(int(6 * self.scale), 0))
        return row_index + 1

    def _render_registry_unknown(self, row_index: int) -> int:
        """Неопознанные документы (папка «unknown») в конце реестра."""
        docs = [
            rec for rec in self.store.documents
            if rec.get("kind") == storage.KIND_UNKNOWN
            and not storage.is_deleted(rec)
        ]
        if not docs:
            return row_index

        holder = ctk.CTkFrame(self.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(10 * self.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            holder, text=f"Неопознанные документы ({len(docs)})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, int(2 * self.scale)))
        ctk.CTkLabel(
            holder,
            text="Файлы, которые не удалось распознать, скопированы "
                 "в папку «unknown» внутри папки документов.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left",
            wraplength=int(560 * self.scale),
        ).grid(row=1, column=0, sticky="ew", pady=(0, int(4 * self.scale)))

        for offset, rec in enumerate(docs, start=2):
            source = str((rec.get("extra") or {}).get("source_file") or "")
            name = os.path.basename(source) if source else "Документ"
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=name, font=ui.font(12),
                text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            # Неопознанный файл нельзя запускать по клику: только показать в папке.
            ui.ghost_button(
                row, "Показать в папке",
                lambda s=source: show_in_folder(s),
                width=int(140 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(int(8 * self.scale), 0))
            ui.ghost_button(
                row, "Удалить", lambda d=rec: self._delete_document(d),
                width=int(92 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=2, padx=(int(6 * self.scale), 0))
        return row_index + 1

    # -- Корзина в реестре ------------------------------------------------ #
    def _render_registry_trash(self, row_index: int) -> int:
        """Раздел «Корзина»: восстановление и окончательная очистка (п. 2.6)."""
        trashed_parties = self.store.trashed_counterparties()
        trashed_docs = [
            rec for rec in self.store.trashed_documents()
            if not rec.get("deleted_with_party")
        ]
        if not trashed_parties and not trashed_docs:
            return row_index

        holder = ctk.CTkFrame(self.registry_list, fg_color="transparent")
        holder.grid(row=row_index, column=0, sticky="ew", pady=(int(12 * self.scale), 0))
        holder.grid_columnconfigure(0, weight=1)
        total = len(trashed_parties) + len(trashed_docs)
        ctk.CTkLabel(
            holder, text=f"Корзина ({total})",
            font=ui.font(14, "bold"), text_color=ui.COLORS["text"], anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ui.ghost_button(
            holder, "Очистить корзину", self._empty_trash,
            width=int(150 * self.scale), height=int(26 * self.scale),
        ).grid(row=0, column=1, padx=(int(8 * self.scale), 0))
        ctk.CTkLabel(
            holder,
            text="Записи удалены из реестра, файлы на диске остались. "
                 "Восстановление вернёт документы в работу.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left", wraplength=int(560 * self.scale),
        ).grid(row=1, column=0, columnspan=2, sticky="ew",
               pady=(int(2 * self.scale), int(4 * self.scale)))

        offset = 2
        for party in trashed_parties:
            docs = self.store.documents_for(party.get("id"), include_deleted=True)
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, columnspan=2, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row,
                text=f"Контрагент «{storage.display_name(party)}» · документов: {len(docs)}",
                font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Восстановить", lambda r=party: self._restore_counterparty(r),
                width=int(120 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(int(8 * self.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda r=party: self._purge_counterparty(r),
                width=int(140 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=2, padx=(int(6 * self.scale), 0))
            offset += 1

        for rec in trashed_docs:
            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            number = rec.get("number") or "б/н"
            date_text = rec.get("date") or ""
            party_name = self.store.party_name(rec.get("party_id"))
            text = f"{kind_label} № {number}"
            if date_text:
                text += f" от {date_text}"
            text += f" · {party_name}"
            row = ctk.CTkFrame(holder, fg_color="transparent")
            row.grid(row=offset, column=0, columnspan=2, sticky="ew", pady=2)
            row.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                row, text=text, font=ui.font(12),
                text_color=ui.COLORS["text_secondary"], anchor="w",
            ).grid(row=0, column=0, sticky="ew")
            ui.ghost_button(
                row, "Восстановить", lambda d=rec: self._restore_document(d),
                width=int(120 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=1, padx=(int(8 * self.scale), 0))
            ui.ghost_button(
                row, "Удалить навсегда", lambda d=rec: self._purge_document(d),
                width=int(140 * self.scale), height=int(26 * self.scale),
            ).grid(row=0, column=2, padx=(int(6 * self.scale), 0))
            offset += 1
        return row_index + 1

    # -- Сверка реестра с папкой документов ------------------------------- #
    def _scan_documents_consistency(self) -> None:
        """Фоновая сверка: ищет файлы документов вне реестра (п. 2.4)."""
        try:
            base = self._documents_dir(interactive=False)
            orphans = reconcile.find_orphans(base, self.store)
        except Exception as exc:  # noqa: BLE001 — фоновая задача не должна падать
            self._log.warning("Сверка реестра с папкой не удалась: %s", type(exc).__name__)
            return
        self._orphans = orphans
        if orphans:
            self._log.info("Файлов вне реестра: %d", len(orphans))
        if "registry" in self.pages:
            self.refresh_registry()

    def _open_reconcile_dialog(self) -> None:
        """Диалог разбора файлов вне реестра: восстановить или убрать в _trash."""
        base = self._documents_dir()
        orphans = reconcile.find_orphans(base, self.store)
        self._orphans = orphans
        if not orphans:
            ui.alert(
                self.root, "Файлы вне реестра",
                "Все файлы документов учтены в реестре.", "info",
            )
            self.refresh_registry()
            return

        preview = "\n".join(f"• {os.path.basename(path)}" for path in orphans[:8])
        if len(orphans) > 8:
            preview += f"\n• …и ещё {len(orphans) - 8}"
        choice = self._choice_dialog(
            "Файлы вне реестра",
            f"В папке документов найдено файлов вне реестра: {len(orphans)}.\n\n"
            f"{preview}\n\n"
            "«Восстановить» — создать записи реестра (тип документа и контрагент "
            "определяются по имени файла и папки).\n"
            "«В корзину» — перенести файлы в служебную папку _trash внутри папки "
            "документов.",
            [
                ("skip", "Оставить", False),
                ("trash", "В корзину", False),
                ("restore", "Восстановить в реестр", True),
            ],
        )
        if choice == "restore":
            added = reconcile.restore_orphans(self.store, base, orphans, safe_filename)
            if self.store.last_error:
                ui.alert(
                    self.root, "Восстановление не удалось",
                    self.store.last_error, "error",
                )
            else:
                ui.alert(
                    self.root, "Реестр восстановлен",
                    f"Добавлено записей в реестр: {added}.\n\n"
                    "Проверьте типы документов и контрагентов в реестре.",
                    "success",
                )
            self._orphans = reconcile.find_orphans(base, self.store)
            self._refresh_party_combo_values()
            self.refresh_registry()
            self._update_status()
        elif choice == "trash":
            moved = reconcile.move_orphans_to_trash(base, orphans)
            ui.alert(
                self.root, "Файлы перемещены",
                f"Перемещено файлов: {moved}.\n\n"
                f"Папка: {os.path.join(base, reconcile.TRASH_DIR_NAME)}",
                "info",
            )
            self._orphans = reconcile.find_orphans(base, self.store)
            self.refresh_registry()

    # ------------------------------------------------------------------ #
    #  Помощники страниц
    # ------------------------------------------------------------------ #
    def _digits_vcmd_for(self, max_len: int):
        """Валидатор «только цифры, не длиннее max_len»."""
        return (
            self.root.register(lambda value: value == "" or (value.isdigit() and len(value) <= max_len)),
            "%P",
        )

    def _add_date_row(self, card, page_key: str, day_var: tk.StringVar, month_var: tk.StringVar):
        holder = ctk.CTkFrame(card.body, fg_color="transparent")

        day = ui.entry(
            holder, width=int(DAY_WIDTH * self.scale),
            textvariable=day_var, validatecommand=(self._day_vcmd, "%P"),
        )
        day.pack(side="left")
        self._page_widgets[(page_key, F.DAY)] = day

        ctk.CTkLabel(
            holder, text="число", font=ui.font(12), text_color=ui.COLORS["text_tertiary"],
        ).pack(side="left", padx=(int(8 * self.scale), int(10 * self.scale)))

        month = ui.option_menu(
            holder, MONTHS_GENITIVE, width=int(MONTH_WIDTH * self.scale), variable=month_var,
        )
        month.pack(side="left")

        ui.ghost_button(
            holder, "Сегодня", lambda: self._fill_date(day_var, month_var),
            width=int(92 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(12 * self.scale), 0))

        card.row("Дата", holder, stretch=True,
                 error_var=self._error_var_for(page_key, F.DAY))

    def _fill_date(self, day_var: tk.StringVar, month_var: tk.StringVar):
        today = datetime.now()
        day_var.set(format_date(today.day))
        month_var.set(MONTHS_GENITIVE[today.month - 1])

    def _error_var_for(self, page_key: str, key: str) -> tk.StringVar:
        return self._error_vars.setdefault((page_key, key), tk.StringVar(value=""))

    def _set_page_error(self, page_key: str, key: str, message: str):
        self._error_var_for(page_key, key).set(message)
        widget = self._page_widgets.get((page_key, key))
        if widget is not None and hasattr(widget, "set_error"):
            widget.set_error(True)

    def _clear_page_errors(self, page_key: str):
        for (page, key), var in self._error_vars.items():
            if page != page_key:
                continue
            var.set("")
            widget = self._page_widgets.get((page, key))
            if widget is not None and hasattr(widget, "set_error"):
                widget.set_error(False)

    def _on_doc_field_changed(self, page_key: str, key: str):
        self._error_var_for(page_key, key).set("")
        widget = self._page_widgets.get((page_key, key))
        if widget is not None and hasattr(widget, "set_error"):
            widget.set_error(False)

    def _focus_page_widget(self, page_key: str, key: str):
        widget = self._page_widgets.get((page_key, key))
        if widget is not None:
            try:
                widget.focus_set()
            except Exception:
                pass

    # ------------------------------------------------------------------ #
    #  Контрагенты: выпадающие списки
    # ------------------------------------------------------------------ #
    def _party_duplicate_bases(self) -> set[str]:
        """Названия, которые встречаются у нескольких карточек (п. 4.1)."""
        counts: dict[str, int] = {}
        for rec in self.store.active_counterparties():
            base = storage.display_with_inn(rec)
            counts[base] = counts.get(base, 0) + 1
        return {base for base, count in counts.items() if count > 1}

    def _party_display(self, rec: dict, duplicates: set[str] | None = None) -> str:
        """Имя в списке; при совпадении названий добавляется ID."""
        base = storage.display_with_inn(rec)
        if duplicates and base in duplicates:
            return f"{base} · ID {rec.get('id')}"
        return base

    def _party_display_values(self) -> list[str]:
        recs = self.store.active_counterparties()
        if not recs:
            return ["— контрагентов нет —"]
        duplicates = self._party_duplicate_bases()
        return [self._party_display(rec, duplicates) for rec in recs]

    def _refresh_party_combo_values(self):
        """Обновляет списки и сбрасывает устаревший выбор (п. 4.2).

        Если выбранного контрагента больше нет в реестре (удалён, переименован),
        поле очищается — недействительное имя не должно оставаться на экране.
        """
        values = self._party_display_values()
        valid = set(values)
        placeholders = {"— контрагентов нет —"}
        for combo, attr in (
            (getattr(self, "contract_party_combo", None), "contract_party_var"),
            (getattr(self, "addl_party_combo", None), "addl_party_var"),
            (getattr(self, "spec_party_combo", None), "spec_party_var"),
        ):
            if combo is not None:
                combo.set_values(values)
            var = getattr(self, attr, None)
            if var is None:
                continue
            current = var.get().strip()
            if current and current not in valid and current not in placeholders:
                var.set("")
        if self._current_party_id is not None:
            if self.store.get_counterparty(self._current_party_id) is None:
                self._current_party_id = None
        self._update_current_party_labels()

    def _party_from_var(self, var: tk.StringVar) -> dict | None:
        """Возвращает запись контрагента по выбранному значению списка."""
        value = var.get().strip()
        if not value:
            return None
        recs = self.store.active_counterparties()
        duplicates = self._party_duplicate_bases()
        for rec in recs:
            if self._party_display(rec, duplicates) == value:
                return rec
        # Совместимость: значение могло быть установлено программно без ID.
        for rec in recs:
            if storage.display_with_inn(rec) == value:
                return rec
        return None

    # -- Единый «текущий контрагент» (п. 4.2) --------------------------- #
    def _current_party_label(self) -> str:
        if self._current_party_id is None:
            return "Документ для: контрагент не выбран"
        rec = self.store.get_counterparty(self._current_party_id)
        if rec is None:
            return "Документ для: выбранный контрагент не найден — выберите заново"
        return f"Документ для: {self._party_display(rec, self._party_duplicate_bases())}"

    def _update_current_party_labels(self) -> None:
        text = self._current_party_label()
        for var in getattr(self, "_current_party_vars", {}).values():
            try:
                var.set(text)
            except Exception:
                pass

    def _set_current_party(self, rec: dict | None, source: str | None = None) -> None:
        """Запоминает контрагента и переносит выбор на все страницы документов."""
        if not isinstance(rec, dict):
            return
        self._current_party_id = rec.get("id")
        display = self._party_display(rec, self._party_duplicate_bases())
        self._syncing_party = True
        try:
            for key, attr in (
                ("contract", "contract_party_var"),
                ("additional", "addl_party_var"),
                ("spec", "spec_party_var"),
            ):
                if key == source:
                    continue
                var = getattr(self, attr, None)
                if var is None:
                    continue
                try:
                    if var.get() != display:
                        var.set(display)
                except Exception:
                    pass
        finally:
            self._syncing_party = False
        self._update_current_party_labels()

    def _on_party_selected(self, key: str) -> None:
        """Пользователь выбрал контрагента на одной из страниц документов."""
        if getattr(self, "_syncing_party", False):
            return
        attr = {
            "contract": "contract_party_var",
            "additional": "addl_party_var",
            "spec": "spec_party_var",
        }.get(key)
        var = getattr(self, attr, None) if attr else None
        if var is None:
            return
        rec = self._party_from_var(var)
        if rec is None:
            if not var.get().strip():
                self._current_party_id = None
                self._update_current_party_labels()
            return
        self._set_current_party(rec, source=key)

    def _update_contract_info(self, page: str) -> None:
        """Показывает выбранный договор-основание явно (п. 4.2)."""
        if page == "additional":
            var = getattr(self, "addl_contract_var", None)
            info = getattr(self, "addl_contract_info_var", None)
        else:
            var = getattr(self, "spec_contract_var", None)
            info = getattr(self, "spec_contract_info_var", None)
        if var is None or info is None:
            return
        value = var.get().strip()
        if not value or value.startswith("—"):
            info.set("Договор-основание не выбран — укажите его явно.")
        else:
            info.set(f"Договор-основание: {value}")

    def _contract_display(self, rec: dict) -> str:
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        return f"№ {number} от {date_text}" if date_text else f"№ {number}"

    def _contract_from_var(self, var: tk.StringVar, party_id) -> dict | None:
        value = var.get().strip()
        if not value:
            return None
        for rec in self.store.contracts_for(party_id):
            if self._contract_display(rec) == value:
                return rec
        return None

    # ------------------------------------------------------------------ #
    #  Формирование текста документов
    # ------------------------------------------------------------------ #
    @staticmethod
    def _set_text(textbox, text: str):
        if textbox is None:
            return
        try:
            textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.insert("1.0", text)
        finally:
            # Предпросмотр только для чтения: правки в нём не попадают в .docx.
            try:
                textbox.configure(state="disabled")
            except Exception:
                pass

    # -- Печатная форма: сборка .docx напрямую из шаблонов ---------------- #
    def _default_documents_dir(self, interactive: bool = True) -> str:
        """Папка документов по умолчанию: «Документы/EW.DocHub» (п. 6.2).

        Раньше по умолчанию использовался скрытый %APPDATA%\\DogovorGenerator\\documents:
        пользователь не видел, где лежат его файлы. Теперь — видимая папка
        «Документы», которая не исчезает при переустановке системы и попадает
        в обычные резервные копии. Если «Документы» недоступны — возврат
        к прежнему месту в профиле пользователя.
        """
        base = os.path.join(_user_documents_dir(), "EW.DocHub")
        try:
            os.makedirs(base, exist_ok=True)
            return base
        except OSError as exc:
            self._log.error(
                "Папка документов по умолчанию недоступна: %s", type(exc).__name__,
            )
            if interactive:
                ui.alert(
                    self.root, "Папка документов недоступна",
                    "Не удалось создать папку документов по умолчанию:\n"
                    f"{base}\n\n{exc}\n\n"
                    "Выберите другую папку в «Настройках» или проверьте права доступа.",
                    "error",
                )
        base = os.path.join(
            os.environ.get("APPDATA") or os.path.expanduser("~"),
            SETTINGS_DIR_NAME, "documents",
        )
        try:
            os.makedirs(base, exist_ok=True)
        except OSError:
            pass
        return base

    def _resolve_unavailable_documents_dir(self, custom: str, exc: OSError) -> str:
        """Спрашивает, куда сохранять, если выбранная папка недоступна (п. 2.8)."""
        self._log.warning("Папка документов недоступна: %s", type(exc).__name__)
        if self._docs_dir_fallback_active or self._docs_dir_dialog_active:
            return self._default_documents_dir()
        if self._docs_dir_retry_count >= 2:
            self._docs_dir_fallback_active = True
            return self._default_documents_dir()

        self._docs_dir_dialog_active = True
        try:
            choice = self._choice_dialog(
                "Папка документов недоступна",
                f"Не удалось открыть папку:\n{custom}\n\n{exc}\n\n"
                "Куда сохранять новые документы?",
                [
                    ("default", "Папка по умолчанию", False),
                    ("choose", "Выбрать другую", False),
                    ("retry", "Повторить", True),
                ],
            )
        finally:
            self._docs_dir_dialog_active = False

        if choice == "retry":
            self._docs_dir_retry_count += 1
            return self._documents_dir(interactive=True)
        if choice == "choose":
            chosen = filedialog.askdirectory(
                parent=self.root, title="Выберите папку для документов",
            )
            if chosen:
                self._settings["documents_dir"] = chosen
                self._docs_dir_retry_count = 0
                self._docs_dir_fallback_active = False
                self._persist_settings()
                return chosen
        self._docs_dir_fallback_active = True
        return self._default_documents_dir()

    def _documents_dir(self, interactive: bool = True) -> str:
        """Папка для готовых документов: выбранная в настройках либо %APPDATA%.

        Недоступная папка больше не подменяется молча (п. 2.8): пользователь
        выбирает «повторить / выбрать другую / по умолчанию» и видит, куда
        реально сохранится документ.
        """
        custom = str(self._settings.get("documents_dir") or "").strip()
        if custom:
            try:
                os.makedirs(custom, exist_ok=True)
                self._docs_dir_retry_count = 0
                return custom
            except OSError as exc:
                if interactive:
                    return self._resolve_unavailable_documents_dir(custom, exc)
                self._log.warning(
                    "Папка документов недоступна (фоновая задача): %s",
                    type(exc).__name__,
                )
        return self._default_documents_dir(interactive=interactive)

    def _documents_dir_for_import(self) -> str:
        """Папка документов для фонового импорта: без Tk-диалогов.

        Импорт выполняется в отдельном потоке, где нельзя показывать окна.
        Если выбранная пользователем папка недоступна, возвращается ошибка —
        она попадёт в отчёт импорта, а не приведёт к молчаливой подмене пути.
        """
        custom = str(self._settings.get("documents_dir") or "").strip()
        if custom:
            try:
                os.makedirs(custom, exist_ok=True)
                return custom
            except OSError as exc:
                raise OSError(
                    "Папка документов недоступна:\n"
                    f"{custom}\n{exc}\n\n"
                    "Выберите другую папку в «Настройках» и повторите импорт."
                ) from exc
        return self._default_documents_dir(interactive=False)

    @staticmethod
    def _is_inside_directory(base: str, path: str) -> bool:
        """Проверяет, что ``path`` не выходит за пределы ``base`` (п. 2.7)."""
        try:
            base_norm = os.path.normcase(os.path.normpath(base))
            path_norm = os.path.normcase(os.path.normpath(path))
            return os.path.commonpath([base_norm, path_norm]) == base_norm
        except (ValueError, OSError):
            return False

    def _fit_document_path(self, directory: str, base: str, ext: str = ".docx") -> str:
        """Собирает путь, укладывающийся в ``MAX_PATH_LENGTH`` (п. 2.7)."""
        path = os.path.join(directory, base + ext)
        if len(path) <= MAX_PATH_LENGTH:
            return path
        available = MAX_PATH_LENGTH - len(directory) - len(ext) - 1
        if available < 12:
            raise OSError(
                "Слишком длинный путь к папке документов — выберите папку короче:\n"
                + directory
            )
        digest = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8]
        shortened = base[: max(1, available - 9)].rstrip(" .") + "~" + digest
        return os.path.join(directory, shortened + ext)

    def _party_documents_dir(self, party_rec: dict | None,
                             interactive: bool = True) -> str:
        """Индивидуальная папка контрагента внутри общей папки документов.

        Все документы реестра хранятся по пути <папка документов>/<контрагент>/.
        Имя папки санитизируется, итоговый путь проверяется на выход за
        пределы папки документов.
        """
        base = self._documents_dir(interactive=interactive)
        if party_rec is None:
            return base
        folder = safe_filename(storage.display_name(party_rec), "Контрагент")
        path = os.path.normpath(os.path.join(base, folder))
        if (
            not self._is_inside_directory(base, path)
            or os.path.normcase(path) == os.path.normcase(os.path.normpath(base))
        ):
            self._log.error(
                "Имя папки контрагента выходит за пределы папки документов",
            )
            ui.alert(
                self.root, "Недопустимое имя папки",
                "Имя контрагента не может быть использовано как папка. "
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        try:
            os.makedirs(path, exist_ok=True)
        except OSError as exc:
            self._log.warning(
                "Не удалось создать папку контрагента: %s", type(exc).__name__,
            )
            ui.alert(
                self.root, "Не удалось создать папку контрагента",
                f"{path}\n\n{exc}\n\n"
                "Документ будет сохранён в общую папку документов.",
                "warning",
            )
            return base
        return path

    def _render_contract_docx(self, data: dict, is_ip: bool) -> "Document":
        """Договор из шаблона: печатная форма сохраняется без искажений."""
        template_path = resource_path("dogovor_template_ip.docx" if is_ip else "dogovor_template.docx")
        doc = load_template_cached(template_path)
        replace_in_doc(doc, build_replacements(data, is_ip))
        return doc

    def _render_additional_docx(self, data: dict, is_ip: bool, number: str,
                                limit_display: str, term: str) -> "Document":
        """Доп. соглашение из шаблона с подстановкой лимита и отсрочки."""
        template_path = resource_path("additional_agreement_ip.docx" if is_ip else "additional_agreement.docx")
        doc = load_template_cached(template_path)
        replacements = build_replacements(data, is_ip)
        replacements[PH.AGREEMENT_NUMBER] = number or "б/н"
        limit = digits_only(limit_display)
        replacements.update({
            PH.LIMIT: limit_display,
            PH.LIMIT_WORDS: amount_to_words_ru(int(limit)) if limit else "",
            PH.TERM: term,
            PH.TERM_WORDS: number_to_words_ru(int(term)) if term.isdigit() else "",
        })
        replace_in_doc(doc, replacements)
        return doc

    def _render_spec_docx(self, rec: dict) -> "Document":
        """Спецификация из шаблона (книжная ориентация).

        У ООО и ИП разные шаблоны: у ИП нет КПП, а вместо ОГРН — ОГРНИП.
        Общий spec_template.docx написан «в терминах ООО», поэтому для ИП
        используется отдельная копия spec_template_ip.docx (п. 1.1):
        строка «ОГРН {{OGRN}}» заменена на «ОГРНИП {{OGRNIP}}»,
        «ИНН {{INN}} / КПП {{KPP}}» — на «ИНН {{INN}}». Без этого в печатной
        форме ИП оставались пустые подписи «ОГРН» и «/ КПП».
        """
        is_ip = self._rec_is_ip(rec)
        template_path = resource_path(
            "spec_template_ip.docx" if is_ip else "spec_template.docx"
        )
        doc = load_template_cached(template_path)
        data = dict(rec.get("data") or {})
        extra = rec.get("extra") or {}
        replacements = build_replacements(data, is_ip)

        if is_ip:
            # Шаблон спецификации написан в терминах ООО-покупателя;
            # для ИП подставляем эквивалентные значения.
            name_ip = data.get(F.IP_NAME, "")
            replacements[PH.FULL_NAME] = (
                f"Индивидуальный предприниматель {name_ip}".strip()
            )
            replacements[PH.SHORT_NAME] = f"ИП {name_ip}".strip()
            replacements[PH.DIRECTOR] = name_ip
            replacements[PH.DIRECTOR_INITIALS] = format_name_ip2(name_ip)
            replacements[PH.BASIS] = (
                "свидетельства о государственной регистрации в качестве "
                "индивидуального предпринимателя"
            )

        # Договор-основание: номер и дата подставляются из записи договора.
        contract = (
            self.store.get_document(extra.get("contract_id"))
            if extra.get("contract_id") else None
        )
        if contract is None or contract.get("kind") != storage.KIND_CONTRACT:
            contract = None
        contract_data = dict((contract or {}).get("data") or {})
        replacements[PH.NUMBER] = (
            (contract or {}).get("number") or extra.get("contract_number") or ""
        )
        replacements[PH.DATE] = format_date(contract_data.get(F.DAY, ""))
        replacements[PH.MONTH] = contract_data.get(F.MONTH, "")

        # Кол-во/цена в реестре хранятся числами (float/int) — переводим в Decimal.
        items = [
            dict(item, qty=parse_decimal(item.get("qty")),
                 price=parse_decimal(item.get("price")))
            for item in (extra.get("items") or [])
        ]
        total = parse_decimal(extra.get("total"))
        if total == 0 and items:
            total = sum((i["qty"] * i["price"] for i in items), Decimal("0"))
        replacements.update({
            PH.APPENDIX_NUMBER: SPEC_APPENDIX_NUMBER,
            PH.SPEC_NUMBER: rec.get("number") or "б/н",
            PH.SPEC_DAY: str(extra.get("spec_day", "")),
            PH.SPEC_MONTH: str(extra.get("spec_month", "")),
            PH.PRICES_VALID_UNTIL: prices_valid_until_default(),
            PH.TOTAL: format_amount(total),
            PH.TOTAL_IN_WORDS: amount_in_words_ru(total) if total else "",
            PH.VAT: format_amount(
                spec_vat_amount_from_items(items) if items else spec_vat_amount(total)
            ),
            PH.VAT_RATE: f"{SPEC_VAT_RATE}%",
            PH.DELIVERY_TERM: str(extra.get("delivery_term", "")),
            PH.PAYMENT_TERMS: str(extra.get("payment_terms", "")),
            PH.DELIVERY_BASIS: str(extra.get("delivery_basis", "")),
        })
        mark_spec_final_page(doc)
        replace_in_doc(doc, replacements)
        fill_spec_items(doc, items)
        return doc

    def _rec_is_ip(self, rec: dict) -> bool:
        """Признак ИП у сохранённой записи (учитывая старые записи реестра)."""
        extra = rec.get("extra") or {}
        if "is_ip" in extra:
            return bool(extra.get("is_ip"))
        party = self.store.get_counterparty(rec.get("party_id"))
        return bool(party and party.get("type") == "ip")

    def _build_record_docx(self, rec: dict) -> "Document":
        """Восстанавливает .docx записи реестра из шаблона (договор и доп. соглашение)."""
        kind = rec.get("kind")
        extra = rec.get("extra") or {}
        if kind == storage.KIND_CONTRACT:
            return self._render_contract_docx(rec.get("data") or {}, self._rec_is_ip(rec))
        if kind == storage.KIND_ADDITIONAL:
            return self._render_additional_docx(
                rec.get("data") or {}, self._rec_is_ip(rec),
                rec.get("number") or "", extra.get("limit", ""), extra.get("term", ""),
            )
        if kind == storage.KIND_SPEC:
            return self._render_spec_docx(rec)
        return docx_from_text(rec.get("text", ""))

    def _write_docx(self, doc, rec: dict) -> str:
        """Пишет .docx в папку контрагента атомарно и возвращает итоговый путь."""
        directory = self._party_documents_dir(
            self.store.get_counterparty(rec.get("party_id"))
        )
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        party_name = self.store.party_name(rec.get("party_id"))
        number = rec.get("number") or "б/н"
        base = safe_filename(f"{kind_label} №{number} {party_name}", "Документ")
        path = self._fit_document_path(directory, base, ".docx")
        counter = 2
        while os.path.exists(path):
            path = self._fit_document_path(directory, f"{base} ({counter})", ".docx")
            counter += 1
        with applog.timed("Запись .docx", 500, "documents"):
            save_docx_atomic(doc, path)
        self._log.info("Документ записан: тип=%s, id=%s", rec.get("kind"), rec.get("id"))
        return path

    def _save_and_open_document(self, rec: dict) -> str:
        """Готовит .docx в индивидуальной папке контрагента и открывает его в Word."""
        path = self._write_docx(self._build_record_docx(rec), rec)
        open_file(path)
        return path

    def _rollback_document(self, rec: dict, path: str) -> None:
        """Откатывает запись и файл, если реестр не удалось записать на диск."""
        try:
            self.store.documents.remove(rec)
        except ValueError:
            pass
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except OSError:
            pass
        ui.alert(
            self.root, "Не удалось сохранить документ",
            f"{self.store.last_error or 'Ошибка записи реестра.'}\n\n"
            "Файл удалён, запись в реестр не добавлена.",
            "error",
        )

    def _regenerate_spec(self, *_args):
        """Откладывает пересборку предпросмотра спецификации (п. 5.6).

        Раньше предпросмотр пересобирался на каждый символ в любом из шести
        полей и в каждой ячейке таблицы позиций. Теперь, как и у договора
        и доп. соглашения, пересборка откладывается на 120 мс.
        """
        pending = getattr(self, "_spec_preview_after", None)
        if pending:
            try:
                self.root.after_cancel(pending)
            except Exception:
                pass
        self._spec_preview_after = self.root.after(120, self._render_spec_preview)

    def _flush_spec_preview(self):
        """Синхронно дособирает отложенный предпросмотр (перед чтением текста)."""
        pending = getattr(self, "_spec_preview_after", None)
        if pending:
            try:
                self.root.after_cancel(pending)
            except Exception:
                pass
            self._spec_preview_after = None
        self._render_spec_preview()

    def _render_spec_preview(self):
        self._spec_preview_after = None
        party = self._party_from_var(self.spec_party_var)
        if party is None:
            self._set_text(self.spec_text, "Выберите контрагента и договор-основание слева.")
            return
        contract = self._contract_from_var(self.spec_contract_var, party["id"])
        data = dict(party.get("fields", {}))

        lines: list[str] = []
        if contract is not None:
            lines.append(f"Приложение № {SPEC_APPENDIX_NUMBER} к Договору поставки "
                         f"{self._contract_display(contract)}")
        lines.append("")
        number = self.spec_number_var.get().strip() or "б/н"
        lines.append(f"СПЕЦИФИКАЦИЯ № {number}")
        lines.append(
            f"г. Хабаровск    «{self.spec_day_var.get().strip() or '__'}» "
            f"{self.spec_month_var.get().strip() or '__'} 2026 г."
        )
        lines.append("")
        lines.append("Поставщик: ООО «Эдельвейс»")
        lines.append(f"Покупатель: {storage.display_name(party)}, "
                     f"ИНН {data.get(F.INN, '')}, {data.get(F.LEGAL_ADDRESS, '')}")
        lines.append("")
        prices_until = prices_valid_until_default()
        self.spec_prices_until_var.set(prices_until)
        lines.append(f"Цены в Спецификации действительны до {prices_until}.")
        lines.append("Наименование, количество и стоимость:")

        items = self._spec_items()
        if items:
            lines.append("№   Наименование                Кол-во   Ед. изм.   Цена, руб.     Сумма, руб.")
            for index, item in enumerate(items, start=1):
                amount = (item["qty"] * item["price"]).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )
                lines.append(
                    f"{index}.  {item['name']}   —   {format_quantity(item['qty'])} "
                    f"{item['unit']} × {format_amount(item['price'])} = "
                    f"{format_amount(amount)} руб."
                )
            total = sum((i["qty"] * i["price"] for i in items), Decimal("0"))
            lines.append(
                f"Итого: {format_amount(total)} руб. ({amount_in_words_ru(total)})"
            )
            lines.append(
                f"в т.ч. НДС {SPEC_VAT_RATE}%: "
                f"{format_amount(spec_vat_amount_from_items(items))} руб."
            )
        else:
            lines.append("Позиции не заполнены — заполните таблицу позиций справа.")
        lines.append("")
        conditions = [
            (F.DELIVERY_TERM, self.spec_delivery_term_var.get().strip()),
            (F.PAYMENT_TERMS, self.spec_payment_terms_var.get().strip()),
            (F.DELIVERY_BASIS, self.spec_delivery_basis_var.get().strip()),
        ]
        for condition_label, condition_value in conditions:
            if condition_value:
                lines.append(f"{condition_label}: {condition_value}")
        lines.append("")
        lines.append("Подписи Сторон: от Поставщика — А.Н. Верещагин, "
                     f"от Покупателя — {data.get(F.DIRECTOR, data.get(F.IP_NAME, ''))}")
        self._set_text(self.spec_text, "\n".join(lines))

    # ------------------------------------------------------------------ #
    #  Сохранение документов в реестр
    # ------------------------------------------------------------------ #
    def _confirm_document_duplicate(
        self, party_id, kind: str, number: str, date_text: str
    ) -> bool:
        """Предупреждает о документе с тем же номером у того же контрагента (п. 4.1)."""
        number = str(number or "").strip()
        if not number:
            return True
        for rec in self.store.documents_for(party_id):
            if rec.get("kind") != kind:
                continue
            if str(rec.get("number", "")).strip() != number:
                continue
            existing_date = str(rec.get("date", "")).strip()
            title = storage.KIND_LABELS.get(kind, "Документ")
            details = f"{title} № {number}"
            if existing_date:
                details += f" от {existing_date}"
            return ui.confirm(
                self.root, "Повтор номера",
                f"У этого контрагента уже есть {details}.\n\n"
                "Создать ещё один документ с тем же номером?",
                ok_text="Создать ещё", cancel_text="Отмена",
            )
        return True

    def save_contract(self):
        try:
            rec = self._party_from_var(self.contract_party_var)
            if rec is None:
                ui.alert(self.root, "Контрагент не выбран",
                         "Выберите контрагента из списка или добавьте нового "
                         "через кнопку «Добавить контрагента».", "error")
                return

            data = dict(rec.get("fields", {}))
            data[F.NUMBER] = self.contract_number_var.get().strip()
            data[F.DAY] = self.contract_day_var.get().strip()
            data[F.MONTH] = self.contract_month_var.get().strip()
            is_ip = rec.get("type") == "ip"

            errors = validate_contract_fields(data, is_ip)
            if errors:
                self._clear_page_errors("contract")
                for key, message in errors.items():
                    self._set_page_error("contract", key, message)
                for key in (F.NUMBER, F.DAY):
                    self._focus_page_widget("contract", key)
                    break
                ui.alert(
                    self.root, "Ошибки в данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            date_text = f"{data[F.DAY]} {data[F.MONTH]}".strip()
            if not self._confirm_document_duplicate(
                rec["id"], storage.KIND_CONTRACT, data[F.NUMBER], date_text
            ):
                return
            extra = {"day": data[F.DAY], "month": data[F.MONTH], "is_ip": is_ip}
            draft = {
                "kind": storage.KIND_CONTRACT,
                "party_id": rec["id"],
                "number": data[F.NUMBER],
                "data": data,
                "extra": extra,
                "text": "",
            }
            # Сначала успешно пишем .docx, затем регистрируем запись: при сбое
            # записи в реестр «осиротевший» файл удаляется, а не остаётся навсегда.
            path = self._write_docx(self._build_record_docx(draft), draft)
            rec_doc = self.store.add_document(
                storage.KIND_CONTRACT, rec["id"], data[F.NUMBER], date_text,
                data=data, extra=extra, text="",
            )
            if self.store.last_error:
                self._rollback_document(rec_doc, path)
                return
            self._reset_form_baseline("contract")
            self.refresh_registry()
            self._update_status()
            opened = open_file(path)
            self._notify(
                f"Договор № {data[F.NUMBER]} от {date_text} для "
                f"{storage.display_name(rec)} сохранён"
                + (" и открыт." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(self.root, "Ошибка", f"Не удалось сохранить договор:\n{exc}", "error")

    def save_additional(self):
        try:
            party = self._party_from_var(self.addl_party_var)
            if party is None:
                ui.alert(self.root, "Контрагент не выбран",
                         "Выберите контрагента из списка слева.", "error")
                return
            contract = self._contract_from_var(self.addl_contract_var, party["id"])
            if contract is None:
                ui.alert(self.root, "Договор не выбран",
                         "Выберите договор-основание. Если договоров нет — создайте "
                         "договор поставки на соответствующей странице.", "error")
                return

            limit_display = self._limit_var.get().strip()
            limit = digits_only(limit_display)
            term = self.term_var.get().strip()
            number = self.addl_number_var.get().strip() or "б/н"

            errors = {}
            if not limit:
                errors[F.LIMIT] = "Заполните поле"
            if not term:
                errors[F.TERM] = "Заполните поле"
            elif not term.isdigit():
                errors[F.TERM] = "Должно быть целое число дней"
            if errors:
                self._clear_page_errors("additional")
                for key, message in errors.items():
                    self._set_page_error("additional", key, message)
                ui.alert(
                    self.root, "Не хватает данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {error}" for error in errors.values()),
                    "error",
                )
                return

            contract_number = contract.get("number") or ""
            date_text = contract.get("date") or ""
            if number and number != "б/н" and not self._confirm_document_duplicate(
                party["id"], storage.KIND_ADDITIONAL, number, date_text
            ):
                return
            extra = {
                "contract_id": contract["id"],
                "contract_number": contract_number,
                "limit": limit_display,
                "limit_digits": limit,
                "term": term,
                "is_ip": party.get("type") == "ip",
            }
            draft = {
                "kind": storage.KIND_ADDITIONAL,
                "party_id": party["id"],
                "number": number,
                "data": dict(contract.get("data") or {}),
                "extra": extra,
                "text": "",
            }
            path = self._write_docx(self._build_record_docx(draft), draft)
            rec_doc = self.store.add_document(
                storage.KIND_ADDITIONAL, party["id"], number, date_text,
                data=dict(contract.get("data") or {}),
                extra=extra,
                text="",
            )
            if self.store.last_error:
                self._rollback_document(rec_doc, path)
                return
            self._reset_form_baseline("additional")
            self.refresh_registry()
            self._update_status()
            opened = open_file(path)
            self._notify(
                f"Доп. соглашение № {number} к договору № {contract_number or 'б/н'} "
                "сохранено"
                + (" и открыто." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(self.root, "Ошибка", f"Не удалось сохранить доп. соглашение:\n{exc}", "error")

    def save_spec(self):
        try:
            party = self._party_from_var(self.spec_party_var)
            if party is None:
                ui.alert(self.root, "Контрагент не выбран",
                         "Выберите контрагента из списка слева.", "error")
                return
            contract = self._contract_from_var(self.spec_contract_var, party["id"])
            if contract is None:
                ui.alert(self.root, "Договор не выбран",
                         "Выберите договор-основание. Если договоров нет — создайте "
                         "договор поставки на соответствующей странице.", "error")
                return

            number = self.spec_number_var.get().strip()
            if not number:
                self._set_page_error("spec", F.NUMBER, "Заполните поле")
                self._focus_page_widget("spec", F.NUMBER)
                ui.alert(self.root, "Не хватает данных",
                         "Укажите номер спецификации.", "error")
                return

            items = self._spec_items()
            if not items:
                ui.alert(self.root, "Нет позиций",
                         "Добавьте хотя бы одну позицию с наименованием.", "error")
                return
            for item in items:
                if item["qty"] <= 0:
                    ui.alert(self.root, "Проверьте позиции",
                             f"У позиции «{item['name']}» должно быть количество больше нуля.", "error")
                    return
                if item["price"] <= 0:
                    ui.alert(self.root, "Проверьте позиции",
                             f"У позиции «{item['name']}» должна быть цена больше нуля.", "error")
                    return

            # Предпросмотр собирается с задержкой — перед сохранением
            # дособираем его синхронно, чтобы текст записи был актуальным.
            self._flush_spec_preview()
            text = self.spec_text.get("1.0", "end").strip()
            date_text = f"{self.spec_day_var.get().strip()} {self.spec_month_var.get().strip()}".strip()
            contract_number = contract.get("number") or ""
            if not self._confirm_document_duplicate(
                party["id"], storage.KIND_SPEC, number, date_text
            ):
                return

            # Decimal не сериализуется в JSON — сохраняем кол-во/цены float'ами.
            total = sum((i["qty"] * i["price"] for i in items), Decimal("0"))
            items_json = [
                {"name": i["name"], "qty": float(i["qty"]),
                 "unit": i["unit"], "price": float(i["price"])}
                for i in items
            ]
            extra = {
                "contract_id": contract["id"],
                "contract_number": contract_number,
                "items": items_json,
                "total": float(total),
                "spec_day": self.spec_day_var.get().strip(),
                "spec_month": self.spec_month_var.get().strip(),
                "prices_until": self.spec_prices_until_var.get().strip(),
                "delivery_term": self.spec_delivery_term_var.get().strip(),
                "payment_terms": self.spec_payment_terms_var.get().strip(),
                "delivery_basis": self.spec_delivery_basis_var.get().strip(),
                "is_ip": party.get("type") == "ip",
            }
            draft = {
                "kind": storage.KIND_SPEC,
                "party_id": party["id"],
                "number": number,
                "data": dict(party.get("fields", {})),
                "extra": extra,
                "text": text,
            }
            path = self._write_docx(self._build_record_docx(draft), draft)
            rec_doc = self.store.add_document(
                storage.KIND_SPEC, party["id"], number, date_text,
                data=dict(party.get("fields", {})),
                extra=extra,
                text=text,
            )
            if self.store.last_error:
                self._rollback_document(rec_doc, path)
                return
            self._reset_form_baseline("spec")
            self.refresh_registry()
            self._update_status()
            opened = open_file(path)
            self._notify(
                f"Спецификация № {number} к договору № {contract_number or 'б/н'} "
                "сохранена"
                + (" и открыта." if opened else "."),
                path,
            )
        except Exception as exc:
            ui.alert(self.root, "Ошибка", f"Не удалось сохранить спецификацию:\n{exc}", "error")

    # ------------------------------------------------------------------ #
    #  Реестр: экспорт и удаление
    # ------------------------------------------------------------------ #
    def export_document(self, rec: dict):
        """Открывает документ из реестра.

        Импортированные документы открываются как исходный файл из папки
        контрагента; созданные в программе — собираются из шаблона .docx.
        Если оригинал импортированного документа потерян, пользователь явно
        выбирает: показать папку или пересобрать из шаблона (п. 2.9).
        """
        try:
            extra = rec.get("extra") or {}
            source = str(extra.get("source_file") or "")
            if rec.get("kind") == storage.KIND_UNKNOWN:
                # «Неопознанные» не открываем через os.startfile — только папку.
                show_in_folder(source)
                return
            if extra.get("imported"):
                if source and os.path.exists(source):
                    open_file(source)
                    return
                if not self._confirm_missing_source(rec, source):
                    return
                # Пользователь разрешил пересобрать документ из шаблона.

            kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
            party_rec = self.store.get_counterparty(rec.get("party_id"))
            party_name = self.store.party_name(rec.get("party_id"))
            number = rec.get("number") or "б/н"
            default_filename = f"{kind_label} №{number} {safe_filename(party_name)}.docx"

            file_path = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".docx",
                filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")],
                initialfile=default_filename,
                initialdir=self._party_documents_dir(party_rec),
                title="Сохранить документ Word",
            )
            if not file_path:
                return

            doc = self._build_record_docx(rec)
            save_docx_atomic(doc, file_path)
            open_file(file_path)
            if extra.get("imported"):
                ui.alert(
                    self.root, "Документ собран из шаблона",
                    "Оригинал импортированного документа не найден, поэтому "
                    "документ собран из шаблона программы.\n\n"
                    "Сверьте его с оригиналом перед отправкой контрагенту.",
                    "warning",
                )
        except Exception as exc:
            self._log.warning(
                "Экспорт документа не удался: id=%s, %s",
                rec.get("id"), type(exc).__name__,
            )
            ui.alert(self.root, "Не удалось сохранить документ", str(exc), "error")

    def _confirm_missing_source(self, rec: dict, source: str) -> bool:
        """Диалог при утере оригинала импортированного документа (п. 2.9)."""
        self._log.warning(
            "Оригинал импортированного документа не найден: id=%s", rec.get("id"),
        )
        choice = self._choice_dialog(
            "Оригинал не найден",
            "Исходный файл импортированного документа отсутствует:\n"
            f"{source or 'путь не сохранён'}\n\n"
            "Можно показать папку документов или пересобрать документ из шаблона. "
            "Пересобранный документ может отличаться от оригинала — сверьте его.",
            [
                ("cancel", "Отмена", False),
                ("folder", "Показать папку", False),
                ("rebuild", "Пересобрать из шаблона", True),
            ],
        )
        if choice == "folder":
            show_in_folder(source or self._documents_dir())
            return False
        return choice == "rebuild"

    def _delete_document(self, rec: dict):
        """Перемещает документ в корзину после подтверждения (п. 2.5, 2.6)."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        party_name = self.store.party_name(rec.get("party_id"))
        source = str((rec.get("extra") or {}).get("source_file") or "")
        message = (
            f"{title}\nКонтрагент: {party_name}\n\n"
            "Переместить документ в корзину?\n\n"
            "Файл на диске останется:\n"
            f"{source or '—'}\n\n"
            "Восстановить документ можно в разделе «Корзина»."
        )
        if not ui.confirm(self.root, "Удалить документ", message, ok_text="В корзину"):
            return
        if not self.store.trash_document(rec.get("id")):
            ui.alert(
                self.root, "Не удалось удалить документ",
                self.store.last_error or "Запись не найдена.", "error",
            )
            return
        self.refresh_registry()
        self._update_status()
        self._preview_contract()
        self._preview_additional()

    def _delete_counterparty(self, rec: dict):
        """Перемещает контрагента и его документы в корзину (п. 2.5, 2.6)."""
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = self.store.documents_for(party_id)
        folder = self._party_documents_dir(rec, interactive=False)
        lines = []
        for doc in docs[:10]:
            kind_label = storage.KIND_LABELS.get(doc.get("kind"), "Документ")
            number = doc.get("number") or "б/н"
            date_text = doc.get("date") or ""
            lines.append(
                f"• {kind_label} № {number}" + (f" от {date_text}" if date_text else "")
            )
        if len(docs) > 10:
            lines.append(f"• …и ещё {len(docs) - 10}")
        listing = "\n".join(lines) if lines else "• документов нет"
        message = (
            f"Удалить контрагента «{name}»?\n\n"
            f"В корзину будут перемещены карточка и документы ({len(docs)} шт.):\n"
            f"{listing}\n\n"
            "Файлы на диске останутся в папке:\n"
            f"{folder}\n\n"
            "Восстановить можно в разделе «Корзина»."
        )
        if not ui.confirm(self.root, "Удалить контрагента", message, ok_text="В корзину"):
            return
        if len(docs) > 20:
            if not ui.confirm(
                self.root, "Подтвердите удаление",
                f"У контрагента «{name}» {len(docs)} документов.\n\n"
                "Точно переместить всё в корзину?",
                ok_text="Да, в корзину",
            ):
                return
        removed = self.store.trash_counterparty(party_id)
        if self.store.last_error:
            ui.alert(
                self.root, "Не удалось удалить контрагента",
                self.store.last_error, "error",
            )
            return
        self._log.info(
            "Контрагент перемещён в корзину: id=%s, документов=%d", party_id, removed,
        )
        if self._registry_selected_id == party_id:
            self._registry_selected_id = None
        # Списки контрагентов в выпадающих меню и предпросмотры устарели.
        self._refresh_party_combo_values()
        self._preview_contract()
        self._preview_additional()
        self.refresh_registry()
        self._update_status()

    # -- Корзина: восстановление и окончательная очистка ------------------ #
    def _restore_document(self, rec: dict):
        """Возвращает документ из корзины (п. 2.6)."""
        if not self.store.restore_document(rec.get("id")):
            ui.alert(
                self.root, "Не удалось восстановить документ",
                self.store.last_error or "Запись не найдена.", "error",
            )
            return
        self._refresh_party_combo_values()
        self._preview_contract()
        self._preview_additional()
        self.refresh_registry()
        self._update_status()

    def _purge_document(self, rec: dict):
        """Удаляет запись документа из корзины окончательно."""
        kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
        number = rec.get("number") or "б/н"
        date_text = rec.get("date") or ""
        title = f"{kind_label} № {number}" + (f" от {date_text}" if date_text else "")
        if not ui.confirm(
            self.root, "Удалить навсегда",
            f"{title}\n\nУдалить запись окончательно?\n\n"
            "Файл на диске останется, но восстановить документ из корзины "
            "будет уже нельзя.",
            ok_text="Удалить навсегда",
        ):
            return
        if not self.store.purge_document(rec.get("id")):
            ui.alert(
                self.root, "Не удалось удалить запись",
                self.store.last_error or "Запись не найдена.", "error",
            )
            return
        self.refresh_registry()
        self._update_status()

    def _restore_counterparty(self, rec: dict):
        """Возвращает карточку контрагента и удалённые вместе с ней документы."""
        restored = self.store.restore_counterparty(rec.get("id"))
        if self.store.last_error:
            ui.alert(
                self.root, "Не удалось восстановить контрагента",
                self.store.last_error, "error",
            )
            return
        self._log.info(
            "Контрагент восстановлен из корзины: id=%s, документов=%d",
            rec.get("id"), restored,
        )
        self._refresh_party_combo_values()
        self._preview_contract()
        self._preview_additional()
        self.refresh_registry()
        self._update_status()

    def _purge_counterparty(self, rec: dict):
        """Удаляет карточку контрагента и его документы окончательно."""
        party_id = rec.get("id")
        name = storage.display_name(rec)
        docs = self.store.documents_for(party_id, include_deleted=True)
        if not ui.confirm(
            self.root, "Удалить навсегда",
            f"Удалить контрагента «{name}» окончательно?\n\n"
            f"Вместе с карточкой будут удалены записи документов — {len(docs)} шт.\n"
            "Файлы на диске останутся. Восстановить записи будет нельзя.",
            ok_text="Удалить навсегда",
        ):
            return
        self.store.purge_counterparty(party_id)
        if self.store.last_error:
            ui.alert(
                self.root, "Не удалось удалить контрагента",
                self.store.last_error, "error",
            )
            return
        self._refresh_party_combo_values()
        self.refresh_registry()
        self._update_status()

    def _empty_trash(self):
        """Полностью очищает корзину (записи; файлы на диске остаются)."""
        count = self.store.trash_count()
        if not count:
            return
        if not ui.confirm(
            self.root, "Очистить корзину",
            f"В корзине записей: {count}.\n\n"
            "Удалить их окончательно? Файлы документов на диске останутся, "
            "но записи восстановить будет нельзя.",
            ok_text="Очистить корзину",
        ):
            return
        removed = self.store.purge_all_trash()
        if not removed and self.store.last_error:
            ui.alert(
                self.root, "Не удалось очистить корзину",
                self.store.last_error, "error",
            )
            return
        self._refresh_party_combo_values()
        self.refresh_registry()
        self._update_status()

    def _edit_counterparty(self, rec: dict):
        """Окно редактирования контрагента: поля карточки + «Сохранить»/«Отмена».

        Изменяется только карточка в реестре. Ранее созданные документы не
        затрагиваются: каждый документ хранит собственную копию реквизитов
        на момент создания.
        """
        mode = "ip" if rec.get("type") == "ip" else "ooo"
        is_ip = mode == "ip"
        saved = rec.get("fields", {})

        win = ctk.CTkToplevel(self.root)
        win.title(f"Редактирование контрагента — {storage.display_name(rec)}")
        win.geometry(f"{int(760 * self.scale)}x{int(680 * self.scale)}")
        win.minsize(int(660 * self.scale), int(520 * self.scale))
        win.transient(self.root)
        win.after(120, win.grab_set)

        header = ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * self.scale), pady=(int(18 * self.scale), int(6 * self.scale)))
        ctk.CTkLabel(
            header,
            text=f"Редактирование контрагента · {'ИП' if is_ip else 'Юр. лицо'}",
            font=ui.font(18, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=24,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Изменится только карточка в реестре. Ранее созданные документы останутся без изменений.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * self.scale), 0))

        scroller = ctk.CTkScrollableFrame(
            win, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        scroller.pack(fill="both", expand=True, padx=int(24 * self.scale), pady=(int(4 * self.scale), int(8 * self.scale)))
        self.scrollers.append(scroller)

        card = ui.Card(scroller, compact=True)
        card.pack(fill="x", pady=(0, int(8 * self.scale)))

        entries: dict[str, tk.Widget] = {}
        error_vars: dict[str, tk.StringVar] = {}

        def error_var(key: str) -> tk.StringVar:
            return error_vars.setdefault(key, tk.StringVar(value=""))

        def prefill(key: str, widget):
            value = str(saved.get(key, "")).strip()
            if not value:
                return
            try:
                widget.delete(0, tk.END)
                widget.insert(0, value)
            except Exception:
                try:
                    widget.set(value)
                except Exception:
                    pass

        def add_field(key, label, kind="text", width=None, required=False,
                      stretch=True, max_len=None, hint=None):
            width = width or int(ENTRY_WIDTH * self.scale)
            validate = None
            if kind == "digits":
                validate = self._digits_vcmd_for(max_len) if max_len else (self._digits_vcmd, "%P")
            elif kind == "phone":
                validate = (self._phone_vcmd, "%P")
            widget = ui.entry(card.body, width=width, validatecommand=validate)
            if kind == "digits":
                # Валидатор отбрасывает вставку целиком, если в буфере есть
                # недопустимые символы. Вставляем только цифры.
                ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
            elif kind == "phone":
                phone_var = tk.StringVar()
                widget.entry.configure(textvariable=phone_var)
                phone_var.trace_add("write", make_phone_formatter(phone_var, widget.entry))
                ui.enable_sanitized_paste(
                    widget.entry,
                    lambda s: "".join(ch for ch in str(s) if ch.isdigit() or ch in "+-()/. "),
                )
            prefill(key, widget)
            entries[key] = widget
            card.row(label, widget, stretch=stretch, required=required,
                     error_var=error_var(key), hint=hint)
            return widget

        # -- Коды -------------------------------------------------------- #
        card.section("Коды")
        add_field(F.INN, F.INN, kind="digits", max_len=12,
                  width=int(180 * self.scale), stretch=False, required=True)
        if is_ip:
            add_field(F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                      width=int(180 * self.scale), stretch=False, required=True)
        else:
            add_field(F.KPP, F.KPP, kind="digits", max_len=9,
                      width=int(140 * self.scale), stretch=False, required=True)
            add_field(F.OGRN, F.OGRN, kind="digits", max_len=13,
                      width=int(180 * self.scale), stretch=False, required=True)

        # -- Организация / Предприниматель -------------------------------- #
        card.section("Организация" if not is_ip else "Предприниматель")
        if is_ip:
            add_field(F.IP_NAME, F.IP_NAME, required=True)
        else:
            add_field(F.FULL_NAME, "Полное название", required=True)
            add_field(F.SHORT_NAME, "Сокращённое название", required=True)

        # -- Руководитель ------------------------------------------------- #
        if not is_ip:
            card.section("Руководитель")
            add_field(F.DIRECTOR, F.DIRECTOR, required=True)
            widget = ui.option_menu(card.body, ["Устав", "Доверенность"],
                                    width=int(ENTRY_WIDTH * self.scale))
            basis = str(saved.get(F.BASIS, "")).strip()
            widget.set(basis if basis in ("Устав", "Доверенность") else "Устав")
            entries[F.BASIS] = widget
            card.row(F.BASIS, widget)

        # -- Адреса ------------------------------------------------------- #
        card.section("Адреса")
        add_field(F.LEGAL_ADDRESS, F.LEGAL_ADDRESS, required=True)
        if is_ip:
            card.row(
                F.FACT_ADDRESS,
                ctk.CTkLabel(
                    card.body, text="как юридический", font=ui.font(13),
                    text_color=ui.COLORS["text_tertiary"], anchor="w",
                ),
            )
        else:
            add_field(F.FACT_ADDRESS, F.FACT_ADDRESS)

        # -- Контакты ----------------------------------------------------- #
        # Email и телефон необязательны (п. 4.5).
        card.section("Контакты")
        add_field(F.EMAIL, F.EMAIL, kind="email", hint="Необязательно.")
        add_field(F.PHONE, F.PHONE, kind="phone", hint="Необязательно.")

        # -- Банк ---------------------------------------------------------- #
        card.section("Банковские реквизиты")
        add_field(F.ACCOUNT, F.ACCOUNT, kind="digits", max_len=20,
                  width=int(260 * self.scale), required=True)
        add_field(F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                  width=int(260 * self.scale), required=True)
        add_field(F.BANK, F.BANK, required=True)
        add_field(F.BIK, F.BIK, kind="digits", max_len=9,
                  width=int(140 * self.scale), stretch=False, required=True)

        def highlight_errors(errors: dict[str, str]):
            for var in error_vars.values():
                var.set("")
            for widget in entries.values():
                if hasattr(widget, "set_error"):
                    widget.set_error(False)
            for key, message in errors.items():
                error_var(key).set(message)
                widget = entries.get(key)
                if widget is not None and hasattr(widget, "set_error"):
                    widget.set_error(True)
            first = entries.get(next(iter(errors)))
            if first is not None:
                try:
                    first.focus_set()
                except Exception:
                    pass

        def save():
            data = {}
            for key, widget in entries.items():
                try:
                    data[key] = widget.get().strip()
                except Exception:
                    data[key] = ""
            if is_ip:
                data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")

            errors = validate_party_fields(data, is_ip)
            # ИНН должен оставаться уникальным в реестре.
            duplicate = self.store.find_by_inn(data.get(F.INN, "").strip())
            if duplicate is not None and duplicate.get("id") != rec.get("id"):
                errors.setdefault(F.INN, "Контрагент с таким ИНН уже есть в реестре")
            if errors:
                highlight_errors(errors)
                ui.alert(
                    win, "Ошибки в данных",
                    "Проверьте подсвеченные поля:\n\n"
                    + "\n".join(f"• {message}" for message in errors.values()),
                    "error",
                )
                return

            # Обновляем только карточку контрагента. Документы не трогаем:
            # они хранят собственную копию реквизитов на момент создания.
            updated = self.store.update_counterparty(rec["id"], fields=data)
            if self.store.last_error:
                ui.alert(
                    win, "Не удалось сохранить изменения",
                    self.store.last_error, "error",
                )
                return
            name = storage.display_name(updated or rec)
            # Переименование не должно «терять» выбранного контрагента (п. 4.2).
            if self._current_party_id == rec.get("id"):
                self._set_current_party(updated or rec)
            else:
                self._update_current_party_labels()
            self._refresh_party_combo_values()
            self.refresh_registry()
            self._preview_contract()
            self._preview_additional()
            self._update_status()
            win.destroy()
            self._notify(
                f"Карточка «{name}» обновлена. Ранее созданные документы не изменялись."
            )

        footer = ctk.CTkFrame(win, fg_color="transparent")
        footer.pack(fill="x", padx=int(24 * self.scale), pady=(0, int(16 * self.scale)))
        ui.secondary_button(
            footer, "Отмена", win.destroy,
            width=int(140 * self.scale), height=int(30 * self.scale),
        ).pack(side="right")
        ui.primary_button(
            footer, "Сохранить", save,
            width=int(170 * self.scale), height=int(30 * self.scale),
        ).pack(side="right", padx=(0, int(10 * self.scale)))

    # ------------------------------------------------------------------ #
    #  Страница «Настройки»
    # ------------------------------------------------------------------ #
    def _build_settings_page(self):
        _page, left, _right, _textbox = self._make_page(
            "settings", "Настройки",
            "Тема оформления, папка документов и ключ DaData.",
        )

        # -- Тема -------------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Тема оформления")
        theme_toggle = ui.ThemeToggle(
            card.body,
            command=self._on_theme_change,
            mode="dark" if ctk.get_appearance_mode().lower() == "dark" else "light",
        )
        card.wide(theme_toggle, pady=(int(2 * self.scale), int(6 * self.scale)))
        # Переключатель — по ширине двух иконок, не растягиваем на всю карточку.
        theme_toggle.grid_configure(sticky="w")

        # -- Папка документов -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("Папка для сохранения документов")
        self._documents_dir_var = tk.StringVar(value=self._documents_dir())
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=self._documents_dir_var,
                font=ui.font(12), text_color=ui.COLORS["text"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 48) * self.scale),
            ),
            pady=(int(2 * self.scale), int(6 * self.scale)),
        )
        buttons = ctk.CTkFrame(card.body, fg_color="transparent")
        card.wide(buttons)
        ui.secondary_button(
            buttons, "Выбрать папку…", self._choose_documents_dir,
            width=int(170 * self.scale), height=int(30 * self.scale),
        ).pack(side="left")
        ui.secondary_button(
            buttons, "Открыть папку", lambda: open_file(self._documents_dir()),
            width=int(150 * self.scale), height=int(30 * self.scale),
        ).pack(side="left", padx=(int(10 * self.scale), 0))
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Здесь сохраняются договоры и доп. соглашения, созданные в программе. "
                     "Для каждого контрагента заводится своя папка с его документами.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 48) * self.scale),
            ),
            pady=(int(4 * self.scale), int(2 * self.scale)),
        )

        # Добавление документов перенесено в раздел «Реестр» (п. 4.5).

        # -- DaData ------------------------------------------------------ #
        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new", padx=int(4 * self.scale), pady=(0, int(12 * self.scale)))
        card.section("DaData: автозаполнение по ИНН")
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Ключ не входит в поставку — укажите свой ключ с dadata.ru. "
                     "Без ключа доступен только ручной ввод реквизитов.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 48) * self.scale),
            ),
            pady=(int(2 * self.scale), int(6 * self.scale)),
        )
        self._dadata_key_entry = ui.entry(card.body, width=int(ENTRY_WIDTH * self.scale))
        try:
            self._dadata_key_entry.entry.configure(show="•")
        except Exception:
            pass
        current_key = dadata.load_api_key()
        if current_key:
            self._dadata_key_entry.insert(0, current_key)
        card.row(
            "API-ключ", self._dadata_key_entry, stretch=True,
            hint="Ключ хранится в профиле пользователя отдельно от реестра.",
        )
        key_buttons = ctk.CTkFrame(card.body, fg_color="transparent")
        card.wide(key_buttons)
        ui.secondary_button(
            key_buttons, "Сохранить ключ", self._save_dadata_key,
            width=int(160 * self.scale), height=int(30 * self.scale),
        ).pack(side="left")
        ui.ghost_button(
            key_buttons, "Удалить ключ", self._clear_dadata_key,
            width=int(130 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(10 * self.scale), 0))

    def _save_dadata_key(self):
        """Сохраняет введённый пользователем ключ DaData."""
        key = self._dadata_key_entry.get().strip()
        if not key:
            ui.alert(self.root, "Ключ не указан",
                     "Введите API-ключ DaData или оставьте поле пустым.", "error")
            return
        if dadata.save_api_key(key):
            ui.alert(
                self.root, "Ключ сохранён",
                "Автозаполнение по ИНН будет использовать этот ключ.", "success",
            )
        else:
            ui.alert(
                self.root, "Не удалось сохранить ключ",
                "Проверьте права на запись в профиль пользователя:\n"
                f"{dadata.key_storage_path()}", "error",
            )

    def _clear_dadata_key(self):
        """Удаляет сохранённый ключ DaData."""
        if not ui.confirm(
            self.root, "Удалить ключ DaData",
            "Удалить сохранённый ключ? Автозаполнение по ИНН перестанет работать.",
            ok_text="Удалить",
        ):
            return
        dadata.clear_api_key()
        try:
            self._dadata_key_entry.delete(0, "end")
        except Exception:
            pass
        ui.alert(self.root, "Ключ удалён",
                 "Ключ DaData удалён из профиля пользователя.", "success")

    def _choose_documents_dir(self):
        current = self._documents_dir()
        chosen = filedialog.askdirectory(
            parent=self.root, initialdir=current,
            title="Папка для сохранения договоров и документов",
        )
        if not chosen:
            return
        self._settings["documents_dir"] = chosen
        self._documents_dir_var.set(chosen)
        # Сохраняем сразу, а не только при закрытии: иначе при аварийном
        # завершении выбранная папка теряется.
        if self._persist_settings():
            ui.alert(
                self.root, "Папка сохранена",
                f"Документы будут сохраняться в:\n{chosen}", "success",
            )

    # ------------------------------------------------------------------ #
    #  Добавление готовых документов (фоновый импорт)
    # ------------------------------------------------------------------ #
    def _import_documents_dialog(self):
        """Выбор файлов для загрузки в реестр."""
        if getattr(self, "_import_running", False):
            return
        files = filedialog.askopenfilenames(
            parent=self.root,
            title="Выберите договоры, доп. соглашения и спецификации",
            filetypes=[
                ("Документы Word и Excel", "*.docx *.xlsx"),
                ("Документы Word", "*.docx"),
                ("Таблицы Excel", "*.xlsx"),
            ],
        )
        if not files:
            return
        self._start_documents_import(list(files))

    def _start_documents_import(self, files: list[str]):
        """Проверяет лимиты пакета и запускает фоновый анализ файлов.

        До запуска отклоняются: слишком большой пакет, слишком большой
        суммарный объём и ситуация «на диске не хватит места» (п. 3.4).
        Размер каждого файла дополнительно проверяется в analyze_file().
        """
        files = [str(path) for path in files if path]
        if not files:
            return
        if len(files) > IMPORT_MAX_FILES:
            ui.alert(
                self.root, "Слишком много файлов",
                f"За один раз можно добавить не более {IMPORT_MAX_FILES} файлов, "
                f"выбрано {len(files)}.\n\n"
                "Разбейте загрузку на несколько частей.",
                "warning",
            )
            return

        total_size = 0
        for path in files:
            try:
                total_size += os.path.getsize(path)
            except OSError:
                pass
        if total_size > IMPORT_MAX_TOTAL_SIZE:
            ui.alert(
                self.root, "Слишком большой пакет",
                f"Суммарный объём файлов {total_size / (1024 * 1024):.0f} МБ "
                f"превышает допустимые {IMPORT_MAX_TOTAL_SIZE / (1024 * 1024):.0f} МБ.\n\n"
                "Разбейте загрузку на несколько частей.",
                "warning",
            )
            return

        try:
            target_dir = self._documents_dir_for_import()
        except OSError as exc:
            ui.alert(self.root, "Папка документов недоступна", str(exc), "error")
            return
        try:
            free = shutil.disk_usage(target_dir).free
        except OSError:
            free = None
        if free is not None and free < total_size + IMPORT_MIN_FREE_SPACE:
            ui.alert(
                self.root, "Недостаточно места на диске",
                "Для импорта нужно около "
                f"{(total_size + IMPORT_MIN_FREE_SPACE) / (1024 * 1024):.0f} МБ "
                f"свободного места, доступно {free / (1024 * 1024):.0f} МБ.\n\n"
                "Освободите место или выберите другую папку документов в «Настройках».",
                "warning",
            )
            return

        self._import_running = True
        self._import_queue = queue.Queue()
        self._import_results: list[dict] = []
        self._import_total = max(1, len(files))
        self._import_cancelled = False
        self._import_cancelled_count = 0

        win = ctk.CTkToplevel(self.root)
        win.title("Загрузка документов")
        win.resizable(False, False)
        win.transient(self.root)
        self._import_win = win

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.pack(fill="both", expand=True,
                   padx=int(28 * self.scale), pady=int(22 * self.scale))
        ctk.CTkLabel(
            frame, text="Загрузка документов", font=ui.font(16, "bold"),
            text_color=ui.COLORS["text"], anchor="w",
        ).pack(anchor="w")
        self._import_status_var = tk.StringVar(value="Подготовка…")
        ctk.CTkLabel(
            frame, textvariable=self._import_status_var, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", justify="left",
            wraplength=int(360 * self.scale),
        ).pack(anchor="w", pady=(int(8 * self.scale), int(12 * self.scale)))
        self._import_progress = ctk.CTkProgressBar(frame, width=int(360 * self.scale))
        self._import_progress.set(0)
        self._import_progress.pack(fill="x")
        ctk.CTkLabel(
            frame,
            text="Документы анализируются в фоне. Не закрывайте окно программы.",
            font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
            anchor="w", justify="left", wraplength=int(360 * self.scale),
        ).pack(anchor="w", pady=(int(10 * self.scale), 0))
        self._import_cancel_btn = ui.ghost_button(
            frame, "Отменить", self._cancel_import,
            width=int(120 * self.scale), height=int(28 * self.scale),
        )
        self._import_cancel_btn.pack(anchor="e", pady=(int(8 * self.scale), 0))

        # Окно по центру родительского, поверх остальных.
        win.update_idletasks()
        width = int(430 * self.scale)
        height = int(240 * self.scale)
        x = self.root.winfo_x() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - height) // 2
        win.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
        try:
            win.grab_set()
        except tk.TclError:
            pass

        threading.Thread(target=self._import_worker, args=(files,), daemon=True).start()
        self.root.after(80, self._import_tick)

    def _cancel_import(self):
        """Отменяет импорт: остаток пакета не анализируется и не регистрируется (п. 5.3)."""
        if not getattr(self, "_import_running", False):
            return
        self._import_cancelled = True
        try:
            self._import_status_var.set("Отмена…")
            self._import_cancel_btn.configure(state="disabled")
        except Exception:
            pass

    def _import_worker(self, files: list[str]):
        """Фоновый поток: анализ каждого файла и копирование в папку контрагента.

        После нажатия «Отменить» остаток пакета не анализируется (п. 5.3).
        """
        for index, path in enumerate(files):
            if self._import_cancelled:
                break
            name = os.path.basename(path)
            self._import_queue.put(
                ("status", f"Анализ файла {index + 1} из {len(files)}:\n{name}")
            )
            ext = os.path.splitext(path)[1].lower()
            if ext not in IMPORT_ALLOWED_EXTENSIONS:
                # Allowlist: .exe/.bat/.lnk и прочее не анализируем и не копируем.
                self._import_queue.put(("file", {
                    "path": path,
                    "name": name,
                    "kind": storage.KIND_UNKNOWN,
                    "number": "",
                    "date": "",
                    "party": None,
                    "text": "",
                    "error": f"Тип файла не поддерживается: {ext or 'без расширения'}",
                    "warning": "",
                    "recognized": False,
                    "ok": False,
                    "copied": "",
                }))
                continue
            result = importer.analyze_file(path)
            try:
                result["copied"] = self._copy_imported_file(result)
            except Exception as exc:
                result["copied"] = ""
                if not result.get("error"):
                    result["error"] = f"Не удалось скопировать файл: {exc}"
            self._import_queue.put(("file", result))
        self._import_queue.put(("done",))

    def _copy_imported_file(self, result: dict) -> str:
        """Копирует в папку контрагента только разобранные .docx/.xlsx.

        Неопознанные и потенциально опасные файлы не копируются: раньше они
        попадали в папку документов и открывались по клику (запуск кода).
        """
        if not result.get("ok") or not result.get("party"):
            return ""
        src = str(result.get("path") or "")
        ext = os.path.splitext(src)[1].lower()
        if ext not in IMPORT_ALLOWED_EXTENSIONS:
            return ""
        base = self._documents_dir_for_import()
        folder = safe_filename(
            importer.party_display_name(result["party"]), "Контрагент"
        )
        dest_dir = os.path.normpath(os.path.join(base, folder))
        if not self._is_inside_directory(base, dest_dir):
            self._log.error("Папка импортированного контрагента вне папки документов")
            return ""
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except OSError as exc:
            self._log.warning(
                "Не удалось создать папку для импорта: %s", type(exc).__name__,
            )
            raise

        stem = safe_filename(os.path.splitext(os.path.basename(src))[0], "Документ")
        dest = self._fit_document_path(dest_dir, stem, ext)
        counter = 2
        while os.path.exists(dest):
            dest = self._fit_document_path(dest_dir, f"{stem} ({counter})", ext)
            counter += 1
        copy_with_zone_identifier(src, dest)
        try:
            size = os.path.getsize(dest)
        except OSError:
            size = 0
        self._log.info(
            "Импортированный файл скопирован: тип=%s, размер=%d байт", ext, size,
        )
        return dest

    def _import_tick(self):
        """Опрос очереди фонового потока из главного потока Tk."""
        try:
            while True:
                item = self._import_queue.get_nowait()
                kind = item[0]
                if kind == "status":
                    self._import_status_var.set(item[1])
                elif kind == "file":
                    self._import_results.append(item[1])
                    self._import_progress.set(
                        len(self._import_results) / self._import_total
                    )
                elif kind == "done":
                    results = self._import_results
                    win = self._import_win
                    self._import_queue = None
                    self._import_results = []
                    self._import_win = None
                    self._import_progress = None
                    self._import_running = False
                    # Часть пакета могла быть пропущена по «Отмене» (п. 5.3).
                    self._import_cancelled_count = max(
                        0, self._import_total - len(results)
                    )
                    try:
                        win.grab_release()
                        win.destroy()
                    except Exception:
                        pass
                    self._finish_documents_import(results)
                    return
        except queue.Empty:
            pass
        except (tk.TclError, AttributeError):
            return  # окно программы закрывается — импорт прекращаем
        self.root.after(80, self._import_tick)

    def _finish_documents_import(self, results: list[dict]):
        """Регистрирует результаты анализа в реестре и показывает отчёт.

        Правила (п. 3.3, 3.5):
          * карточка создаётся автоматически только при найденном ИНН;
          * без ИНН — диалог «привязать к существующему / создать / пропустить»;
          * документ, который уже есть в реестре, повторно не добавляется;
          * по каждому файлу в отчёте видны тип, результат и причина.

        Пакетная запись (п. 5.2): все изменения пакета пишутся на диск ОДНИМ
        ``save()`` в конце. Раньше каждое «добавить» полностью перезаписывало
        JSON с fsync и ротацией копий — импорт 200 файлов давал 200 полных
        записей реестра.
        """
        stats = {
            "added": 0,
            "new_parties": 0,
            "skipped": 0,
            "duplicates": 0,
            "cancelled": max(0, int(getattr(self, "_import_cancelled_count", 0))),
        }
        save_errors: list[str] = []
        entries: list[dict] = []
        store_dirty = False

        for result in results:
            entry = self._new_import_entry(result)
            entries.append(entry)
            if self._register_import_result(result, entry, stats):
                store_dirty = True

        if store_dirty:
            with applog.timed("Запись реестра (импорт)", 500, "storage"):
                if not self.store.save():
                    save_errors.append(self.store.last_error or "реестр не записан")
            self._log.info("Пакет импорта записан одним save() (п. 5.2)")

        if stats["added"]:
            self._refresh_party_combo_values()
            self._preview_contract()
            self._preview_additional()
            self.refresh_registry()
            self._update_status()

        message = (
            "Загрузка завершена.\n"
            f"Добавлено документов: {stats['added']} (новых контрагентов: {stats['new_parties']})\n"
            f"Пропущено: {stats['skipped']}"
        )
        if stats["duplicates"]:
            message += f", уже были в реестре: {stats['duplicates']}"
        if stats["cancelled"]:
            message += f"\nОтменено пользователем: {stats['cancelled']}"
        if save_errors:
            message += (
                "\n\nВНИМАНИЕ: часть записей не попала на диск:\n"
                + "\n".join(f"• {err}" for err in dict.fromkeys(save_errors))
            )
        self._log.info(
            "Импорт завершён: добавлено=%d, новых контрагентов=%d, "
            "пропущено=%d, дубликатов=%d, ошибок записи=%d",
            stats["added"], stats["new_parties"], stats["skipped"],
            stats["duplicates"], len(save_errors),
        )
        self._show_import_report(entries, message)
        self._import_cancelled_count = 0

    @staticmethod
    def _new_import_entry(result: dict) -> dict:
        """Заготовка строки отчёта импорта для одного файла."""
        return {
            "name": result.get("name") or "файл",
            "kind": result.get("kind") or storage.KIND_UNKNOWN,
            "status": "skip",
            "detail": "",
            "warning": "",
        }

    def _register_import_result(self, result: dict, entry: dict, stats: dict) -> bool:
        """Регистрирует один результат разбора. Возвращает True, если реестр менялся.

        Изменения остаются в памяти: сохранение на диск делает вызывающий код
        одним ``save()`` на весь пакет (п. 5.2).
        """
        copied = result.get("copied") or ""
        party = result.get("party") or {}

        if not result.get("ok") or not party:
            stats["skipped"] += 1
            entry["detail"] = result.get("error") or "документ не распознан"
            result["copied"] = self._move_import_to_unknown(copied)
            return False

        fields = dict(party.get("fields") or {})
        inn = str(fields.get(F.INN, "")).strip()
        rec = None
        action = ""
        if result.get("recognized") and inn:
            rec = self.store.find_by_inn(inn)
            if rec is not None:
                action = "attach"
            else:
                # ИНН новый, но карточка с таким названием уже может быть
                # (например, заведена вручную без ИНН) — не плодим дубль.
                same_name = self.store.find_by_name(party.get("display"))
                if same_name is not None:
                    attach = ui.confirm(
                        self.root, "Похожее название уже есть",
                        f"Файл: {result.get('name')}\n"
                        f"Распознано: {party.get('display')}\n\n"
                        f"В реестре есть «{storage.display_name(same_name)}» "
                        "с другим ИНН.\n\n"
                        "Привязать документ к этой карточке?",
                        "Привязать", "Создать новую",
                    )
                    if attach:
                        rec, action = same_name, "attach"
                if action != "attach":
                    action = "create"
        else:
            rec, action = self._ask_import_party(result)

        if action == "create":
            rec = self.store.add_counterparty(party.get("type") or "ooo", fields, save=False)
            stats["new_parties"] += 1
            entry["detail"] = "создана новая карточка"
        elif action == "attach" and rec is not None:
            entry["detail"] = f"привязан к «{storage.display_name(rec)}»"

        if action == "skip" or rec is None:
            stats["skipped"] += 1
            entry["detail"] = entry["detail"] or "контрагент не выбран"
            result["copied"] = self._move_import_to_unknown(copied)
            return False

        if not copied:
            stats["skipped"] += 1
            entry["detail"] = result.get("error") or "файл не скопирован"
            return False

        if self._import_document_is_duplicate(rec["id"], result):
            stats["duplicates"] += 1
            entry["status"] = "duplicate"
            entry["detail"] = "такой документ уже есть в реестре"
            result["copied"] = self._move_import_to_unknown(copied)
            return False

        self.store.add_document(
            result.get("kind") or storage.KIND_UNKNOWN,
            rec["id"],
            result.get("number") or "",
            result.get("date") or "",
            data=dict(rec.get("fields") or {}),
            extra={
                "imported": True,
                "source_file": copied,
                "is_ip": rec.get("type") == "ip",
            },
            text=result.get("text") or "",
            save=False,
        )
        stats["added"] += 1
        entry["status"] = "added"
        if not result.get("recognized"):
            entry["warning"] = (
                result.get("warning") or "карточка создана без ИНН"
            )
        return True

    def _ask_import_party(self, result: dict) -> tuple[dict | None, str]:
        """Спрашивает, к какому контрагенту отнести документ без ИНН (п. 3.3).

        Возвращает ``(запись, действие)``: действие — ``attach`` / ``create`` /
        ``skip``. Сначала предлагается найденный по названию контрагент,
        затем — выбранный на страницах документов, затем создание карточки.
        """
        party = result.get("party") or {}
        fields = party.get("fields") or {}
        display = str(party.get("display") or "").strip() or "название не распознано"
        missing = party.get("missing") or "ИНН"

        existing = self.store.find_by_name(display) if display else None
        if existing is not None:
            existing_inn = str(
                (existing.get("fields") or {}).get(F.INN, "")
            ).strip() or "не указан"
            attach = ui.confirm(
                self.root, "Похожий контрагент уже есть",
                f"Файл: {result.get('name')}\n"
                f"Распознано: {display}\n\n"
                f"В реестре есть «{storage.display_name(existing)}» "
                f"(ИНН {existing_inn}).\n\n"
                "Привязать документ к этой карточке?",
                "Привязать", "Создать новую",
            )
            if attach:
                return existing, "attach"
            choice = self._choice_dialog(
                "Новый контрагент",
                f"Создать новую карточку «{display}» без {missing}?",
                [("create", "Создать карточку", True),
                 ("skip", "Пропустить", False)],
            )
            return (None, "create") if choice == "create" else (None, "skip")

        current = self._current_import_party()
        options: list[tuple[str, str, bool]] = []
        if current is not None:
            options.append((
                "attach", f"Привязать к «{storage.display_name(current)}»", False,
            ))
        options.append(("create", "Создать карточку", True))
        options.append(("skip", "Пропустить", False))
        choice = self._choice_dialog(
            "Не найден ИНН",
            f"Файл: {result.get('name')}\n"
            f"Распознано: {display}\n\n"
            f"Не найден {missing}: автоматически карточку создать нельзя.\n"
            "Выберите, что сделать с документом:",
            options,
        )
        if choice == "attach":
            return current, "attach"
        if choice == "create":
            return None, "create"
        return None, "skip"

    def _current_import_party(self) -> dict | None:
        """Контрагент, выбранный на страницах договора/доп./спецификации."""
        for attr in ("contract_party_var", "addl_party_var", "spec_party_var"):
            var = getattr(self, attr, None)
            if var is None:
                continue
            rec = self._party_from_var(var)
            if rec is not None:
                return rec
        return None

    def _import_document_is_duplicate(self, party_id, result: dict) -> bool:
        """Такой документ уже есть у контрагента (тип + номер + дата)."""
        kind = result.get("kind") or ""
        number = str(result.get("number") or "").strip()
        date_text = str(result.get("date") or "").strip()
        if not kind or (not number and not date_text):
            return False  # без номера и даты сравнивать нечего
        for rec in self.store.documents_for(party_id):
            if rec.get("kind") != kind:
                continue
            if str(rec.get("number") or "").strip() != number:
                continue
            if str(rec.get("date") or "").strip() != date_text:
                continue
            return True
        return False

    def _move_import_to_unknown(self, path: str) -> str:
        """Переносит отклонённый файл в папку «unknown» (п. 3.5).

        Документ, который не попал в реестр, не должен оставаться в папке
        контрагента — иначе он выглядит как осиротевший файл.
        """
        if not path:
            return path
        try:
            if not os.path.exists(path):
                return path
            base = self._documents_dir_for_import()
            unknown_dir = os.path.join(base, importer.UNKNOWN_FOLDER)
            os.makedirs(unknown_dir, exist_ok=True)
            ext = os.path.splitext(path)[1]
            stem = safe_filename(os.path.splitext(os.path.basename(path))[0], "Документ")
            dest = self._fit_document_path(unknown_dir, stem, ext)
            counter = 2
            while os.path.exists(dest):
                dest = self._fit_document_path(
                    unknown_dir, f"{stem} ({counter})", ext
                )
                counter += 1
            os.replace(path, dest)
            return dest
        except OSError as exc:
            self._log.warning(
                "Не удалось перенести файл в unknown: %s", type(exc).__name__,
            )
            return path

    def _show_import_report(self, entries: list[dict], summary: str) -> None:
        """Окно отчёта: по каждому файлу тип, результат и причина (п. 3.5)."""
        win = ctk.CTkToplevel(self.root)
        win.title("Отчёт об импорте")
        win.configure(fg_color=ui.COLORS["window"])
        win.transient(self.root)
        win.resizable(True, True)

        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=int(22 * self.scale),
                  pady=int(18 * self.scale))
        ctk.CTkLabel(
            body, text="Отчёт об импорте", font=ui.font(16, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w",
        ).pack(anchor="w")
        ctk.CTkLabel(
            body, text=summary, font=ui.font(12), justify="left", anchor="w",
            text_color=ui.COLORS["text_secondary"], wraplength=int(660 * self.scale),
        ).pack(anchor="w", pady=(int(6 * self.scale), int(10 * self.scale)))

        textbox = ctk.CTkTextbox(
            body, width=int(660 * self.scale), height=int(360 * self.scale),
            font=ui.font(12), wrap="word",
        )
        textbox.pack(fill="both", expand=True)

        icons = {"added": "✓", "duplicate": "•", "skip": "✕"}
        lines: list[str] = []
        for entry in entries:
            label = storage.KIND_LABELS.get(entry.get("kind"), "Документ")
            line = (
                f"{icons.get(entry.get('status'), '•')} {entry.get('name')} — "
                f"{label}"
            )
            if entry.get("detail"):
                line += f": {entry['detail']}"
            if entry.get("warning"):
                line += f" ({entry['warning']})"
            lines.append(line)
        textbox.insert("1.0", "\n".join(lines) or "Файлы не выбраны")
        textbox.configure(state="disabled")

        buttons = ctk.CTkFrame(body, fg_color="transparent")
        buttons.pack(fill="x", pady=(int(12 * self.scale), 0))
        ui.primary_button(
            buttons, "Закрыть", win.destroy,
            width=int(120 * self.scale), height=int(32 * self.scale),
        ).pack(side="right")

        win.update_idletasks()
        width = int(720 * self.scale)
        height = int(520 * self.scale)
        try:
            x = self.root.winfo_rootx() + max(
                0, (self.root.winfo_width() - width) // 2
            )
            y = self.root.winfo_rooty() + max(
                0, (self.root.winfo_height() - height) // 3
            )
            win.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            pass
        win.protocol("WM_DELETE_WINDOW", win.destroy)
        try:
            win.grab_set()
        except tk.TclError:
            pass
        win.wait_window()

    # ------------------------------------------------------------------ #
    #  Окно «Добавить контрагента»
    # ------------------------------------------------------------------ #
    def _build_counterparty_page(self):
        """Страница «Добавить контрагента» в основном окне программы."""
        page = ctk.CTkFrame(self.content, fg_color="transparent")
        page.grid(row=0, column=0, sticky="nsew")
        self.pages["counterparty"] = page

        header = ctk.CTkFrame(page, fg_color="transparent")
        header.pack(fill="x", padx=int(24 * self.scale), pady=(int(18 * self.scale), int(6 * self.scale)))
        ctk.CTkLabel(
            header, text="Добавить контрагента", font=ui.font(20, "bold", display=True),
            text_color=ui.COLORS["text"], anchor="w", height=26,
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text="Заполните карточку — она сохранится в реестре программы.",
            font=ui.font(12), text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(2 * self.scale), 0))

        self._cp_type_var = tk.StringVar(value="Юр. лицо")
        segmented = ui.segmented(
            header, ["Юр. лицо", "ИП"], command=self._on_cp_type_change,
            width=int(240 * self.scale),
        )
        segmented.set("Юр. лицо")
        segmented.pack(anchor="w", pady=(int(8 * self.scale), 0))

        ctk.CTkLabel(
            header, textvariable=self._cp_status_var, font=ui.font(12),
            text_color=ui.COLORS["text_secondary"], anchor="w", height=16,
        ).pack(anchor="w", pady=(int(4 * self.scale), 0))

        scroller = ctk.CTkScrollableFrame(
            page, fg_color="transparent", corner_radius=0,
            scrollbar_button_color=ui.COLORS["scroll"],
            scrollbar_button_hover_color=ui.COLORS["scroll_hover"],
        )
        scroller.pack(fill="both", expand=True, padx=int(24 * self.scale), pady=(int(4 * self.scale), int(12 * self.scale)))
        self.scrollers.append(scroller)

        self._cp_frames: dict[str, ctk.CTkFrame] = {}
        for mode in ("ooo", "ip"):
            frame = ctk.CTkFrame(scroller, fg_color="transparent")
            frame.grid(row=0, column=0, sticky="nsew")
            card = ui.Card(frame, compact=True)
            card.pack(fill="x", pady=(0, int(12 * self.scale)))
            self._build_cp_fields(card, mode)
            self._cp_frames[mode] = frame

        self._on_cp_type_change("Юр. лицо")

    def _build_cp_fields(self, card, mode: str):
        error_var = self._cp_error_var_for(mode)

        # -- Коды -------------------------------------------------------- #
        card.section("Коды")
        inn_holder = ctk.CTkFrame(card.body, fg_color="transparent")
        inn = ui.entry(
            inn_holder, width=int(180 * self.scale),
            validatecommand=self._digits_vcmd_for(12),
        )
        ui.enable_sanitized_paste(inn.entry, digits_only, max_len=12)
        inn.pack(side="left")
        self._cp_entries[mode][F.INN] = inn
        ui.ghost_button(
            inn_holder, "Заполнить по ИНН", self._cp_fill_from_inn,
            width=int(152 * self.scale), height=int(28 * self.scale),
        ).pack(side="left", padx=(int(12 * self.scale), 0))
        card.row(
            "ИНН *", inn_holder, stretch=True,
            hint="Кнопка подтягивает название, адрес, директора и банк из ЕГРЮЛ/ЕГРИП.",
            error_var=error_var(F.INN),
        )

        if mode == "ooo":
            self._cp_add_field(card, mode, F.KPP, F.KPP, kind="digits", max_len=9,
                               width=int(140 * self.scale), required=True, stretch=False)
            self._cp_add_field(card, mode, F.OGRN, F.OGRN, kind="digits", max_len=13,
                               width=int(180 * self.scale), required=True, stretch=False)
        else:
            self._cp_add_field(card, mode, F.OGRNIP, F.OGRNIP, kind="digits", max_len=15,
                               width=int(180 * self.scale), required=True, stretch=False)

        # -- Организация ------------------------------------------------- #
        card.section("Организация" if mode == "ooo" else "Предприниматель")
        if mode == "ooo":
            self._cp_add_field(card, mode, F.FULL_NAME, "Полное название", required=True)
            self._cp_add_field(card, mode, F.SHORT_NAME, "Сокращённое название", required=True)
        else:
            self._cp_add_field(card, mode, F.IP_NAME, F.IP_NAME, required=True)

        # -- Руководитель ------------------------------------------------ #
        if mode == "ooo":
            card.section("Руководитель")
            self._cp_add_field(card, mode, F.DIRECTOR, F.DIRECTOR, required=True)
            widget = ui.option_menu(card.body, ["Устав", "Доверенность"],
                                    width=int(ENTRY_WIDTH * self.scale))
            widget.set("Устав")
            self._cp_entries[mode][F.BASIS] = widget
            card.row(F.BASIS, widget)

        # -- Адреса ------------------------------------------------------ #
        card.section("Адреса")
        self._cp_add_field(card, mode, F.LEGAL_ADDRESS, F.LEGAL_ADDRESS, required=True)
        if mode == "ooo":
            self._cp_add_field(card, mode, F.FACT_ADDRESS, F.FACT_ADDRESS)
        else:
            card.row(
                F.FACT_ADDRESS,
                ctk.CTkLabel(
                    card.body, text="как юридический", font=ui.font(13),
                    text_color=ui.COLORS["text_tertiary"], anchor="w",
                ),
            )

        # -- Контакты ---------------------------------------------------- #
        # Email и телефон необязательны — юридически значимые поля выше (п. 4.5).
        card.section("Контакты")
        self._cp_add_field(card, mode, F.EMAIL, F.EMAIL, kind="email",
                           hint="Необязательно.")
        self._cp_add_field(card, mode, F.PHONE, F.PHONE, kind="phone",
                           hint="Необязательно.")

        # -- Банк -------------------------------------------------------- #
        card.section("Банковские реквизиты")
        self._cp_add_field(card, mode, F.ACCOUNT, F.ACCOUNT, kind="digits", max_len=20,
                           width=int(260 * self.scale), required=True)
        self._cp_add_field(card, mode, F.CORR_ACCOUNT, F.CORR_ACCOUNT, kind="digits", max_len=20,
                           width=int(260 * self.scale), required=True)
        self._cp_add_field(card, mode, F.BANK, F.BANK, required=True)
        self._cp_add_field(card, mode, F.BIK, F.BIK, kind="digits", max_len=9,
                           width=int(140 * self.scale), required=True, stretch=False)

    def _cp_add_field(self, card, mode: str, key: str, label: str, kind="text",
                      width=None, required=False, stretch=True, max_len=None,
                      hint: str | None = None):
        width = width or int(ENTRY_WIDTH * self.scale)
        validate = None
        if kind == "digits":
            validate = self._digits_vcmd_for(max_len) if max_len else (self._digits_vcmd, "%P")
        elif kind == "phone":
            validate = (self._phone_vcmd, "%P")

        widget = ui.entry(card.body, width=width, validatecommand=validate)
        if kind == "digits":
            # Валидатор отбрасывает вставку целиком, если в буфере есть
            # недопустимые символы («р/с 40702810…»). Вставляем только цифры.
            ui.enable_sanitized_paste(widget.entry, digits_only, max_len=max_len)
        elif kind == "phone":
            phone_var = tk.StringVar()
            widget.entry.configure(textvariable=phone_var)
            phone_var.trace_add("write", make_phone_formatter(phone_var, widget.entry))
            ui.enable_sanitized_paste(
                widget.entry,
                lambda s: "".join(ch for ch in str(s) if ch.isdigit() or ch in "+-()/. "),
            )

        self._cp_entries[mode][key] = widget
        card.row(label, widget, stretch=stretch, required=required,
                 error_var=self._cp_error_var_for(mode)(key), hint=hint)
        return widget

    def _cp_error_var_for(self, mode: str):
        def get(key: str) -> tk.StringVar:
            return self._cp_error_vars[mode].setdefault(key, tk.StringVar(value=""))
        return get

    def _on_cp_type_change(self, value: str):
        mode = "ip" if value == "ИП" else "ooo"
        self._cp_frames[mode].tkraise()

    def _cp_collect(self, mode: str) -> dict:
        data = {}
        for key, widget in self._cp_entries[mode].items():
            try:
                data[key] = widget.get().strip()
            except Exception:
                data[key] = ""
        if mode == "ip":
            data[F.FACT_ADDRESS] = data.get(F.LEGAL_ADDRESS, "")
        return data

    def _cp_save(self):
        mode = "ip" if self._cp_type_var.get() == "ИП" else "ooo"
        is_ip = mode == "ip"
        data = self._cp_collect(mode)

        errors = validate_party_fields(data, is_ip)
        if errors:
            for var in self._cp_error_vars[mode].values():
                var.set("")
            for widget in self._cp_entries[mode].values():
                if hasattr(widget, "set_error"):
                    widget.set_error(False)
            for key, message in errors.items():
                self._cp_error_var_for(mode)(key).set(message)
                widget = self._cp_entries[mode].get(key)
                if widget is not None and hasattr(widget, "set_error"):
                    widget.set_error(True)
            first = self._cp_entries[mode].get(next(iter(errors)))
            if first is not None:
                try:
                    first.focus_set()
                except Exception:
                    pass
            ui.alert(
                self.root, "Ошибки в данных",
                "Проверьте подсвеченные поля:\n\n"
                + "\n".join(f"• {error}" for error in errors.values()),
                "error",
            )
            return

        # Дубликаты по ИНН: молча создавать вторую карточку нельзя (п. 4.1).
        inn = data.get(F.INN, "").strip()
        duplicate = self.store.find_by_inn(inn)
        if duplicate is not None:
            choice = self._choice_dialog(
                "Контрагент с таким ИНН уже есть",
                f"ИНН {inn} уже указан у «{storage.display_name(duplicate)}» "
                f"(ID {duplicate.get('id')}).\n\n"
                "Открыть существующего или всё равно создать дубль?",
                [
                    ("open", "Открыть существующего", True),
                    ("create", "Создать всё равно", False),
                    ("cancel", "Отмена", False),
                ],
            )
            if choice != "create":
                if choice == "open":
                    # Пользователь явно выбрал существующего — введённые
                    # значения карточки больше не нужны.
                    self._reset_form_baseline("counterparty")
                    self._set_current_party(duplicate)
                    self._select_page("contract")
                return

        rec = self.store.add_counterparty(mode, data)
        if self.store.last_error:
            # Не оставляем «фантом» в памяти, если запись не легла на диск.
            try:
                self.store.counterparties.remove(rec)
            except ValueError:
                pass
            ui.alert(
                self.root, "Не удалось сохранить контрагента",
                self.store.last_error, "error",
            )
            return
        self._reset_form_baseline("counterparty")
        self._refresh_party_combo_values()
        self._update_status()
        self.refresh_registry()
        self._set_current_party(rec)
        self._notify(f"{storage.display_name(rec)} добавлен в реестр программы.")
        # Сразу переходим к договору с уже выбранным контрагентом.
        self._select_page("contract")

    # ------------------------------------------------------------------ #
    #  DaData: автозаполнение карточки контрагента
    # ------------------------------------------------------------------ #
    def _cp_fill_from_inn(self):
        mode = "ip" if self._cp_type_var.get() == "ИП" else "ooo"
        widget = self._cp_entries[mode].get(F.INN)
        inn = widget.get().strip() if widget is not None else ""
        expected = 12 if mode == "ip" else 10
        if not inn.isdigit() or len(inn) != expected:
            ui.alert(
                self.root, "Проверьте ИНН",
                f"ИНН должен содержать {expected} цифр.", "error",
            )
            return

        api_key = dadata.get_api_key()
        if not api_key:
            ui.alert(
                self.root, "Нет API-ключа",
                "Ключ DaData не задан. Укажите его на странице «Настройки» "
                "или заполните реквизиты вручную.", "error",
            )
            return

        self._cp_busy = True
        self._cp_busy_phase = 0
        self._cp_busy_tick()
        worker = threading.Thread(
            target=self._cp_fetch_worker, args=(inn, api_key, mode), daemon=True,
        )
        worker.start()

    def _cp_busy_tick(self):
        if not self._cp_busy:
            return
        dots = "." * (self._cp_busy_phase % 4)
        self._cp_status_var.set(f"Загружаю реквизиты из ЕГРЮЛ/ЕГРИП{dots}")
        self._cp_busy_phase += 1
        self._cp_busy_after = self.root.after(300, self._cp_busy_tick)

    def _cp_busy_stop(self):
        self._cp_busy = False
        if self._cp_busy_after is not None:
            try:
                self.root.after_cancel(self._cp_busy_after)
            except Exception:
                pass
            self._cp_busy_after = None

    def _cp_fetch_worker(self, inn: str, api_key: str, mode: str):
        try:
            info = dadata.fetch_party(inn, api_key)
        except Exception as exc:
            self.root.after(0, lambda: self._cp_party_failed(str(exc)))
            return
        self.root.after(0, lambda: self._cp_party_loaded(mode, info))

    def _cp_party_failed(self, message: str):
        self._cp_busy_stop()
        self._cp_status_var.set("")
        ui.alert(self.root, "Не удалось заполнить реквизиты", message, "error")

    def _cp_party_loaded(self, mode: str, info: dict):
        self._cp_busy_stop()
        self._cp_status_var.set("")

        filled: list[str] = []

        def put(key: str, value: str):
            if not value:
                return
            widget = self._cp_entries[mode].get(key)
            if widget is None:
                return
            widget.delete(0, tk.END)
            widget.insert(0, value)
            filled.append(key)

        put(F.INN, info["inn"])
        put(F.LEGAL_ADDRESS, info["legal_address"])
        put(F.BANK, info["bank_name"])
        put(F.BIK, info["bik"])
        put(F.CORR_ACCOUNT, info["correspondent_account"])

        if mode == "ooo":
            put(F.FULL_NAME, info["full_name"])
            put(F.SHORT_NAME, info["short_name"])
            put(F.DIRECTOR, info["director"])
            put(F.KPP, info["kpp"])
            put(F.OGRN, info["ogrn"])
            fact = self._cp_entries[mode].get(F.FACT_ADDRESS)
            if fact is not None and not fact.get().strip():
                put(F.FACT_ADDRESS, info["legal_address"])
        else:
            put(F.IP_NAME, info["director"])
            put(F.OGRNIP, info["ogrn"])

        if not filled:
            ui.alert(
                self.root, "Данных нет",
                "Сервис вернул карточку без данных для подстановки.", "error",
            )
            return

        self._cp_status_var.set("Реквизиты заполнены из ЕГРЮЛ/ЕГРИП")

    # ------------------------------------------------------------------ #
    #  Несохранённые данные (п. 4.4)
    # ------------------------------------------------------------------ #
    @staticmethod
    def _form_label(key: str) -> str:
        return {
            "contract": "Договор поставки",
            "additional": "Доп. соглашение",
            "spec": "Спецификация",
            "counterparty": "Карточка контрагента",
        }.get(key, key)

    def _form_signature(self, key: str) -> tuple:
        """Данные формы, потеря которых требует подтверждения.

        Выбор контрагента/договора намеренно не входит в подпись: он
        переносится между страницами (п. 4.2) и не является введёнными данными.
        """
        try:
            if key == "contract":
                return (
                    self.contract_number_var.get(),
                    self.contract_day_var.get(),
                    self.contract_month_var.get(),
                )
            if key == "additional":
                return (
                    self.addl_number_var.get(),
                    self._limit_var.get(),
                    self.term_var.get(),
                )
            if key == "spec":
                items = tuple(
                    (row["name"].get(), row["qty"].get(), row["unit"].get(), row["price"].get())
                    for row in self._spec_rows
                )
                return (
                    self.spec_number_var.get(),
                    self.spec_day_var.get(),
                    self.spec_month_var.get(),
                    self.spec_delivery_term_var.get(),
                    self.spec_payment_terms_var.get(),
                    self.spec_delivery_basis_var.get(),
                    self.spec_prices_until_var.get(),
                    items,
                )
            if key == "counterparty":
                mode = "ip" if self._cp_type_var.get() == "ИП" else "ooo"
                return tuple(sorted(self._cp_collect(mode).items()))
        except Exception:
            return ()
        return ()

    def _reset_form_baseline(self, key: str) -> None:
        self._form_baseline[key] = self._form_signature(key)

    def _reset_all_form_baselines(self) -> None:
        for key in ("contract", "additional", "spec", "counterparty"):
            self._reset_form_baseline(key)

    def _is_form_dirty(self, key: str) -> bool:
        if key not in self._form_baseline:
            return False
        return self._form_signature(key) != self._form_baseline[key]

    def _dirty_pages(self) -> list[str]:
        return [
            key for key in ("contract", "additional", "spec", "counterparty")
            if self._is_form_dirty(key)
        ]

    def _confirm_leave_page(self) -> bool:
        key = self.current_page
        if key not in self.pages or not self._is_form_dirty(key):
            return True
        return ui.confirm(
            self.root, "Несохранённые изменения",
            f"На странице «{self._form_label(key)}» есть несохранённые данные.\n\n"
            "Перейти и потерять изменения?",
            ok_text="Перейти", cancel_text="Остаться",
        )

    def _save_current_page(self) -> None:
        commands = {
            "contract": self.save_contract,
            "additional": self.save_additional,
            "spec": self.save_spec,
            "counterparty": self._cp_save,
        }
        command = commands.get(self.current_page)
        if command is not None:
            command()

    def _on_save_shortcut(self, _event=None):
        self._save_current_page()
        return "break"

    @staticmethod
    def _is_descendant(widget, ancestor) -> bool:
        if ancestor is None:
            return False
        try:
            while widget is not None:
                if widget is ancestor:
                    return True
                widget = widget.master
        except Exception:
            pass
        return False

    def _on_enter_save(self, _event=None):
        """Enter в форме сохраняет документ, но не мешает таблице позиций."""
        if self.current_page not in ("contract", "additional", "spec", "counterparty"):
            return None
        widget = self.root.focus_get()
        if widget is None:
            return None
        if self.current_page == "spec" and self._is_descendant(widget, getattr(self, "spec_table", None)):
            return None
        try:
            if isinstance(widget, tk.Text):
                return None
        except Exception:
            pass
        self._save_current_page()
        return "break"

    # ------------------------------------------------------------------ #
    #  Drag&drop файлов в окно (п. 4.5)
    # ------------------------------------------------------------------ #
    def _enable_drop_files(self) -> None:
        """Принимает файлы, перетащенные в окно, через WM_DROPFILES.

        Дополнительных зависимостей не требуется. При любой ошибке
        возможность просто отключается — остаётся кнопка «Добавить документы…».
        """
        if sys.platform != "win32":
            return
        try:
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            user32.GetParent.restype = ctypes.c_void_p
            user32.GetParent.argtypes = [ctypes.c_void_p]
            shell32.DragAcceptFiles.argtypes = [ctypes.c_void_p, ctypes.c_bool]
            hwnd = user32.GetParent(ctypes.c_void_p(self.root.winfo_id()))
            if not hwnd:
                hwnd = self.root.winfo_id()
            shell32.DragAcceptFiles(ctypes.c_void_p(hwnd), True)

            GWLP_WNDPROC = -4
            user32.GetWindowLongPtrW.restype = ctypes.c_void_p
            user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
            user32.SetWindowLongPtrW.restype = ctypes.c_void_p
            user32.SetWindowLongPtrW.argtypes = [
                ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p,
            ]
            user32.CallWindowProcW.restype = ctypes.c_longlong
            user32.CallWindowProcW.argtypes = [
                ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint,
                ctypes.c_ulonglong, ctypes.c_longlong,
            ]
            old_proc = user32.GetWindowLongPtrW(ctypes.c_void_p(hwnd), GWLP_WNDPROC)
            if not old_proc:
                return

            wndproc_type = ctypes.WINFUNCTYPE(
                ctypes.c_longlong, ctypes.c_void_p, ctypes.c_uint,
                ctypes.c_ulonglong, ctypes.c_longlong,
            )

            def handler(h, msg, wparam, lparam):
                if msg == 0x0233:  # WM_DROPFILES
                    try:
                        count = shell32.DragQueryFileW(
                            ctypes.c_void_p(wparam), 0xFFFFFFFF, None, 0
                        )
                        files: list[str] = []
                        for index in range(count):
                            length = shell32.DragQueryFileW(
                                ctypes.c_void_p(wparam), index, None, 0
                            )
                            buffer = ctypes.create_unicode_buffer(length + 1)
                            shell32.DragQueryFileW(
                                ctypes.c_void_p(wparam), index, buffer, length + 1
                            )
                            if buffer.value:
                                files.append(buffer.value)
                        shell32.DragFinish(ctypes.c_void_p(wparam))
                        if files:
                            self.root.after(
                                0, lambda f=list(files): self._start_documents_import(f)
                            )
                        return 0
                    except Exception:
                        pass
                return user32.CallWindowProcW(old_proc, h, msg, wparam, lparam)

            self._drop_proc = wndproc_type(handler)
            self._drop_old_proc = old_proc
            self._drop_hwnd = hwnd
            user32.SetWindowLongPtrW(
                ctypes.c_void_p(hwnd), GWLP_WNDPROC,
                ctypes.cast(self._drop_proc, ctypes.c_void_p),
            )
        except Exception as exc:
            self._log.warning("Drag&drop недоступен: %s", type(exc).__name__)

    # ------------------------------------------------------------------ #
    #  Навигация и статус
    # ------------------------------------------------------------------ #
    def _select_page(self, key: str):
        if key != self.current_page and not self._confirm_leave_page():
            return
        self.current_page = key
        self._update_footer()
        for page_key, frame in self.pages.items():
            if page_key == key:
                frame.tkraise()
        for nav_key, item in self.nav_items.items():
            item.set_selected(nav_key == key)
        if key == "registry":
            self.refresh_registry()
        self.root.after(60, self._refresh_scrollbars)

    def _refresh_scrollbars(self):
        """Прячет полосу прокрутки, если содержимое помещается целиком (как в macOS)."""
        for scroller in self.scrollers:
            try:
                canvas = scroller._parent_canvas
                first, last = canvas.yview()
                if first <= 0.001 and last >= 0.999:
                    scroller._scrollbar.grid_remove()
                else:
                    scroller._scrollbar.grid()
            except Exception:
                continue

    def _update_status(self, *_args):
        """Строка в футере: сколько контрагентов и документов в реестре."""
        if not hasattr(self, "status_label"):
            return
        parties = len(self.store.active_counterparties())
        active_docs = self.store.active_documents()
        documents = sum(
            1 for rec in active_docs
            if rec.get("kind") != storage.KIND_UNKNOWN
        )
        contracts = sum(1 for rec in active_docs if rec.get("kind") == storage.KIND_CONTRACT)
        trash = self.store.trash_count()
        status = f"Контрагентов: {parties} · Документов: {documents} (договоров: {contracts})"
        if trash:
            status += f" · В корзине: {trash}"
        self.status_label.configure(
            text=status,
            text_color=ui.resolve(ui.COLORS["text_secondary"]),
        )

    # ------------------------------------------------------------------ #
    #  Доп. соглашение: суммы прописью и формат лимита
    # ------------------------------------------------------------------ #
    def _update_words(self, *_args):
        value = self.term_var.get().strip()
        if not value:
            self.words_label.configure(text="Срок прописью: —")
        elif value.isdigit():
            self.words_label.configure(text=f"Срок прописью: {number_to_words_ru(int(value))}")
        else:
            self.words_label.configure(text="Срок прописью: введите целое число дней")

    def _format_limit(self, *_args):
        """Форматирует лимит с разделителями разрядов и обновляет сумму прописью."""
        raw = self._limit_var.get()
        formatted = format_number_grouped(raw)
        if raw != formatted:
            entry = None
            caret = None
            try:
                entry = self.limit_entry.entry
                caret = entry.index(tk.INSERT)
            except Exception:
                pass
            self._limit_var.set(formatted)
            restore_caret_after_reformat(entry, caret, len(raw), len(formatted))
        self._update_limit_words()

    def _update_limit_words(self):
        digits = digits_only(self._limit_var.get())
        if not digits:
            self.limit_words_label.configure(text="Сумма прописью: —")
            return
        try:
            self.limit_words_label.configure(text=f"Сумма прописью: {amount_to_words_ru(int(digits))}")
        except Exception:
            self.limit_words_label.configure(text="Сумма прописью: —")

    def run(self):
        self.root.mainloop()


def main():
    """Точка входа: один экземпляр, логирование, корректное освобождение блокировки."""
    lock = single_instance.InstanceLock()
    if not lock.acquire():
        # Второй экземпляр писал бы в тот же database.json и мог затереть данные.
        single_instance.activate_existing_window()
        single_instance.show_already_running()
        return

    applog.setup_logging(APP_VERSION)
    log = applog.get_logger("main")
    try:
        app = ContractApp()
        app.run()
    except Exception:
        log.exception("Необработанная ошибка при работе приложения")
        raise
    finally:
        lock.release()


if __name__ == "__main__":
    main()
