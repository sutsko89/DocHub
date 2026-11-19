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

import copy
import counterparty_controller
import ctypes
import json
import os
import platform
import queue
import re
import subprocess
import sys
import threading
from datetime import datetime

from decimal import Decimal, ROUND_HALF_UP

import customtkinter as ctk
import tkinter as tk
from customtkinter.windows.widgets.scaling.scaling_tracker import ScalingTracker
from docx import Document
from tkinter import filedialog

import applog
import dadata_client as dadata
import dadata_controller
import document_importer as importer
import document_save_controller
import import_controller
import lifecycle_controller
import reconcile
import registry_controller
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


# Сервисный слой (п. 9.2): валидация, кэш шаблонов, запись документов и
# транзакция сохранения переехали в services.py; имена реэкспортируются,
# чтобы main.py и smoke_test.py работали как раньше.
import services
from services import (
    MAX_PATH_LENGTH, MAX_FILENAME_LENGTH, resource_path, _TEMPLATE_CACHE,
    load_template_cached, validate_ogrn, validate_ogrnip, validate_inn,
    validate_kpp, validate_email, validate_phone, validate_account,
    validate_correspondent_account, validate_bik, _PHONE_ALLOWED_RE,
    REQUIRED_COMMON, REQUIRED_OOO, REQUIRED_IP, REQUIRED_PARTY_COMMON,
    validate_contract_fields, validate_party_fields, validate_all_fields,
    save_docx_atomic, _RESERVED_FILENAMES, _FORBIDDEN_FILENAME_CHARS,
    safe_filename, copy_with_zone_identifier,
)


# Составные виджеты вынесены в widgets.py (п. 9.4).
from widgets import SpecTable


APP_VERSION = "2.6.0"
# Оставляем запас под полный путь Windows: лимит 260 символов, минус имя файла
# и расширение, минус служебные суффиксы « (2).docx».

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


# ---------- Кэш шаблонов печатных форм (п. 5.5 этапа 5) ----------
# Шаблон .docx разбирается один раз; на каждый рендер отдаётся deepcopy.
# Раньше предпросмотр при каждом вводе заново читал и разбирал ZIP шаблона.
# Кэш инвалидируется по mtime — правка шаблона подхватывается без перезапуска.


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


# ---------- Настройки печатной формы спецификации (п. 9.7) ----------
# Магические значения печатной формы спецификации вынесены в settings.json:
# поставщик, подписант, город, год, ставка НДС и номер приложения меняются
# без правки кода. Отсутствующие или битые значения заменяются дефолтами.
# Год здесь НЕ фиксируется: раньше datetime.now().year вычислялся при импорте
# модуля, и приложение, пережившее Новый год в запущенном виде, ставило в
# спецификации прошлый год (аудит 2.0.1). Актуальный год берётся на момент
# вызова get_spec_settings().
DEFAULT_SPEC_SETTINGS = {
    "spec_supplier": "ООО «Эдельвейс»",
    "spec_supplier_signatory": "А.Н. Верещагин",
    "spec_city": "Хабаровск",
    "spec_vat_rate": SPEC_VAT_RATE,
    "spec_appendix_number": SPEC_APPENDIX_NUMBER,
}


def get_spec_settings() -> dict:
    """Печатная форма спецификации: дефолты, перекрытые settings.json."""
    saved = load_settings()
    out = dict(DEFAULT_SPEC_SETTINGS)
    out["spec_year"] = datetime.now().year  # на момент вызова, не импорта
    for key in out:
        value = saved.get(key)
        if value is None:
            continue
        if key == "spec_vat_rate":
            try:
                out[key] = max(0, min(int(value), 100))
            except (TypeError, ValueError):
                pass
        elif key == "spec_year":
            try:
                out[key] = int(value)
            except (TypeError, ValueError):
                pass
        else:
            out[key] = str(value)
    return out


# ---------- Функции валидации ----------


# Регулярка одна на оба места: проверку сохранённых данных и фильтр ввода на лету.


# ---------- Ключи полей и плейсхолдеров (единый источник имён) ----------


# Email и телефон — контактные, а не юридически обязательные поля (п. 4.5).

# Константы печатной формы спецификации (spec_template.docx).

# Импорт принимает только документы, которые программа умеет разбирать.
# Всё остальное (в т. ч. .exe/.bat/.lnk) не копируется и не открывается.
# Канонический набор расширений — import_controller.ALLOWED_EXTENSIONS.
IMPORT_ALLOWED_EXTENSIONS = import_controller.ALLOWED_EXTENSIONS

# Лимиты пакетного импорта (п. 3.4): файл, количество, суммарный объём,
# свободное место на диске. Проверяются до запуска фонового разбора.
IMPORT_MAX_FILE_SIZE = importer.MAX_FILE_SIZE
IMPORT_MAX_FILES = importer.MAX_FILES_PER_BATCH
IMPORT_MAX_TOTAL_SIZE = importer.MAX_TOTAL_SIZE
IMPORT_MIN_FREE_SPACE = importer.MIN_FREE_SPACE


# Поля карточки контрагента (без реквизитов договора).
# Email и телефон необязательны: без них карточку можно сохранить (п. 4.5).


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


# --------------------------------------------------------------------------- #
#  Таблица позиций спецификации (excel-подобная, вставка из буфера обмена)
# --------------------------------------------------------------------------- #


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
        # Оркестрация DaData (заполнение по ИНН) — в отдельном контроллере
        # (шаг декомпозиции ContractApp, аудит 2.0.1 P2): busy-анимация,
        # воркер и опрос очереди живут в dadata_controller.DadataController.
        self._dadata_flow = dadata_controller.DadataController(self)
        # Страница «Реестр»: рендер, поиск, выбор, сверка — контроллер (шаг 4
        # декомпозиции ContractApp, аудит 2.0.1 P2). Создаётся ДО построения
        # страниц: _build_registry_page уже идёт через контроллер.
        self._registry_flow = registry_controller.RegistryController(self)
        # Спецификация и сохранение документов — тоже контроллер (шаг 5
        # декомпозиции ContractApp, аудит 2.0.1 P2). Создаётся ДО построения
        # страниц: _build_spec_page уже идёт через контроллер.
        self._doc_save_flow = document_save_controller.DocumentSaveController(self)

        # Страница «Добавить контрагента» и редактирование карточки — тоже
        # контроллер (шаг 6 декомпозиции ContractApp, аудит 2.0.1 P2).
        # Создаётся ДО построения страниц: _build_counterparty_page уже
        # идёт через контроллер.
        self._counterparty_flow = counterparty_controller.CounterpartyController(self)

        # Единый «текущий контрагент» и защита несохранённых данных (п. 4.2, 4.4).
        self._current_party_id = None
        self._current_party_vars: dict[str, tk.StringVar] = {}
        self._syncing_party = False
        self._form_baseline: dict[str, tuple | None] = {}
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
        # Оркестрация пакетного импорта — в отдельном контроллере (шаг
        # декомпозиции ContractApp, аудит 2.0.1 P2).
        self._import_flow = import_controller.ImportController(self)
        # Жизненный цикл записей реестра — экспорт и корзина — тоже контроллер
        # (третий шаг декомпозиции ContractApp, аудит 2.0.1 P2).
        self._lifecycle_flow = lifecycle_controller.LifecycleController(self)
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
        """Модальный диалог с N кнопками — примитив переехал в ui_macos.

        ``options``: список ``(ключ, подпись, главная_кнопка)`` слева направо.
        Возвращает ключ выбранного варианта или ``None`` при закрытии окна.
        """
        return ui.choice_dialog(self.root, title, message, options, self.scale)

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
    #  Спецификация и сохранение документов — DocumentSaveController
    # ------------------------------------------------------------------ #
    #  Код страницы «Спецификация» — таблица позиций, договор-основание,
    #  предпросмотр — переехал в document_save_controller.DocumentSave-
    #  Controller (пятый шаг декомпозиции ContractApp, аудит 2.0.1 P2).
    #  Состояние страницы остаётся на приложении с прежними именами — его
    #  читают защита несохранённых данных (п. 4.2, 4.4), страницы договора
    #  и доп. соглашения и smoke-тест; здесь остаются фасады с прежними
    #  именами для внешних вызывающих.
    def _build_spec_page(self):
        """Строит страницу «Спецификация» (контроллер сохранения документов)."""
        self._doc_save_flow.build_spec_page()

    @property
    def _spec_rows(self) -> list[dict]:
        """Строки позиций страницы «Спецификация» (контроллер сохранения)."""
        return self._doc_save_flow.spec_rows

    def _add_spec_row(self, name="", qty="", price="", unit=""):
        """Добавляет строку в таблицу позиций (контроллер сохранения)."""
        return self._doc_save_flow.add_spec_row(name=name, qty=qty, price=price, unit=unit)

    def _remove_spec_row(self):
        """Удаляет последнюю строку таблицы позиций (контроллер сохранения)."""
        self._doc_save_flow.remove_spec_row()

    def _spec_items(self) -> list[dict]:
        """Позиции с заполненным наименованием (контроллер сохранения)."""
        return self._doc_save_flow.spec_items()

    def _spec_total(self) -> Decimal:
        """Сумма позиций спецификации (контроллер сохранения)."""
        return self._doc_save_flow.spec_total()

    def _on_spec_item_changed(self, *_args):
        """Пересобирает предпросмотр после правки ячейки таблицы (контроллер)."""
        self._doc_save_flow.on_spec_item_changed(*_args)

    def _on_spec_party_changed(self, *_args):
        """Подставляет договор-основание при выборе контрагента (контроллер)."""
        self._doc_save_flow.on_spec_party_changed(*_args)

    # ------------------------------------------------------------------ #
    #  Страница «Реестр» — рендер и сверки в RegistryController
    # ------------------------------------------------------------------ #
    #  Построение страницы, поиск, выбор контрагента, рендер секций
    #  (документы / «unknown» / корзина) и сверка реестра с папкой
    #  документов переехали в registry_controller.RegistryController
    #  (четвёртый шаг декомпозиции ContractApp, аудит 2.0.1 P2). Состояние
    #  страницы остаётся на приложении с прежними именами — его читают
    #  другие страницы, сценарии сохранения и smoke-тест; здесь остаются
    #  фасады с прежними именами для внешних вызывающих.
    def _build_registry_page(self):
        """Строит страницу «Реестр» (контроллер реестра)."""
        self._registry_flow.build_page()

    def refresh_registry(self):
        """Перерисовывает реестр целиком (контроллер реестра)."""
        self._registry_flow.refresh()

    def _registry_apply_filter(self):
        """Применяет поиск и перерисовывает реестр (контроллер реестра)."""
        self._registry_flow.apply_filter()

    def _select_registry_party(self, rec: dict):
        """Клик по контрагенту: показать/скрыть его документы (контроллер)."""
        self._registry_flow.select_party(rec)

    def _scan_documents_consistency(self) -> None:
        """Фоновая сверка реестра с папкой документов (контроллер реестра)."""
        self._registry_flow.scan_consistency()

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

    # ------------------------------------------------------------------ #
    #  Сохранение документов — запись и сценарии в DocumentSaveController
    # ------------------------------------------------------------------ #
    #  Папки контрагентов, транзакция «файл + запись реестра» (п. 9.2),
    #  печатные формы из шаблонов и записей, предпросмотр спецификации,
    #  подтверждение повтора номера (п. 4.1) и save_contract /
    #  save_additional / save_spec переехали в
    #  document_save_controller.DocumentSaveController (пятый шаг
    #  декомпозиции ContractApp, аудит 2.0.1 P2); здесь остаются фасады
    #  с прежними именами — страницы форм, Ctrl+S и smoke-тест продолжают
    #  работать как раньше. Открытие файла после сохранения идёт через
    #  фасад _open_file: smoke-тест подменяет модульное main.open_file.
    def _persist_document(self, draft: dict) -> dict:
        """Пишет .docx и регистрирует запись (контроллер сохранения)."""
        return self._doc_save_flow.persist_document(draft)

    def _party_documents_dir(self, party_rec: dict | None,
                             interactive: bool = True) -> str:
        """Индивидуальная папка контрагента (контроллер сохранения)."""
        return self._doc_save_flow.party_documents_dir(party_rec, interactive=interactive)

    def _render_contract_docx(self, data: dict, is_ip: bool) -> "Document":
        """Договор из шаблона (контроллер сохранения)."""
        return self._doc_save_flow.render_contract_docx(data, is_ip)

    def _render_additional_docx(self, data: dict, is_ip: bool, number: str,
                                limit_display: str, term: str) -> "Document":
        """Доп. соглашение из шаблона (контроллер сохранения)."""
        return self._doc_save_flow.render_additional_docx(
            data, is_ip, number, limit_display, term)

    def _render_spec_docx(self, rec: dict) -> "Document":
        """Спецификация из шаблона (контроллер сохранения)."""
        return self._doc_save_flow.render_spec_docx(rec)

    def _rec_is_ip(self, rec: dict) -> bool:
        """Признак ИП у сохранённой записи (контроллер сохранения)."""
        return self._doc_save_flow.rec_is_ip(rec)

    def _build_record_docx(self, rec: dict) -> "Document":
        """Собирает .docx записи реестра по её типу (контроллер сохранения)."""
        return self._doc_save_flow.build_record_docx(rec)

    def _write_docx(self, doc, rec: dict) -> str:
        """Пишет .docx в папку контрагента атомарно (контроллер сохранения)."""
        return self._doc_save_flow.write_docx(doc, rec)

    def _save_and_open_document(self, rec: dict) -> str:
        """Готовит .docx и открывает его в Word (контроллер сохранения)."""
        return self._doc_save_flow.save_and_open_document(rec)

    def _persist_document_result_alert(self, result: dict):
        """Сообщение при сбое транзакции сохранения (контроллер сохранения)."""
        self._doc_save_flow.persist_document_result_alert(result)

    def _regenerate_spec(self, *_args):
        """Откладывает пересборку предпросмотра спецификации (контроллер)."""
        self._doc_save_flow.regenerate_spec(*_args)

    def _flush_spec_preview(self):
        """Синхронно дособирает отложенный предпросмотр (контроллер)."""
        self._doc_save_flow.flush_spec_preview()

    def _render_spec_preview(self):
        """Собирает текст предпросмотра спецификации (контроллер)."""
        self._doc_save_flow.render_spec_preview()

    def _confirm_document_duplicate(
        self, party_id, kind: str, number: str, date_text: str
    ) -> bool:
        """Предупреждает о повторе номера документа (контроллер сохранения)."""
        return self._doc_save_flow.confirm_document_duplicate(party_id, kind, number, date_text)

    def save_contract(self):
        """Сохраняет договор поставки (контроллер сохранения)."""
        self._doc_save_flow.save_contract()

    def save_additional(self):
        """Сохраняет доп. соглашение (контроллер сохранения)."""
        self._doc_save_flow.save_additional()

    def save_spec(self):
        """Сохраняет спецификацию (контроллер сохранения)."""
        self._doc_save_flow.save_spec()

    # ------------------------------------------------------------------ #
    #  Реестр: экспорт и удаление — оркестрация в LifecycleController
    # ------------------------------------------------------------------ #
    #  Сценарии экспорта/открытия документа и корзины переехали в
    #  lifecycle_controller.LifecycleController (третий шаг декомпозиции
    #  ContractApp, аудит 2.0.1 P2); здесь остаются фасады с прежними
    #  именами — кнопки реестра и smoke-тест продолжают работать как раньше.
    def export_document(self, rec: dict):
        """Открывает/экспортирует документ из реестра (контроллер жизненного цикла)."""
        self._lifecycle_flow.export_document(rec)

    def _delete_document(self, rec: dict):
        """Перемещает документ в корзину после подтверждения (п. 2.5, 2.6)."""
        self._lifecycle_flow.delete_document(rec)

    def _delete_counterparty(self, rec: dict):
        """Перемещает контрагента и его документы в корзину (п. 2.5, 2.6)."""
        self._lifecycle_flow.delete_counterparty(rec)

    def _restore_document(self, rec: dict):
        """Возвращает документ из корзины (п. 2.6)."""
        self._lifecycle_flow.restore_document(rec)

    def _restore_counterparty(self, rec: dict):
        """Возвращает контрагента и его документы из корзины (п. 2.6)."""
        self._lifecycle_flow.restore_counterparty(rec)

    def _purge_document(self, rec: dict):
        """Удаляет запись документа из корзины окончательно (контроллер)."""
        self._lifecycle_flow.purge_document(rec)

    def _purge_counterparty(self, rec: dict):
        """Удаляет контрагента и его документы окончательно (контроллер)."""
        self._lifecycle_flow.purge_counterparty(rec)

    def _empty_trash(self):
        """Полностью очищает корзину (записи; файлы на диске остаются)."""
        self._lifecycle_flow.empty_trash()

    # -- Оболочка: открытие файлов для контроллера жизненного цикла -------- #
    def _open_file(self, file_path) -> bool:
        """Открывает файл системным приложением — модульная main.open_file.

        Отдельный фасад нужен потому, что smoke-тест подменяет модульное имя
        ``main.open_file``: тело метода резолвит имя из globals модуля main
        на каждый вызов и поэтому видит подмену.
        """
        return open_file(file_path)

    def _show_in_folder(self, file_path: str) -> None:
        """Показывает файл в проводнике — модульная main.show_in_folder.

        Маршрутизация через приложение — как у _open_file: smoke-тест
        подменяет ``main.show_in_folder``, вызов из тела метода приложения
        видит подмену.
        """
        show_in_folder(file_path)

    # ------------------------------------------------------------------ #
    #  Редактирование контрагента — оркестрация в CounterpartyController
    # ------------------------------------------------------------------ #
    #  Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2): окно
    #  редактирования карточки живёт в counterparty_controller.
    #  CounterpartyController; здесь — фасад с прежним именем. Изменяется
    #  только карточка: документы хранят собственную копию реквизитов
    #  на момент создания.

    def _edit_counterparty(self, rec: dict):
        """Окно редактирования контрагента (контроллер карточки)."""
        self._counterparty_flow.edit_counterparty(rec)

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
                     "Без ключа доступен только ручной ввод реквизитов. При "
                     "автозаполнении сервису отправляется только ИНН — публичный "
                     "идентификатор организации/ИП; реквизиты карточки и документы "
                     "никуда не отправляются.",
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
        # Аудит 2.0.1: сохранённый ключ НЕ подставляется в поле — реальное
        # значение не показывается и не копируется по Ctrl+C. Поле остаётся
        # пустым, пока пользователь не решит сменить ключ; состояние — строкой ниже.
        self._dadata_key_status_var = tk.StringVar(value=self._dadata_key_status_text())
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=self._dadata_key_status_var,
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((LEFT_PANEL_WIDTH - 48) * self.scale),
            ),
            pady=(int(2 * self.scale), int(4 * self.scale)),
        )
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

    def _dadata_key_status_text(self) -> str:
        """Статус сохранённого ключа — без подстановки самого значения."""
        if dadata.load_api_key():
            if dadata.dpapi_available():
                return "Ключ сохранён и зашифрован (DPAPI, профиль пользователя)."
            return "Ключ сохранён в профиле пользователя."
        return "Ключ не задан."

    def _save_dadata_key(self):
        """Сохраняет введённый пользователем ключ DaData."""
        key = self._dadata_key_entry.get().strip()
        if not key:
            ui.alert(self.root, "Ключ не указан",
                     "Введите API-ключ DaData или оставьте поле пустым.", "error")
            return
        if dadata.save_api_key(key):
            try:
                self._dadata_key_entry.delete(0, "end")
            except Exception:
                pass
            self._dadata_key_status_var.set(self._dadata_key_status_text())
            ui.alert(
                self.root, "Ключ сохранён",
                "Автозаполнение по ИНН будет использовать этот ключ."
                + (" Значение зашифровано (DPAPI)." if dadata.dpapi_available() else ""),
                "success",
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
        self._dadata_key_status_var.set(self._dadata_key_status_text())
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

    # ------------------------------------------------------------------ #
    #  Импорт документов: оркестрация в ImportController
    # ------------------------------------------------------------------ #
    #  Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 — «по
    #  одному контроллеру за итерацию»): логика пакетного импорта живёт в
    #  import_controller.ImportController, здесь остаются фасады с прежними
    #  именами — smoke-тест, drag-and-drop-хук и внутренние вызовы продолжают
    #  работать как раньше. Состояние потока импорта публикуется свойствами:
    #  smoke-тест читает и подменяет его через прежние имена атрибутов.

    @property
    def _import_running(self) -> bool:
        """Выполняется ли сейчас пакетный импорт (состояние контроллера)."""
        return self._import_flow.running

    @_import_running.setter
    def _import_running(self, value) -> None:
        self._import_flow.running = bool(value)

    @property
    def _import_queue(self):
        """Очередь результатов текущего импорта (или None между запусками)."""
        return self._import_flow.queue

    @_import_queue.setter
    def _import_queue(self, value) -> None:
        self._import_flow.queue = value

    @property
    def _import_win(self):
        """Окно прогресса импорта (или None между запусками)."""
        return self._import_flow.win

    @_import_win.setter
    def _import_win(self, value) -> None:
        self._import_flow.win = value

    @property
    def _import_cancelled(self) -> bool:
        """Запрошена ли отмена остатка пакета (п. 5.3)."""
        return self._import_flow.cancelled

    @_import_cancelled.setter
    def _import_cancelled(self, value) -> None:
        self._import_flow.cancelled = bool(value)

    def _start_documents_import(self, files: list[str]) -> None:
        """Проверяет лимиты пакета и запускает фоновый анализ файлов.

        Guard повторного входа и проверки пакета — в ImportController.start;
        drag-and-drop вызывает метод напрямую, поэтому защита стоит в начале
        запуска (аудит 2.0.1, P1). Фабрика потока передаётся из main на
        каждый вызов: smoke-тест подменяет main.threading.Thread синхронным
        исполнителем, и подмена должна действовать в момент вызова.
        """
        self._import_flow.start(files, thread_factory=threading.Thread)

    def _cancel_import(self) -> None:
        """Отменяет импорт: остаток пакета не анализируется и не регистрируется (п. 5.3)."""
        self._import_flow.cancel()

    def _import_worker(self, files: list[str], result_queue: "queue.Queue") -> None:
        """Фоновый поток: анализ каждого файла и копирование в папку контрагента."""
        self._import_flow.worker(files, result_queue)

    def _copy_imported_file(self, result: dict) -> str:
        """Копирует в папку контрагента только разобранные .docx/.xlsx."""
        return self._import_flow.copy_imported_file(result)

    def _import_tick(self) -> None:
        """Опрос очереди фонового потока из главного потока Tk."""
        self._import_flow.tick()

    def _finish_documents_import(self, results: list[dict]) -> None:
        """Регистрирует результаты анализа в реестре и показывает отчёт."""
        self._import_flow.finish(results)

    @staticmethod
    def _new_import_entry(result: dict) -> dict:
        """Заготовка строки отчёта импорта для одного файла."""
        return import_controller.ImportController.new_entry(result)

    def _register_import_result(self, result: dict, entry: dict, stats: dict) -> bool:
        """Регистрирует один результат разбора; True — если реестр менялся."""
        return self._import_flow.register_result(result, entry, stats)

    def _ask_import_party(self, result: dict) -> "tuple[dict | None, str]":
        """Диалог: к какому контрагенту отнести документ без ИНН (п. 3.3)."""
        return self._import_flow.ask_party(result)

    def _import_document_is_duplicate(self, party_id, result: dict) -> bool:
        """Такой документ уже есть у контрагента (тип + номер + дата)."""
        return self._import_flow.document_is_duplicate(party_id, result)

    def _move_import_to_unknown(self, path: str) -> str:
        """Переносит отклонённый файл в папку «unknown» (п. 3.5)."""
        return self._import_flow.move_to_unknown(path)

    def _show_import_report(self, entries: list[dict], summary: str) -> None:
        """Окно отчёта: по каждому файлу тип, результат и причина (п. 3.5)."""
        self._import_flow.show_report(entries, summary)

    # ------------------------------------------------------------------ #
    #  Окно «Добавить контрагента»
    # ------------------------------------------------------------------ #
    # ------------------------------------------------------------------ #
    #  Страница «Добавить контрагента» — оркестрация в CounterpartyController
    # ------------------------------------------------------------------ #
    #  Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2): сборка
    #  страницы и полей, сбор данных, валидация с подсветкой ошибок, выбор
    #  при дубликате ИНН и сохранение карточки живут в
    #  counterparty_controller.CounterpartyController; здесь — фасады с
    #  прежними именами. Состояние страницы (_cp_entries, _cp_error_vars,
    #  _cp_frames, _cp_type_var, _cp_status_var) остаётся на приложении:
    #  его читают DadataController, защита несохранённых данных и smoke-тест.

    def _build_counterparty_page(self):
        """Страница «Добавить контрагента» (контроллер карточки)."""
        self._counterparty_flow.build_counterparty_page()

    def _build_cp_fields(self, card, mode: str):
        """Поля карточки одного режима (юридическое лицо / ИП)."""
        self._counterparty_flow.build_cp_fields(card, mode)

    def _cp_add_field(self, card, mode: str, key: str, label: str, kind="text",
                      width=None, required=False, stretch=True, max_len=None,
                      hint: str | None = None):
        """Одно поле карточки с валидатором и санитизацией вставки."""
        return self._counterparty_flow.add_field(
            card, mode, key, label, kind=kind, width=width, required=required,
            stretch=stretch, max_len=max_len, hint=hint,
        )

    def _cp_error_var_for(self, mode: str):
        """Переменная ошибки поля карточки (создаётся лениво)."""
        return self._counterparty_flow.error_var_for(mode)

    def _on_cp_type_change(self, value: str):
        """Переключение «Юр. лицо / ИП»: поднимаем соответствующий кадр."""
        self._counterparty_flow.on_type_change(value)

    def _cp_collect(self, mode: str) -> dict | None:
        """Собирает данные карточки; None — сбой чтения (аудит 2.0.1)."""
        return self._counterparty_flow.collect(mode)

    def _cp_save(self):
        """Сохраняет карточку контрагента в реестр (контроллер карточки)."""
        self._counterparty_flow.save_card()

    # ------------------------------------------------------------------ #
    #  DaData: автозаполнение — оркестрация в DadataController
    # ------------------------------------------------------------------ #
    #  Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2): guard
    #  повторных кликов, busy-анимация, воркер и опрос очереди живут в
    #  dadata_controller.DadataController; здесь — фасады с прежними именами.
    #  Результат приходит через очередь и опрашивается только из главного
    #  потока Tk — root.after() из рабочего потока запрещён (аудит 2.0.1).

    @property
    def _cp_busy(self) -> bool:
        """Выполняется ли сейчас запрос реквизитов по ИНН."""
        return self._dadata_flow.busy

    @_cp_busy.setter
    def _cp_busy(self, value) -> None:
        self._dadata_flow.busy = bool(value)

    def _cp_fill_from_inn(self) -> None:
        """Заполняет карточку контрагента из ЕГРЮЛ/ЕГРИП по ИНН.

        Guard повторных кликов — в DadataController.fill_from_inn (аудит
        2.0.1). Фабрика потока передаётся из main на каждый вызов:
        smoke-тест подменяет main.threading.Thread синхронным исполнителем.
        """
        self._dadata_flow.fill_from_inn(thread_factory=threading.Thread)

    def _cp_busy_tick(self) -> None:
        """Кадр анимации «Загружаю реквизиты…»."""
        self._dadata_flow.busy_tick()

    def _cp_busy_stop(self) -> None:
        """Останавливает анимацию и цикл опроса."""
        self._dadata_flow.busy_stop()

    def _cp_fetch_worker(self, inn: str, api_key: str, mode: str,
                         result_queue: "queue.Queue") -> None:
        """Рабочий поток DaData: только HTTP и очередь — никаких вызовов Tk."""
        self._dadata_flow.fetch_worker(inn, api_key, mode, result_queue)

    def _cp_result_tick(self) -> None:
        """Опрашивает очередь результата DaData из главного потока Tk."""
        self._dadata_flow.result_tick()

    def _cp_party_failed(self, message: str) -> None:
        """Запрос не удался: стоп-анимация, ошибка пользователю."""
        self._dadata_flow.party_failed(message)

    def _cp_party_loaded(self, mode: str, info: dict) -> None:
        """Карточка получена: подстановка реквизитов в поля формы."""
        self._dadata_flow.party_loaded(mode, info)

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

    def _form_signature(self, key: str) -> tuple | None:
        """Данные формы, потеря которых требует подтверждения.

        Выбор контрагента/договора намеренно не входит в подпись: он
        переносится между страницами (п. 4.2) и не является введёнными данными.

        Аудит 2.0.1: при сбое чтения возвращается None, а не пустой кортеж —
        None не совпадает ни с каким baseline, поэтому защита несохранённых
        данных считает страницу «грязной» и просит подтверждение, вместо
        того чтобы молча решить, что терять нечего.
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
                data = self._cp_collect(mode)
                if data is None:
                    return None
                return tuple(sorted(data.items()))
        except Exception as exc:
            # Сбой чтения формы — повод предположить несохранённые данные,
            # а не «чистоту» страницы (аудит 2.0.1).
            self._log.warning(
                "Не удалось снять подпись формы «%s»: %s", key, type(exc).__name__,
            )
            return None
        return None

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
