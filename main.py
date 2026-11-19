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
import document_page_controller
import document_context_controller
import document_save_controller
import import_controller
import lifecycle_controller
import reconcile
import registry_controller
import settings_controller
import single_instance
import storage
import updater_controller
import ui_macos as ui
import win_shell

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


APP_VERSION = "2.15.0"
# Оставляем запас под полный путь Windows: лимит 260 символов, минус имя файла
# и расширение, минус служебные суффиксы « (2).docx».

# Масштабирование берём на себя: иначе customtkinter масштабирует повторно.
ScalingTracker.deactivate_automatic_dpi_awareness = True

# Per-Monitor DPI awareness (п. 6.3) — до создания первого окна Tk.
win_shell.enable_dpi_awareness()


# ---------- Кэш шаблонов печатных форм (п. 5.5 этапа 5) ----------
# Шаблон .docx разбирается один раз; на каждый рендер отдаётся deepcopy.
# Раньше предпросмотр при каждом вводе заново читал и разбирал ZIP шаблона.
# Кэш инвалидируется по mtime — правка шаблона подхватывается без перезапуска.


# ---------- Сохранение настроек между запусками ----------
SETTINGS_DIR_NAME = "DogovorGenerator"
MIN_WINDOW_WIDTH = 1120  # логические пиксели
MIN_WINDOW_HEIGHT = 620


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
        # Страница «Реестр»: рендер, поиск, выбор, сверка — RegistryController
        # (v2.4.0). Создаётся ДО построения страниц: build_page вызывается
        # напрямую у контроллера.
        self._registry_flow = registry_controller.RegistryController(self)
        # Спецификация и сохранение документов — DocumentSaveController
        # (v2.5.0). Создаётся ДО построения страниц: build_spec_page
        # вызывается напрямую у контроллера.
        self._doc_save_flow = document_save_controller.DocumentSaveController(self)

        # Страница «Добавить контрагента» и редактирование карточки —
        # CounterpartyController (v2.6.0). Создаётся ДО построения страниц:
        # build_counterparty_page вызывается напрямую у контроллера.
        self._counterparty_flow = counterparty_controller.CounterpartyController(self)

        # Страницы «Договор» и «Доп. соглашение» с отложенными
        # предпросмотрами — DocumentPageController (v2.7.0). Создаётся ДО
        # построения страниц: build_contract_page / build_additional_page
        # вызываются напрямую у контроллера.
        self._doc_page_flow = document_page_controller.DocumentPageController(self)

        # Страница «Настройки», ключ DaData и папка документов —
        # SettingsController (v2.8.0). Создаётся ДО построения страниц:
        # build_settings_page вызывается напрямую у контроллера.
        self._settings_flow = settings_controller.SettingsController(self)

        # Проверка обновлений (2.15.0): тихая автопроверка при старте и
        # кнопка «Проверить обновления» на странице «Настройки» —
        # UpdateController. Создаётся ДО построения страниц: карточка
        # «Обновления» привязывает кнопку к контроллеру.
        self._update_flow = updater_controller.UpdateController(self)

        # Единый «текущий контрагент» и защита несохранённых данных (п. 4.2, 4.4).
        self._current_party_id = None
        self._current_party_vars: dict[str, tk.StringVar] = {}
        self._syncing_party = False
        self._form_baseline: dict[str, tuple | None] = {}
        # Оркестрация контекста — DocumentContextController (v2.9.0):
        # списки контрагентов, перенос выбора между страницами и
        # подтверждение потери изменений. Создаётся ДО построения страниц:
        # refresh_party_combo_values вызывается напрямую у контроллера.
        self._doc_context_flow = document_context_controller.DocumentContextController(self)
        self._registry_search_after = None
        self._toast_after = None
        self._toast_path: str | None = None

        self._place_window()
        ui.make_borderless(self.root)

        self._build_titlebar()
        self._build_body()
        self._build_pages()
        self._doc_context_flow.reset_all_form_baselines()
        self._install_resize_handles()
        self._select_page("contract")
        # Первый запуск: сразу открываем добавление контрагента (п. 4.5).
        if not self.store.active_counterparties():
            self.root.after(250, lambda: self._select_page("counterparty"))
        self.root.after(250, self._refresh_scrollbars)
        self.root.after(600, self._cleanup_temp_documents)
        # Автоочистка старой корзины и сверка «реестр ↔ папка документов».
        self.root.after(900, self._cleanup_expired_trash)
        self.root.after(1500, self._registry_flow.scan_consistency)
        # Тихая автопроверка обновлений (2.15.0): после всех стартовых задач,
        # не чаще раза в сутки, выключается пользователем и env-глушкой
        # EWDOCHUB_NO_UPDATE_CHECK (тесты/smoke не ходят в сеть).
        self.root.after(2200, self._update_flow.maybe_auto_check)

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
        # Курсоры — системные имена Windows (п. 6.3): «top_side» и родственные
        # X-имена не во всех средах resolves, «size_*» есть в любой Windows.
        # TclError при создании хэндла отлавливается ниже — окно остаётся
        # работоспособным и без курсоров растяжения (Xvfb и минималистичные X).
        specs = (
            ("n",  "size_ns",            "chrome",  dict(relx=0.0, rely=0.0, relwidth=1.0, height=t), self.titlebar),
            ("nw", "size_nwse",          "chrome",  dict(width=c, height=c), self.titlebar),
            ("ne", "size_nesw",          "chrome",  dict(relx=1.0, x=-c, width=c, height=c), self.titlebar),
            ("s",  "size_ns",            "window",  dict(relx=0.0, rely=1.0, y=-t, relwidth=1.0, height=t), self.root),
            ("w",  "size_we",            "sidebar", dict(rely=0.0, relheight=1.0, width=t), self.root),
            ("e",  "size_we",            "content", dict(relx=1.0, x=-t, rely=0.0, relheight=1.0, width=t), self.root),
            ("sw", "size_nesw",          "sidebar", dict(rely=1.0, y=-c, width=c, height=c), self.root),
            ("se", "size_nwse",          "window",  dict(relx=1.0, rely=1.0, x=-c, y=-c, width=c, height=c), self.root),
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
        dirty = [self._doc_context_flow.form_label(key) for key in self._doc_context_flow.dirty_pages()]
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
            base = self._settings_flow.documents_dir(interactive=False)
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
            self._registry_flow.refresh()
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
            inner, "Сохранить договор", self._doc_save_flow.save_contract,
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
            "counterparty": ("Сохранить контрагента", self._counterparty_flow.save_card),
            "contract": ("Сохранить договор", self._doc_save_flow.save_contract),
            "additional": ("Сохранить доп. соглашение", self._doc_save_flow.save_additional),
            "spec": ("Сохранить спецификацию", self._doc_save_flow.save_spec),
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
        self._doc_page_flow.build_contract_page()
        self._doc_page_flow.build_additional_page()
        self._doc_save_flow.build_spec_page()
        self._counterparty_flow.build_counterparty_page()
        self._registry_flow.build_page()
        self._settings_flow.build_settings_page()
        # Обновляем списки только после того, как все страницы созданы:
        # иначе в «Доп. соглашении» и «Спецификации» не будет контрагентов.
        self._doc_context_flow.refresh_party_combo_values()

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
    # ------------------------------------------------------------------ #
    #  Страницы «Договор» и «Доп. соглашение» — DocumentPageController
    # ------------------------------------------------------------------ #
    #  Сборка страниц, отложенные предпросмотры (п. 5.6), поиск по ИНН,
    #  подстановка договора-основания (п. 5.1, 4.2) и поля «прописью» —
    #  в document_page_controller.DocumentPageController (v2.7.0). Фасады
    #  ContractApp удалены (v2.12.0): вызовы идут через
    #  self.app._doc_page_flow, состояние страниц осталось на приложении.

    #  Спецификация и сохранение документов — DocumentSaveController
    #  (document_save_controller, v2.5.0). Фасады ContractApp удалены
    #  (v2.12.0): вызовы идут через self.app._doc_save_flow.

    #  Страница «Реестр» — RegistryController (registry_controller,
    #  v2.4.0). Фасады ContractApp удалены (v2.12.0): вызовы идут через
    #  self.app._registry_flow.

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

    #  Контрагенты/договоры на страницах, единый «текущий контрагент»
    #  (п. 4.2) и защита несохранённых данных (п. 4.4) —
    #  DocumentContextController (document_context_controller, v2.9.0).
    #  Фасады ContractApp удалены (v2.12.0): вызовы идут через
    #  self.app._doc_context_flow; состояние контекста (_current_party_id,
    #  _syncing_party, _current_party_vars, _form_baseline) на приложении.

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

    #  Экспорт/открытие документов и корзина — LifecycleController
    #  (lifecycle_controller, v2.3.0). Фасады ContractApp удалены
    #  (v2.12.0): вызовы идут через self.app._lifecycle_flow.

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


    #  Страница «Настройки», ключ DaData и папка документов —
    #  SettingsController (settings_controller, v2.8.0). Фасады ContractApp
    #  удалены (v2.12.0): вызовы идут через self.app._settings_flow.

    def _import_documents_dialog(self):
        """Выбор файлов для загрузки в реестр."""
        if self._import_flow.running:
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
        self._import_flow.start(list(files))

    #  Импорт документов — ImportController (import_controller, v2.1.0).
    #  Фасады и свойства состояния ContractApp удалены (v2.12.0):
    #  состояние потока читается и подменяется через self._import_flow
    #  (running / queue / win / cancelled).

    #  Страница «Добавить контрагента» и окно редактирования карточки —
    #  CounterpartyController (counterparty_controller, v2.6.0). Фасады
    #  ContractApp удалены (v2.12.0): вызовы идут через
    #  self.app._counterparty_flow; состояние страницы (_cp_entries,
    #  _cp_error_vars, _cp_frames, _cp_type_var, _cp_status_var) —
    #  на приложении, его читают DadataController, защита несохранённых
    #  данных и smoke-тест.

    #  DaData: автозаполнение — DadataController (dadata_controller,
    #  v2.2.0). Фасады ContractApp удалены (v2.12.0): вызовы идут через
    #  self.app._dadata_flow; очередь результата опрашивается только из
    #  главного потока Tk — root.after() из рабочего потока запрещён
    #  (аудит 2.0.1).

    #  Подписи форм, baseline, «грязные» страницы, подтверждение при
    #  уходе и сохранение текущей страницы — DocumentContextController
    #  (document_context_controller, v2.9.0); базовые подписи
    #  (_form_baseline) остаются на приложении.

    def _on_save_shortcut(self, _event=None):
        self._doc_context_flow.save_current_page()
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
        self._doc_context_flow.save_current_page()
        return "break"

    # ------------------------------------------------------------------ #
    #  Drag&drop файлов в окно (п. 4.5) и смена DPI монитора (п. 6.3)
    # ------------------------------------------------------------------ #
    def _enable_drop_files(self) -> None:
        """Принимает файлы, перетащенные в окно, через WM_DROPFILES.

        Win32-обвязка (DragAcceptFiles, подкласс WndProc) вынесена в
        win_shell (п. 9.3): один хук на окно обслуживает и WM_DROPFILES,
        и WM_DPICHANGED (п. 6.3). Колбэки уходят в цикл Tk через
        after(0) — выполняются в главном потоке. При любой ошибке
        возможность просто отключается — остаётся кнопка
        «Добавить документы…», а геометрию при смене DPI правит Tk.
        """
        win_shell.install_window_hooks(
            self.root,
            on_files=self._import_flow.start,
            on_dpi_changed=self._on_dpichanged,
        )

    def _on_dpichanged(self, dpi: int, rect: tuple[int, int, int, int]) -> None:
        """Окно перенесли на монитор с другим масштабом (п. 6.3).

        Окно встаёт в подсказанный Windows прямоугольник (иначе после
        переноса вылезает за границы монитора), масштаб интерфейса
        пересчитывается — диалоги, тосты и меню, созданные после
        переноса, используют новый масштаб. Уже построенные страницы под
        новый масштаб не пересобираются: живая пересборка интерфейса —
        отдельная задача (см. «Известные долги» README).
        """
        self.scale = max(1.0, dpi / 96.0)
        ui.set_scale(self.scale)
        x, y, width, height = rect
        if width > 0 and height > 0:
            self.root.geometry(f"{width}x{height}+{x}+{y}")
        self._refresh_scrollbars()
        self._log.info(
            "Смена DPI монитора: %d → масштаб интерфейса %.2f", dpi, self.scale,
        )

    # ------------------------------------------------------------------ #
    #  Навигация и статус
    # ------------------------------------------------------------------ #
    def _select_page(self, key: str):
        if key != self.current_page and not self._doc_context_flow.confirm_leave_page():
            return
        self.current_page = key
        self._update_footer()
        for page_key, frame in self.pages.items():
            if page_key == key:
                frame.tkraise()
        for nav_key, item in self.nav_items.items():
            item.set_selected(nav_key == key)
        if key == "registry":
            self._registry_flow.refresh()
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

    #  «Сумма/срок прописью» и формат лимита — методы
    #  DocumentPageController (update_words / format_limit /
    #  update_limit_words).

    def run(self):
        self.root.mainloop()


def main():
    """Точка входа: один экземпляр, логирование, корректное освобождение блокировки."""
    lock = single_instance.InstanceLock()
    if not lock.acquire():
        # Второй экземпляр писал бы в ту же базу реестра и мог затереть данные.
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
