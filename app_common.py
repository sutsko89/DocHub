"""GUI-free константы и настройки приложения (этап M1 миграции Tk → PySide6).

Развязка GUI-контроллеров от main.py: раньше ширины разметки, печатные
настройки спецификации, версию и хелперы каретки контроллеры брали поздним
``import main`` внутри методов (16 мест — цикл «main импортирует контроллер,
контроллер импортирует main»). Теперь они лежат здесь и импортируются
напрямую. main.py реэкспортирует имена, поэтому smoke_test.py и тесты,
обращающиеся к ``m.APP_VERSION`` / ``m.get_spec_settings`` и т. п., работают
как раньше; Tk-сборка поведенчески не меняется.

Состав (план миграции, этап M1):

* ``APP_VERSION`` — единый источник версии приложения: его читает
  EWDochub.spec при сборке (VersionInfo exe) и обновляет
  tools/bump_version.py. Строка остаётся чисто числовой «X.Y.Z» —
  регулярка spec-файла и updater.parse_version не ломаются;
* ``APP_UI_FLAVOR`` — признак интерфейса сборки (``"tk"`` | ``"qt"``)
  для канала обновлений (план миграции, раздел И.2): сравнение версий
  идёт по числовой части, флейвор в сравнение не входит;
* константы разметки, которые используют контроллеры страниц
  (ENTRY_WIDTH / SHORT_WIDTH / LEFT_PANEL_WIDTH);
* чтение и атомарная запись settings.json (settings_path / load_settings /
  write_settings) и печатные настройки спецификации (DEFAULT_SPEC_SETTINGS /
  get_spec_settings) — с прежним клампом ``spec_vat_rate`` и годом
  «на момент вызова» (аудит 2.0.1);
* хелперы каретки и телефона (restore_caret_after_reformat /
  make_phone_formatter) — Tk-виджеты передаются им аргументами, сам модуль
  tkinter не импортирует (константа ``tk.INSERT`` заменена равнозначным
  литералом ``"insert"``).

Модуль не импортирует main и не зависит от Tk/CTk: ядро (storage, services,
document_builder, updater) остаётся без изменений (план миграции, В.3).
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime

from document_builder import (
    SPEC_APPENDIX_NUMBER, SPEC_VAT_RATE, digits_only, format_phone,
)

APP_VERSION = "2.28.2"

# Признак интерфейса сборки (план миграции, И.2): "tk" — Tk/CustomTkinter
# сборка, "qt" — сборка на PySide6 (каркас появляется на этапе M2).
# Отображается как «2.28.2-tk»; updater.parse_version сравнивает числовую
# часть, поэтому флейвор не делает сборку «новее» той же версии.
APP_UI_FLAVOR = "tk"

# ---------- Константы разметки, используемые контроллерами страниц ----------
# Логические пиксели: умножаются на app.scale в месте использования.
# Геометрия окна Tk-сборки (DAY_WIDTH, SIDEBAR_WIDTH, BASE_*, MIN_WINDOW_*)
# остаётся в main.py — она принадлежит Tk-каркасу и уходит вместе с ним (M8).
ENTRY_WIDTH = 300
SHORT_WIDTH = 150
LEFT_PANEL_WIDTH = 400

# ---------- Сохранение настроек между запусками ----------
SETTINGS_DIR_NAME = "DogovorGenerator"


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


# ---------- Хелперы каретки и телефона ----------


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
                # tk.INSERT == "insert": литерал вместо константы tkinter,
                # чтобы модуль оставался GUI-free (M1) — значение то же.
                caret = entry.index("insert")
            except Exception:
                caret = None
        var.set(formatted)
        restore_caret_after_reformat(entry, caret, len(raw), len(formatted))

    return _format
