"""Чистое ядро таблицы позиций спецификации (этап M3 миграции Tk → PySide6).

Логика вставки блока из Excel, нормализация чисел и валидаторы ячеек
вынесены из Tk-виджета widgets.SpecTable в GUI-free модуль, чтобы Qt-таблица
(spec_table_qt.SpecTableQt) и Tk-таблица считали одинаково: один
``paste_plan`` (PERF-1), одна нормализация «1 000,50» → «1000.50», одни
правила санитизации вставки. Расхождение лимита или правил вставки между
сборками исключено по построению — тот же приём, что у spec_total (U12).

Зависимостей от интерфейса нет: только document_builder.parse_decimal.
Функции чистые — покрыты существующими тестами test_spec_table_validation
(через делегирование widgets.SpecTable) и новыми Qt-тестами CP1.
"""

from __future__ import annotations

from document_builder import parse_decimal

# Лимиты и состав колонок — прежние значения Tk-таблицы (PERF-1, D2/R-08).
MIN_ROWS = 10
MAX_ROWS = 500
HEADERS = ("№", "Наименование", "Кол-во", "Ед. изм.", "Цена за ед. с НДС (руб.)")
KEYS = ("name", "qty", "unit", "price")
# Ширина фиксированных колонок в логических пикселях (№, Кол-во, Ед. изм., Цена).
COL_WIDTHS = {0: 34, 2: 64, 3: 84, 4: 150}


def paste_plan(matrix: list[list[str]], start_row: int,
               max_rows: int) -> tuple[list[list[str]], int]:
    """Обрезает вставляемый блок до лимита таблицы (PERF-1).

    Возвращает ``(usable, dropped)`` — строки, которые встанут в таблицу
    (итого после вставки не более ``max_rows`` с учётом стартовой ячейки
    ``start_row``), и число отброшенных строк. Логика совпадает с полем
    «Строк:» (set_row_count ограничена тем же MAX_ROWS).
    """
    capacity = max(0, max_rows - max(0, start_row))
    if len(matrix) <= capacity:
        return matrix, 0
    return matrix[:capacity], len(matrix) - capacity


def paste_notice_text(inserted: int, dropped: int, max_rows: int) -> str:
    """Сообщение об обрезке вставки: «Вставлено 500, отброшено 100 …»."""
    return (
        f"Вставлено {inserted}, отброшено {dropped}: лимит таблицы — "
        f"{max_rows} строк. Остальные строки остались в источнике и в "
        "буфере обмена."
    )


def normalize_number(text: str) -> str:
    """«1 000,50» → «1000.50»: убирает разрядные разделители,
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


def validate_digits(new_value: str) -> bool:
    """Кол-во: цифры и запятая/точка (может быть дробным)."""
    return new_value == "" or all(ch.isdigit() or ch in ",." for ch in new_value)


def validate_price(new_value: str) -> bool:
    """Цена: цифры, разряды и запятая/точка (копейки)."""
    return new_value == "" or all(
        ch.isdigit() or ch in " .," for ch in new_value
    )


def sanitize_value(key: str, value: str) -> str:
    """Приводит значение к правилам ячейки (вставка мимо валидатора ввода)."""
    if key in ("qty", "price"):
        value = normalize_number(value)
        validator = validate_digits if key == "qty" else validate_price
        return value if validator(value) else ""
    return value


def map_row(cells: list[str]) -> list[str]:
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
        return [cells[0], normalize_number(cells[1]), "", ""]
    if n == 3:
        # «Наименование, Кол-во, Цена» либо «Наименование, Кол-во, Ед. изм.»
        if normalize_number(cells[2]):
            return [cells[0], normalize_number(cells[1]), "",
                    normalize_number(cells[2])]
        return [cells[0], normalize_number(cells[1]), cells[2], ""]
    return [cells[0], normalize_number(cells[1]), cells[2],
            normalize_number(cells[3])]


def items_from_values(values: list[dict]) -> list[dict]:
    """Позиции с заполненным наименованием; суммы считаются на лету.

    Кол-во и цена — Decimal: цена может содержать копейки, кол-во может
    быть дробным (например, 2,5 м). Вход — значения строк таблицы
    ({«name», «qty», «unit», «price»}) в любом видежет-слое.
    """
    items = []
    for row in values:
        name = str(row.get("name", "") or "").strip()
        if not name:
            continue
        items.append({
            "name": name,
            "qty": parse_decimal(row.get("qty")),
            "unit": str(row.get("unit", "") or "").strip() or "шт.",
            "price": parse_decimal(row.get("price")),
        })
    return items
