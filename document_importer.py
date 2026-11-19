"""Анализ загружаемых пользователем документов (функция «Добавить документы»).

Модуль работает в фоновом потоке и умеет:
  * извлекать текст из .docx (python-docx) и .xlsx (openpyxl);
  * определять тип документа: договор поставки, доп. соглашение, спецификация;
  * находить реквизиты контрагента (название, ИНН, КПП, ОГРН/ОГРНИП, адрес,
    контакты, банковские реквизиты);
  * находить номер и дату документа.

Результат :func:`analyze_file` — словарь::

    {"path", "name", "kind", "number", "date", "party", "text", "error", "ok"}

``party`` — словарь {"recognized": bool, "type": "ooo"|"ip",
"fields": {...}, "display": str}. Ключи в ``fields`` совпадают с ключами
класса ``F`` из main.py (единые имена полей карточки контрагента).

Если документ опознать не удалось (нет признаков типа документа или
реквизитов контрагента), ``ok=False`` — файл попадает в папку «unknown».

Гарантии качества импорта (этап 3 плана):
  * ``recognized=True`` только если найден ИНН и подтверждающие данные
    (название/КПП/ОГРН для юрлица, ФИО для ИП) — «полукопии» без ИНН
    требуют подтверждения пользователя;
  * размер файла и содержимое ZIP проверяются до разбора (защита от
    «архивных бомб»); XML в .docx разбирается python-docx с
    ``resolve_entities=False`` (XXE закрыт), openpyxl использует
    stdlib ElementTree без внешних сущностей;
  * у каждой ошибки — понятная пользователю причина.
"""

from __future__ import annotations

import os
import re
import zipfile

import applog
from document_builder import F

UNKNOWN_FOLDER = "unknown"

KIND_CONTRACT = "contract"
KIND_ADDITIONAL = "additional"
KIND_SPEC = "spec"
KIND_UNKNOWN = "unknown"

# Ключи полей контрагента — единый класс F из document_builder (импорт выше).
# Прежний «монкипатч» с переприсваиванием 18 атрибутов F удалён (аудит 2.0.1,
# P2): значения совпадали, но побочная запись в общий класс — мина при любой
# правке ключа. Совпадение источников закреплено smoke-тестом (di.F is m.F)
# и golden-формами.

# Сколько символов текста сохраняем в записи реестра (для просмотра).
# Ограничение применяется до сохранения: полный текст большого документа
# в реестр не попадает (п. 3.4).
TEXT_KEEP_LIMIT = 100_000

# Лимиты импорта (п. 3.4): защищают от «архивных бомб» и гигантских файлов.
MAX_FILE_SIZE = 50 * 1024 * 1024            # 50 МБ на один файл
MAX_ZIP_MEMBERS = 5_000                     # число частей внутри .docx/.xlsx
MAX_ZIP_UNCOMPRESSED_SIZE = 300 * 1024 * 1024   # 300 МБ после распаковки
MAX_ZIP_RATIO = 200                         # сжатие сильнее 1:200 — подозрительно
ZIP_RATIO_MIN_SIZE = 20 * 1024 * 1024       # порог, ниже которого коэффициент не считаем
MAX_FILES_PER_BATCH = 200                   # файлов за один импорт
MAX_TOTAL_SIZE = 300 * 1024 * 1024          # суммарный объём пакета
MIN_FREE_SPACE = 100 * 1024 * 1024          # запас свободного места на диске

_OFFICE_EXTENSIONS = (".docx", ".xlsx", ".xlsm")


def check_file_limits(path: str) -> None:
    """Проверяет размер файла и защищает от ZIP-бомб (п. 3.4).

    Бросает ``ValueError`` с понятной пользователю причиной — она попадёт
    в отчёт импорта. Проверка выполняется до извлечения текста, поэтому
    «архивная бомба» не успевает раздуть память.
    """
    try:
        size = os.path.getsize(path)
    except OSError as exc:
        raise ValueError(f"Не удалось прочитать файл: {exc}") from exc
    if size <= 0:
        raise ValueError("Файл пустой (0 байт)")
    if size > MAX_FILE_SIZE:
        limit_mb = MAX_FILE_SIZE / (1024 * 1024)
        raise ValueError(
            f"Файл слишком большой: {size / (1024 * 1024):.1f} МБ "
            f"(допустимо до {limit_mb:.0f} МБ)"
        )

    ext = os.path.splitext(path)[1].lower()
    if ext not in _OFFICE_EXTENSIONS:
        return  # офисный архив проверяем ниже; прочие расширения отсеет extract_text

    try:
        with zipfile.ZipFile(path) as archive:
            members = archive.infolist()
    except zipfile.BadZipFile as exc:
        raise ValueError(
            "Файл повреждён: это не корректный документ Office (.docx/.xlsx)"
        ) from exc
    except OSError as exc:
        raise ValueError(f"Не удалось открыть файл: {exc}") from exc

    if len(members) > MAX_ZIP_MEMBERS:
        raise ValueError(
            f"Внутри файла слишком много частей ({len(members)}) — "
            "документ отклонён как подозрительный"
        )
    unpacked = sum(max(0, info.file_size) for info in members)
    if unpacked > MAX_ZIP_UNCOMPRESSED_SIZE:
        raise ValueError(
            "После распаковки документ слишком велик "
            f"({unpacked / (1024 * 1024):.0f} МБ) — вероятна «архивная бомба»"
        )
    if (
        unpacked > ZIP_RATIO_MIN_SIZE
        and size
        and unpacked / size > MAX_ZIP_RATIO
    ):
        raise ValueError(
            "Файл сжат подозрительно сильно (похоже на «архивную бомбу»)"
        )


# --------------------------------------------------------------------------- #
#  Извлечение текста
# --------------------------------------------------------------------------- #
def extract_text(path: str) -> str:
    """Извлекает текст документа; неподдерживаемый формат — исключение."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return _extract_docx(path)
    if ext in (".xlsx", ".xlsm"):
        return _extract_xlsx(path)
    if ext == ".doc":
        raise ValueError(
            "Файлы старого формата .doc не поддерживаются — "
            "пересохраните документ в .docx"
        )
    raise ValueError(f"Неподдерживаемый формат файла: {ext or 'без расширения'}")


def _extract_docx(path: str) -> str:
    from docx import Document

    doc = Document(path)
    chunks: list[str] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            chunks.append(text)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            line = " | ".join(cell for cell in cells if cell)
            if line:
                chunks.append(line)
    return "\n".join(chunks)


def _extract_xlsx(path: str) -> str:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    chunks: list[str] = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [
                    str(value).strip()
                    for value in row
                    if value is not None and str(value).strip()
                ]
                if cells:
                    chunks.append(" | ".join(cells))
    finally:
        wb.close()
    return "\n".join(chunks)


# --------------------------------------------------------------------------- #
#  Определение типа документа
# --------------------------------------------------------------------------- #
_TITLE_RE = re.compile(
    r"^\s*\**\s*(?:договор\s+\S+|доп(?:\.|олнител\w*)\s*соглашени\w*|спецификац\w*)",
    re.IGNORECASE,
)


def detect_kind(text: str) -> str:
    """Тип документа: сначала по строке-заголовку, затем по частоте слов.

    Порядок важен: в договоре и спецификации всегда упоминаются соседние
    документы, поэтому ищем именно строку заголовка (первая строка вида
    «ДОГОВОР ПОСТАВКИ №…», «СПЕЦИФИКАЦИЯ №…», «ДОПОЛНИТЕЛЬНОЕ СОГЛАШЕНИЕ №…»).
    """
    for line in text.splitlines()[:40]:
        match = _TITLE_RE.match(line)
        if match:
            title = match.group(0).lower()
            if title.startswith("договор"):
                return KIND_CONTRACT
            if title.startswith("доп"):
                return KIND_ADDITIONAL
            return KIND_SPEC

    # Заголовок не найден — считаем вхождения ключевых слов в начале текста.
    head = text[:6000].lower()
    counts = {
        KIND_ADDITIONAL: len(re.findall(r"доп(?:\.|олнител\w*)\s*соглашени", head)),
        KIND_SPEC: len(re.findall(r"спецификац", head)),
        KIND_CONTRACT: len(re.findall(r"договор", head)),
    }
    best = max(counts, key=counts.get)
    return best if counts[best] else KIND_UNKNOWN


# --------------------------------------------------------------------------- #
#  Регулярные выражения для реквизитов
# --------------------------------------------------------------------------- #
_INN_LABELED = re.compile(r"ИНН[^\d\n]{0,12}(\d{12}|\d{10})", re.IGNORECASE)
_INN_OOO_RAW = re.compile(r"(?<!\d)(\d{10})(?!\d)")
_INN_IP_RAW = re.compile(r"(?<!\d)(\d{12})(?!\d)")
_KPP_RE = re.compile(r"КПП[^\d\n]{0,12}(\d{9})", re.IGNORECASE)
_OGRNIP_RE = re.compile(r"ОГРНИП[^\d\n]{0,12}(\d{15})", re.IGNORECASE)
_OGRN_RE = re.compile(r"ОГРН[^\d\n]{0,12}(\d{13})", re.IGNORECASE)
# Написания расчётного счёта в реальных документах (п. 3.1):
# «Р/с», «Р/счёт», «р/сч», «расчётный счёт», «расч. счёт», «счёт №».
_ACCOUNT_RE = re.compile(
    r"(?:"
    r"р\s*[\/.]?\s*с(?:ч(?:[её]т\w*)?)?\b"
    r"|расч[её]тн\w*\s+сч[её]т\w*"
    r"|расч\.?\s*сч[её]т\w*"
    r"|(?<![\w\/.])(?:сч[её]т|счет)\s*(?:№|номер)"
    r")"
    r"[^\d\n]{0,15}(\d{20})",
    re.IGNORECASE,
)
_CORR_RE = re.compile(
    r"к\s*[\/.]?\s*с(?:ч(?:[её]т\w*)?)?\b[^\d\n]{0,15}(\d{20})",
    re.IGNORECASE,
)
# «20 цифр в отдельной ячейке таблицы»: счёт без подписи (п. 3.1).
_ACCOUNT_RAW_RE = re.compile(r"(?<!\d)(\d{20})(?!\d)")
_BIK_RE = re.compile(r"БИК[^\d\n]{0,12}(\d{9})", re.IGNORECASE)
_BANK_RE = re.compile(
    r"(?:банк|банк\s+получателя)\s*[:\-]\s*([^\n|]{3,100})", re.IGNORECASE
)
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
_PHONE_RE = re.compile(
    r"(?:телефон|тел\.?|факс)\s*[:\-]?\s*(\+?[\d\s\-()]{7,22}\d)", re.IGNORECASE
)

_OOO_FULL_RE = re.compile(
    r"общество\s+с\s+ограниченной\s+ответственностью\s+«?([^«»\n|]{2,100})»?",
    re.IGNORECASE,
)
_OOO_SHORT_RE = re.compile(
    r"(?<![А-ЯЁа-яёA-Za-z])ООО\s*[«\"“”]?([A-ZА-ЯЁ0-9][^»\"“”|\n]{1,90})"
)
# Имя (ФИО) без флага IGNORECASE: с ним [А-ЯЁ] соответствует и строчным буквам.
_NAME = r"[А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){1,2}"

_IP_RE = re.compile(
    # после «Индивидуальный предприниматель» тоже нужен разделитель:
    # без \s+ имя ИП после полного написания не извлекалось
    r"(?:[Ии]ндивидуальный\s+[Пп]редприниматель\s+|(?<![А-ЯЁа-яё])ИП\s+)" + rf"({_NAME})"
)
_DIRECTOR_RE = re.compile(
    r"(?:[Гг]енеральн\w*\s+)?(?:[Дд]иректор\w*|[Уу]правляющий)\s+" + rf"({_NAME})"
)
_DIRECTOR_PERSON_RE = re.compile(
    r"[Вв]\s+лице\s+(?:[Гг]енеральн\w*\s+)?"
    r"(?:[Дд]иректор\w*|[Уу]правляющ\w*|[Рр]уководител\w*)?\s*"
    r"(?:филиал\w*)?\s*(?:в\s+г\.\s*[\w\-\s]+\s*,?\s*)?(?:,\s*)?" + rf"({_NAME})"
)
_ADDRESS_RE = re.compile(
    r"(?:юридический|фактический|почтовый)?\s*адрес\w*\s*[:\-]\s*([^\n|]{10,200})",
    re.IGNORECASE,
)
_NUMBER_RE = re.compile(r"№\s*([0-9A-Za-zА-Яа-яЁё\/\-\.]{1,20})")

_MONTH_PREFIXES = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма[йя]": 5, "июн": 6,
    "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11,
    "декабр": 12,
}


def _search(pattern, text: str) -> str:
    match = pattern.search(text)
    if not match:
        return ""
    # У шаблонов без группы берём совпадение целиком (например, Email).
    return (match.group(1) if match.groups() else match.group(0)).strip()


def _clean_value(value: str) -> str:
    """Подрезает «хвосты» значений, попавшие из смежных ячеек таблиц."""
    value = value.split("|", 1)[0]
    value = re.sub(r"\s+", " ", value).strip(" \t\n\r,;:")
    return value.strip()


# Маркеры, после которых в названии начинается описание стороны, а не имя.
_NAME_TAIL_RE = re.compile(
    r"(?:"
    r"[,;]?\s*\(?\s*(?:сокращ[её]нн\w*|полн\w*)\s+наименован\w*"
    r"|[,;]?\s*именуем\w*"
    r"|[,;]?\s*в\s+лице\b"
    r"|[,;]?\s*\(?\s*далее\b"
    r"|[,;]?\s*действующ\w*"
    r"|[,;]?\s*на\s+основании\b"
    r")",
    re.IGNORECASE,
)


def _normalize_quotes(value: str) -> str:
    """Приводит кавычки к «ёлочкам»: ``ООО "X"`` -> ``ООО «X»`` (п. 3.2)."""
    value = value.replace("“", "«").replace("”", "»").replace("„", "«")
    if '"' not in value:
        return value
    parts = value.split('"')
    if len(parts) % 2 == 1:  # чётное число кавычек — парные
        rebuilt = [parts[0]]
        for index, part in enumerate(parts[1:]):
            rebuilt.append(("«" if index % 2 == 0 else "»") + part)
        value = "".join(rebuilt)
    else:
        value = value.replace('"', "")
    return value


def _clean_company_name(value: str) -> str:
    """Убирает «хвост» фразы о стороне и нормализует кавычки (п. 3.2).

    ``ООО "ЛИДЕР ДВ", именуемый в дальнейшем`` -> ``ООО «ЛИДЕР ДВ»``.
    """
    value = _clean_value(value)
    value = value.split("\n", 1)[0]
    match = _NAME_TAIL_RE.search(value)
    if match and match.start() > 0:
        value = value[: match.start()]
    value = value.strip(" \t\n\r,;:.-—")
    return _normalize_quotes(value).strip()


def _format_ooo_short(inner: str) -> str:
    """Короткое название в едином виде: ``ООО «X»``."""
    inner = _normalize_quotes(inner).strip()
    if not inner:
        return ""
    if inner.upper().startswith("ООО"):
        return inner
    if inner.lower().startswith("общество с ограниченной ответственностью"):
        return inner  # краткое название не найдено — полное лучше, чем ничего
    return f"ООО «{inner.strip('«»')}»"


def _format_ooo_full(inner: str) -> str:
    """Полное название в едином виде: ``Общество ... «X»``."""
    inner = _normalize_quotes(inner).strip()
    if not inner:
        return ""
    if inner.lower().startswith("общество с ограниченной ответственностью"):
        return inner
    if inner.upper().startswith("ООО"):
        return inner
    return (
        "Общество с ограниченной ответственностью "
        f"«{inner.strip('«»')}»"
    )


# --------------------------------------------------------------------------- #
#  Номер и дата документа
# --------------------------------------------------------------------------- #
def extract_number(text: str) -> str:
    """Первый номер, содержащий цифры: «Договор № 516», «Спецификация № 12-ф»."""
    for match in _NUMBER_RE.finditer(text[:6000]):
        candidate = match.group(1).strip(".")
        if any(ch.isdigit() for ch in candidate):
            return candidate
    return ""


def extract_date(text: str) -> str:
    """Дата документа в формате ДД.ММ.ГГГГ; не найдена — пустая строка."""
    match = re.search(r"«\s*(\d{1,2})\s*»\s*([а-яё]+)\s*(\d{4})", text, re.IGNORECASE)
    if match:
        return _date_from_words(match.group(1), match.group(2), match.group(3))
    match = re.search(r"(?<!\d)(\d{1,2})\.(\d{1,2})\.(\d{4})(?!\d)", text)
    if match:
        day, month = int(match.group(1)), int(match.group(2))
        if 1 <= day <= 31 and 1 <= month <= 12:
            return f"{day:02d}.{month:02d}.{match.group(3)}"
    match = re.search(r"(?<!\d)(\d{1,2})\s+([а-яё]+)\s+(\d{4})(?!\d)", text, re.IGNORECASE)
    if match:
        return _date_from_words(match.group(1), match.group(2), match.group(3))
    return ""


def _date_from_words(day: str, month_word: str, year: str) -> str:
    month_word = month_word.lower()
    for prefix, month_number in _MONTH_PREFIXES.items():
        if re.match(prefix, month_word):
            try:
                return f"{int(day):02d}.{month_number:02d}.{year}"
            except ValueError:
                pass  # некорректный день — пробуем дальше
    return ""


# --------------------------------------------------------------------------- #
#  Реквизиты контрагента
# --------------------------------------------------------------------------- #
def _matches(pattern, text: str) -> list[tuple[str, int]]:
    """Все совпадения: (значение, позиция в тексте)."""
    result: list[tuple[str, int]] = []
    for match in pattern.finditer(text):
        value = match.group(1) if match.groups() else match.group(0)
        result.append((value, match.start()))
    return result


def _nearest(pattern, text: str, anchor: int | None, exclude=()) -> str:
    """Значение, ближайшее к якорю (блоку реквизитов контрагента)."""
    candidates = [
        (value, position)
        for value, position in _matches(pattern, text)
        if value.strip() and value not in exclude
    ]
    if not candidates:
        return ""
    if anchor is None:
        return candidates[0][0]
    return min(candidates, key=lambda item: abs(item[1] - anchor))[0]


_BUYER_MARKER_RE = re.compile(
    r"именуем\w*\s+(?:в\s+дальнейшем\s+)?[«\"“”]?\s*покупател\w*",
    re.IGNORECASE,
)
_IP_MARKER_RE = re.compile(
    r"именуем\w*\s+(?:в\s+дальнейшем\s+)?[«\"“”]?\s*ИП\b",
    re.IGNORECASE,
)


def _pick_ooo_name(text: str) -> tuple[str, str, int | None]:
    """Возвращает (краткое, полное, позиция-якорь) название ООО.

    Контрагент в документах программы — сторона «Покупатель», поэтому ищем
    фразу «именуемое в дальнейшем «Покупатель»» (в т. ч. без кавычек) и берём
    название из её окружения. Якорь — последнее вхождение названия (обычно
    блок реквизитов в конце документа) — по нему выбираем ближайшие реквизиты.
    """
    marker = _BUYER_MARKER_RE.search(text)
    if marker:
        window = text[max(0, marker.start() - 700): marker.end()]
        shorts = _matches(_OOO_SHORT_RE, window)
        fulls = _matches(_OOO_FULL_RE, window)
        if shorts or fulls:
            # Якорь ищем по «сырому» значению: очистка меняет кавычки,
            # и rfind по очищенной строке не нашёл бы её в тексте.
            raw_short = _clean_value(shorts[-1][0]) if shorts else ""
            raw_full = _clean_value(fulls[-1][0]) if fulls else ""
            inner_short = _clean_company_name(raw_short) if raw_short else ""
            inner_full = _clean_company_name(raw_full) if raw_full else ""
            if not inner_short and inner_full:
                inner_short = inner_full
            short = _format_ooo_short(inner_short)
            full = _format_ooo_full(inner_full) if inner_full else short
            anchor = _resolve_anchor(text, raw_short or raw_full, marker.start())
            return short, full, anchor

    # Маркер «Покупатель» не найден — берём первое упомянутое название.
    shorts = _matches(_OOO_SHORT_RE, text)
    fulls = _matches(_OOO_FULL_RE, text)
    raw_short = _clean_value(shorts[0][0]) if shorts else ""
    raw_full = _clean_value(fulls[0][0]) if fulls else ""
    inner_short = _clean_company_name(raw_short) if raw_short else ""
    inner_full = _clean_company_name(raw_full) if raw_full else ""
    short = _format_ooo_short(inner_short)
    full = _format_ooo_full(inner_full) if inner_full else short
    anchor = _resolve_anchor(text, raw_short or raw_full, None)
    return short, full, anchor


def _resolve_anchor(text: str, name: str, fallback: int | None) -> int | None:
    """Позиция названия контрагента в блоке реквизитов.

    Если название в реквизитах записано иначе (например, в преамбуле —
    «Строительно-монтажное управление - 7», а в таблице — «ООО "СМУ-7"»),
    берём последний блок организации: в договорах и спецификациях реквизиты
    покупателя обычно идут последними (п. 3.1, 3.3).
    """
    position = text.rfind(name) if name else -1
    if position >= 0:
        return position
    positions = [
        pos for _value, pos in (
            _matches(_OOO_SHORT_RE, text) + _matches(_OOO_FULL_RE, text)
        )
    ]
    if positions:
        return max(positions)
    return fallback


def _pick_ip_name(text: str) -> tuple[str, int | None]:
    """Возвращает (ФИО ИП, позиция-якорь)."""
    marker = _IP_MARKER_RE.search(text)
    if marker:
        window = text[max(0, marker.start() - 700): marker.end()]
        matches = _matches(_IP_RE, window)
        if matches:
            raw_name = _clean_value(matches[-1][0])
            return _clean_company_name(raw_name), text.rfind(raw_name)
    matches = _matches(_IP_RE, text)
    if matches:
        raw_name = _clean_value(matches[0][0])
        return _clean_company_name(raw_name), text.rfind(raw_name)
    return "", None


# Длина «своего» блока реквизитов после названия контрагента и начало
# следующего блока (другая организация) — граница для поиска счетов.
_REQ_SEGMENT_LENGTH = 1500
_ORG_START_RE = re.compile(
    r"\n\s*(?:ООО|ИП|АО|ПАО|ЗАО|Общество\s+с\s+ограниченной)", re.IGNORECASE
)


def _party_segment(text: str, anchor: int | None) -> str:
    """Фрагмент текста с реквизитами именно этого контрагента.

    В договоре два блока реквизитов (поставщик и покупатель), и «ближайший»
    счёт может принадлежать другой стороне. Поэтому финансовые реквизиты
    ищем в блоке, который начинается с названия контрагента и заканчивается
    перед следующей организацией (п. 3.1).
    """
    if anchor is None or anchor < 0:
        return text
    segment = text[anchor: anchor + _REQ_SEGMENT_LENGTH]
    match = _ORG_START_RE.search(segment, 5)
    if match:
        segment = segment[: match.start()]
    return segment


def _pick_account_by_position(text: str, anchor: int | None,
                              exclude=frozenset()) -> str:
    """Счёт без подписи: 20 цифр в отдельной ячейке таблицы (п. 3.1).

    Берём значение, ближайшее к блоку реквизитов контрагента; корр. счёт,
    БИК и коды уже исключены вызывающей стороной.
    """
    candidates = [
        (value, position)
        for value, position in _matches(_ACCOUNT_RAW_RE, text)
        if value not in exclude
    ]
    if not candidates:
        return ""
    if anchor is None:
        return candidates[0][0]
    return min(candidates, key=lambda item: abs(item[1] - anchor))[0]


def extract_party(text: str, prefer_last_block: bool = False) -> dict | None:
    """Извлекает реквизиты контрагента; None — если нет ни ИНН, ни названия.

    ``prefer_last_block`` — для табличных спецификаций (.xlsx): реквизиты
    покупателя в них идут последним блоком, а название в преамбуле может
    отличаться от написания в таблице (п. 3.1).
    """
    ogrnip = _search(_OGRNIP_RE, text)
    ogrn = _search(_OGRN_RE, text)

    is_ip = bool(
        ogrnip
        or re.search(r"индивидуальный\s+предприниматель", text, re.IGNORECASE)
        or re.search(r"(?<![А-ЯЁа-яё])ИП(?![А-ЯЁа-яё])", text)
    )

    anchor: int | None = None
    if is_ip:
        ip_name, anchor = _pick_ip_name(text)
        short_name = full_name = ""
    else:
        short_name, full_name, anchor = _pick_ooo_name(text)
        ip_name = ""

    if prefer_last_block:
        positions = [
            pos for _value, pos in (
                _matches(_OOO_SHORT_RE, text) + _matches(_OOO_FULL_RE, text)
                + (_matches(_IP_RE, text) if is_ip else [])
            )
        ]
        if positions:
            anchor = max(positions)

    # Реквизиты берём из блока именно этого контрагента: в договоре они есть
    # у обеих сторон, и «ближайший» счёт может оказаться чужим.
    segment = _party_segment(text, anchor)
    # Если в своём блоке есть финансовые реквизиты, значения из чужих блоков
    # (другая сторона договора) не подставляем — лучше пустое поле.
    has_finance = bool(
        _search(_INN_LABELED, segment)
        or _search(_ACCOUNT_RE, segment)
        or _search(_CORR_RE, segment)
        or _search(_BIK_RE, segment)
    )

    def first(pattern, exclude=()):
        """Значение из своего блока; иначе — ближайшее по всему тексту."""
        value = _search(pattern, segment)
        if value and value not in exclude:
            return value
        if has_finance:
            return ""
        return _nearest(pattern, text, anchor, exclude=exclude)

    inn = first(_INN_LABELED)
    kpp = first(_KPP_RE)
    ogrn = first(_OGRN_RE) if not is_ip else ""
    ogrnip = first(_OGRNIP_RE) if is_ip else ""
    bik = first(_BIK_RE)
    # Корр. счёт ищем первым: иначе «счёт № 301…» попадёт в расчётный.
    corr_account = first(_CORR_RE)
    account = first(_ACCOUNT_RE, exclude=(corr_account,) if corr_account else ())
    if not account:
        # «20 цифр в отдельной ячейке» — счёт без подписи (п. 3.1).
        exclude = {v for v in (corr_account, bik, inn, kpp, ogrn, ogrnip) if v}
        account = _pick_account_by_position(segment, None, exclude=exclude)
        if not account:
            account = _pick_account_by_position(text, anchor, exclude=exclude)
    bank = _clean_value(_nearest(_BANK_RE, text, anchor))
    email = _clean_value(_nearest(_EMAIL_RE, text, anchor))
    phone = _clean_value(_nearest(_PHONE_RE, text, anchor))
    address = _clean_value(_nearest(_ADDRESS_RE, text, anchor))
    director_matches = _matches(_DIRECTOR_RE, text) + _matches(_DIRECTOR_PERSON_RE, text)
    if director_matches:
        if anchor is None:
            director = _clean_value(director_matches[0][0])
        else:
            director = _clean_value(
                min(director_matches, key=lambda item: abs(item[1] - anchor))[0]
            )
    else:
        director = ""

    if not inn:
        inn = _guess_inn(text, kpp, ogrn, ogrnip, bik, account, corr_account,
                         is_ip=is_ip, anchor=anchor)

    # Критерий «карточку можно создать автоматически» (п. 3.3):
    # ИНН + подтверждающие данные. Без ИНН документ не должен плодить
    # «полукопии» — пользователь подтверждает карточку вручную.
    if is_ip:
        fields = {
            F.IP_NAME: ip_name, F.INN: inn, F.OGRNIP: ogrnip,
            F.LEGAL_ADDRESS: address, F.FACT_ADDRESS: "",
            F.EMAIL: email, F.PHONE: phone,
            F.ACCOUNT: account, F.CORR_ACCOUNT: corr_account,
            F.BANK: bank, F.BIK: bik,
        }
        display = f"ИП {ip_name}".strip()
        recognized = bool(inn and ip_name)
        missing = "" if recognized else ("ИНН" if not inn else "ФИО ИП")
    else:
        if not full_name and short_name:
            full_name = short_name
        fields = {
            F.FULL_NAME: full_name, F.SHORT_NAME: short_name, F.DIRECTOR: director,
            F.INN: inn, F.KPP: kpp, F.OGRN: ogrn,
            F.LEGAL_ADDRESS: address, F.FACT_ADDRESS: "",
            F.EMAIL: email, F.PHONE: phone,
            F.ACCOUNT: account, F.CORR_ACCOUNT: corr_account,
            F.BANK: bank, F.BIK: bik,
        }
        display = short_name or full_name or ""
        recognized = bool(inn and (short_name or full_name or kpp or ogrn))
        missing = "" if recognized else ("ИНН" if not inn else "название организации")

    if not display and not inn:
        return None
    return {
        "recognized": recognized,
        "type": "ip" if is_ip else "ooo",
        "fields": fields,
        "display": display,
        "inn": inn,
        "missing": missing,
    }


def _guess_inn(text: str, kpp: str, ogrn: str, ogrnip: str, bik: str,
               account: str, corr_account: str, is_ip: bool,
               anchor: int | None = None) -> str:
    """ИНН без пометки: исключаем КПП, ОГРН, счета и БИК."""
    excluded = {v for v in (kpp, ogrn, ogrnip, bik, account, corr_account) if v}
    if ogrnip:
        excluded.add(ogrnip[:12])
    pattern = _INN_IP_RAW if is_ip else _INN_OOO_RAW
    candidates = [
        (match.group(1), match.start())
        for match in pattern.finditer(text)
        if match.group(1) not in excluded
    ]
    if not candidates:
        return ""
    if anchor is None:
        return candidates[0][0]
    return min(candidates, key=lambda item: abs(item[1] - anchor))[0]


def party_display_name(party: dict) -> str:
    """Имя для папки: «ООО ЛИДЕР ДВ» / «ИП Иванов»."""
    return (party.get("display") or "Контрагент").strip() or "Контрагент"


# --------------------------------------------------------------------------- #
#  Итоговый анализ файла
# --------------------------------------------------------------------------- #
def analyze_file(path: str) -> dict:
    """Полный разбор файла. Никогда не бросает исключений — ошибки в «error».

    ``ok=True`` — файл опознан (тип + реквизиты) и его можно копировать в
    папку документов. ``recognized`` — карточку контрагента можно создать
    автоматически (найден ИНН). Если ``ok`` и не ``recognized``, документ
    требует выбора контрагента пользователем (п. 3.3).

    Длительные разборы (более 500 мс) попадают в журнал (п. 10 этапа 5).
    """
    with applog.timed("Анализ файла импорта", 500, "import"):
        return _analyze_file(path)


def _analyze_file(path: str) -> dict:
    """Реализация разбора (см. ``analyze_file``)."""
    result = {
        "path": path,
        "name": os.path.basename(path),
        "kind": KIND_UNKNOWN,
        "number": "",
        "date": "",
        "party": None,
        "text": "",
        "error": "",
        "warning": "",
        "recognized": False,
        "ok": False,
    }
    try:
        check_file_limits(path)
        text = extract_text(path)
    except Exception as exc:
        result["error"] = str(exc)
        return result

    # Ограничение применяется до сохранения в реестр (п. 3.4).
    result["text"] = text[:TEXT_KEEP_LIMIT]
    if not text.strip():
        result["error"] = (
            "Не удалось извлечь текст: документ пуст или состоит только из картинок"
        )
        return result

    kind = detect_kind(text)
    if kind == KIND_UNKNOWN:  # пробуем по имени файла
        kind = detect_kind(os.path.splitext(os.path.basename(path))[0])
    party = extract_party(
        text,
        prefer_last_block=os.path.splitext(path)[1].lower() in (".xlsx", ".xlsm"),
    )
    result["kind"] = kind
    result["number"] = extract_number(text)
    result["date"] = extract_date(text)
    result["party"] = party
    result["ok"] = kind != KIND_UNKNOWN and party is not None
    result["recognized"] = bool(party and party.get("recognized"))

    if not result["ok"]:
        if kind == KIND_UNKNOWN and party is None:
            result["error"] = (
                "Не удалось опознать документ: не найден тип (договор, "
                "доп. соглашение, спецификация) и реквизиты контрагента"
            )
        elif kind == KIND_UNKNOWN:
            result["error"] = (
                "Не удалось определить тип документа: нет заголовка "
                "«Договор поставки», «Доп. соглашение» или «Спецификация»"
            )
        else:
            result["error"] = (
                "Не найдены реквизиты контрагента: нет ИНН и названия организации"
            )
    elif not result["recognized"]:
        missing = (party or {}).get("missing") or "ИНН"
        result["warning"] = (
            f"Не найден {missing}: карточку контрагента нужно подтвердить вручную"
        )
    return result
