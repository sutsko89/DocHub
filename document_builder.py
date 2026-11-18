"""Печатные формы: поля (F), плейсхолдеры (PH), замены и сборка .docx.

Чистый конвейер без зависимости от интерфейса (п. 9.1 плана) — общий
для приложения, smoke-теста и pytest-набора.
"""
from __future__ import annotations

from docx import Document
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from datetime import datetime
import calendar
import copy


def digits_only(text: str) -> str:
    """Оставляет только цифры — для вставки «замусоренного» номера из буфера."""
    return "".join(ch for ch in str(text) if ch.isdigit())


class F:
    """Ключи полей ввода. Имя поля меняется здесь — и больше нигде в коде."""

    NUMBER = "Номер договора"
    DAY = "Дата (число)"
    MONTH = "Месяц (словом)"
    FULL_NAME = "Полное название компании (Покупь)"
    SHORT_NAME = "Сокращенное название компании"
    DIRECTOR = "ФИО директора"
    BASIS = "Основание действий"
    IP_NAME = "ФИО ИП"
    INN = "ИНН"
    KPP = "КПП"
    OGRN = "ОГРН"
    OGRNIP = "ОГРНИП"
    LEGAL_ADDRESS = "Юридический адрес"
    FACT_ADDRESS = "Фактический адрес"
    EMAIL = "Email"
    PHONE = "Телефон"
    ACCOUNT = "Расчётный счёт"
    CORR_ACCOUNT = "Корреспондентский счёт"
    BANK = "Название банка"
    BIK = "БИК"
    LIMIT = "Лимит"
    TERM = "Отсрочка платежа"
    DELIVERY_TERM = "Срок поставки"
    PAYMENT_TERMS = "Условия оплаты"
    DELIVERY_BASIS = "Базис поставки"


class PH:
    """Плейсхолдеры в шаблонах .docx: код и шаблоны всегда синхронны."""

    NUMBER = "{{number}}"
    DATE = "{{date}}"
    MONTH = "{{month}}"                   # раньше в шаблонах был опечаткой {{mouth}}
    IP_NAME = "{{name_IP}}"
    IP_NAME2 = "{{name_IP2}}"
    IP_NAME3 = "{{name_IP3}}"
    OGRNIP = "{{OGRNIP}}"
    INN = "{{INN}}"
    LEGAL_ADDRESS = "{{legal_address}}"   # раньше {{Urid_adress}}
    FACT_ADDRESS = "{{fact_address}}"     # раньше {{fact_adress}}
    EMAIL = "{{E-mail}}"
    PHONE = "{{Telephone}}"
    ACCOUNT = "{{r/schet}}"
    BANK = "{{bank}}"
    CORR_ACCOUNT = "{{k/schet}}"
    BIK = "{{BIK}}"
    FULL_NAME = "{{long_name_company}}"
    SHORT_NAME = "{{short_name_company}}"
    DIRECTOR = "{{direct_name}}"
    DIRECTOR_INITIALS = "{{direct_name2}}"
    BASIS = "{{ustava}}"
    OGRN = "{{OGRN}}"
    KPP = "{{KPP}}"
    LIMIT = "{{limit}}"
    LIMIT_WORDS = "{{limit_in_words}}"
    TERM = "{{term}}"
    TERM_WORDS = "{{term_in_words}}"
    AGREEMENT_NUMBER = "{{agreement_number}}"

    # --- Спецификация (spec_template.docx) -------------------------------- #
    APPENDIX_NUMBER = "{{appendix_number}}"
    SPEC_NUMBER = "{{spec_number}}"
    SPEC_DAY = "{{spec_day}}"
    SPEC_MONTH = "{{spec_month}}"
    PRICES_VALID_UNTIL = "{{prices_valid_until}}"
    ITEM_NO = "{{item_no}}"
    ITEM_NAME = "{{item_name}}"
    ITEM_QTY = "{{item_qty}}"
    ITEM_UNIT = "{{item_unit}}"
    ITEM_PRICE = "{{item_price}}"
    ITEM_SUM = "{{item_sum}}"
    ITEM_NOTE = "{{item_note}}"
    TOTAL = "{{total}}"
    TOTAL_IN_WORDS = "{{total_in_words}}"
    VAT = "{{vat}}"
    VAT_RATE = "{{vat_rate}}"
    DELIVERY_TERM = "{{delivery_term}}"
    PAYMENT_TERMS = "{{payment_terms}}"
    DELIVERY_BASIS = "{{delivery_basis}}"
SPEC_APPENDIX_NUMBER = "1"   # номер приложения к договору, как в образце
SPEC_VAT_RATE = 22           # ставка НДС «в том числе», %


def spec_vat_amount(total, rate: int = SPEC_VAT_RATE) -> Decimal:
    """Сумма НДС «в том числе» из общей суммы: total * rate / (100 + rate)."""
    total_dec = total if isinstance(total, Decimal) else parse_decimal(total)
    return (total_dec * rate / (100 + rate)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )


def spec_vat_amount_from_items(items, rate: int = SPEC_VAT_RATE) -> Decimal:
    """Сумма НДС «в том числе» так же, как в 1С: НДС считается по каждой
    позиции (сумма позиции * rate / (100 + rate), округление до копеек),
    после чего НДС по позициям складывается в итог."""
    vat = Decimal("0")
    for item in items or []:
        qty = item.get("qty") if isinstance(item.get("qty"), Decimal) else parse_decimal(item.get("qty"))
        price = item.get("price") if isinstance(item.get("price"), Decimal) else parse_decimal(item.get("price"))
        line_total = (qty * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        vat += (line_total * rate / (100 + rate)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    return vat


def format_director_name(full_name):
    """Преобразует ФИО из формата 'Фамилия Имя Отчество' в формат 'И.О. Фамилия'"""
    parts = full_name.strip().split()
    if len(parts) < 2:
        return full_name
    surname = parts[0]
    initials = ""
    for i in range(1, min(3, len(parts))):
        if parts[i]:
            initials += parts[i][0] + "."
    return f"{initials} {surname}"


def format_name_ip2(full_name):
    """Преобразует ФИО из формата 'Фамилия Имя Отчество' в формат 'Фамилия И.О.'"""
    parts = full_name.strip().split()
    if len(parts) < 2:
        return full_name
    surname = parts[0]
    initials = ""
    for i in range(1, min(3, len(parts))):
        if parts[i]:
            initials += parts[i][0] + "."
    return f"{surname} {initials}"


def format_date(day):
    """Форматирует число дня: добавляет ведущий ноль для чисел меньше 10"""
    day_str = str(day).strip()
    if day_str.isdigit():
        day_num = int(day_str)
        return f"{day_num:02d}"
    return day_str
MONTHS_GENITIVE = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]


def parse_decimal(value) -> Decimal:
    """Разбирает число: цифры, пробелы, запятая или точка. Ошибка — 0."""
    cleaned = str(value or "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    if not cleaned:
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return Decimal("0")


def format_amount(value) -> str:
    """Денежная сумма: разряды по три, копейки только если есть (2052,59)."""
    dec = (value if isinstance(value, Decimal) else parse_decimal(value))
    dec = dec.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    whole = int(dec)
    frac = int((dec - whole) * 100)
    if frac == 0:
        return format_number_grouped(str(whole))
    return f"{format_number_grouped(str(whole))},{frac:02d}"


def format_quantity(value) -> str:
    """Количество: целое — без дробной части, дробное — с запятой (2,5)."""
    dec = value if isinstance(value, Decimal) else parse_decimal(value)
    if dec == dec.to_integral_value():
        return format_number_grouped(str(int(dec)))
    text = f"{dec.quantize(Decimal('0.001'), rounding=ROUND_HALF_UP):f}"
    text = text.rstrip("0").rstrip(".")
    return text.replace(".", ",")


def format_phone(digits: str) -> str:
    """Собирает маску +7 (999) 123-45-67 из цифр (без «7» — добавит сам)."""
    digits = digits[:11]
    if not digits:
        return ""
    if digits.startswith("8") and len(digits) == 11:
        digits = "7" + digits[1:]
    if not digits.startswith("7"):
        digits = ("7" + digits)[:11]
    out = "+7"
    if len(digits) > 1:
        out += " (" + digits[1:4]
    if len(digits) >= 5:
        out += ") " + digits[4:7]
    if len(digits) >= 8:
        out += "-" + digits[7:9]
    if len(digits) >= 10:
        out += "-" + digits[9:11]
    return out


def format_number_grouped(value: str) -> str:
    """Группирует цифры по три разряда: 1500000 -> 1 500 000."""
    digits = digits_only(value)
    if not digits:
        return ""
    groups = []
    while len(digits) > 3:
        groups.insert(0, digits[-3:])
        digits = digits[:-3]
    groups.insert(0, digits)
    return " ".join(groups)
_UNITS = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_UNITS_FEM = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_TEENS = [
    "десять", "одиннадцать", "двенадцать", "тринадцать",
    "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать",
    "восемнадцать", "девятнадцать",
]
_TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят",
         "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
_HUNDREDS = ["", "сто", "двести", "триста", "четыреста",
             "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"]


def _below_thousand(n: int, feminine: bool = False) -> str:
    units = _UNITS_FEM if feminine else _UNITS
    parts = []
    hundreds, rest = divmod(n, 100)
    if hundreds:
        parts.append(_HUNDREDS[hundreds])
    if 10 <= rest < 20:
        parts.append(_TEENS[rest - 10])
    else:
        tens, unit = divmod(rest, 10)
        if tens:
            parts.append(_TENS[tens])
        if unit:
            parts.append(units[unit])
    return " ".join(parts)


def _plural(n: int, forms: tuple) -> str:
    n %= 100
    if 11 <= n <= 14:
        return forms[2]
    n %= 10
    if n == 1:
        return forms[0]
    if 2 <= n <= 4:
        return forms[1]
    return forms[2]
_SCALES = (
    (("тысяча", "тысячи", "тысяч"), True),
    (("миллион", "миллиона", "миллионов"), False),
    (("миллиард", "миллиарда", "миллиардов"), False),
    (("триллион", "триллиона", "триллионов"), False),
)


def number_to_words_ru(n: int) -> str:
    """Число прописью (до триллионов включительно)."""
    if n < 0:
        return "минус " + number_to_words_ru(-n)
    if n == 0:
        return "ноль"

    groups = []
    while n:
        groups.append(n % 1000)
        n //= 1000

    parts = []
    for index in range(len(groups) - 1, -1, -1):
        value = groups[index]
        if value == 0:
            continue
        if index == 0:
            parts.append(_below_thousand(value))
        else:
            forms, feminine = _SCALES[index - 1]
            parts.append(_below_thousand(value, feminine=feminine))
            parts.append(_plural(value, forms))
    return " ".join(parts)


def amount_to_words_ru(n: int) -> str:
    """Сумма прописью: «пятьсот тысяч рублей»."""
    return f"{number_to_words_ru(n)} {_plural(n, ('рубль', 'рубля', 'рублей'))}"


def amount_in_words_ru(value) -> str:
    """Сумма прописью с копейками: «Одиннадцать тысяч шестьсот пятьдесят
    семь рублей 95 копеек» (как в образце спецификации)."""
    dec = value if isinstance(value, Decimal) else parse_decimal(value)
    dec = dec.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    whole = int(dec)
    kopecks = int(round(float(dec - whole) * 100))
    rubles = (
        f"{number_to_words_ru(whole).capitalize()} "
        f"{_plural(whole, ('рубль', 'рубля', 'рублей'))}"
    )
    return (
        f"{rubles} {kopecks:02d} "
        f"{_plural(kopecks, ('копейка', 'копейки', 'копеек'))}"
    )


def replace_in_paragraph(paragraph, replacements: dict):
    """Заменяет плейсхолдеры за один проход по параграфу (п. 5.6).

    Тексты runs зеркалируются в обычные строки: поиск и замены идут по
    обычному тексту, а в runs результат записывается один раз в конце.
    Раньше после каждой одиночной замены текст параграфа полностью
    пересобирался из runs (каждое обращение к run.text — обход lxml-дерева),
    что давало O(n²) на больших шаблонах.
    """
    runs = paragraph.runs
    if not runs:
        return

    texts = [r.text or "" for r in runs]
    full_text = "".join(texts)
    changed: set[int] = set()

    for raw_key, raw_value in replacements.items():
        key = str(raw_key)
        value = "" if raw_value is None else str(raw_value)
        if not key:
            continue

        search_from = 0
        while True:
            start = full_text.find(key, search_from)
            if start < 0:
                break
            end = start + len(key)

            # Границы runs — то же правило, что раньше: run выбирается,
            # если позиция попадает в его диапазон (пустые runs пропускаются).
            cur = 0
            start_run = None
            start_offset = 0
            end_run = None
            end_offset = 0
            for i, text in enumerate(texts):
                next_cur = cur + len(text)
                if start_run is None and start < next_cur:
                    start_run = i
                    start_offset = start - cur
                if end_run is None and end <= next_cur:
                    end_run = i
                    end_offset = end - cur
                    break
                cur = next_cur

            if start_run is None or end_run is None:
                break

            prefix = texts[start_run][:start_offset]
            suffix = texts[end_run][end_offset:]

            texts[start_run] = prefix + value + suffix
            if end_run != start_run:
                for j in range(start_run + 1, end_run + 1):
                    texts[j] = ""
            changed.add(start_run)
            if end_run != start_run:
                changed.update(range(start_run + 1, end_run + 1))

            full_text = full_text[:start] + value + full_text[end:]
            # Продолжаем поиск ПОСЛЕ вставленного значения: подстановка не
            # обрабатывает собственный результат (защита от зацикливания,
            # если в подстановке снова встречается этот же плейсхолдер).
            search_from = start + len(value)

    for index in sorted(changed):
        runs[index].text = texts[index]


def replace_in_doc(doc, replacements: dict):
    for p in doc.paragraphs:
        replace_in_paragraph(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, replacements)

    for section in doc.sections:
        header = section.header
        footer = section.footer

        for p in header.paragraphs:
            replace_in_paragraph(p, replacements)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, replacements)

        for p in footer.paragraphs:
            replace_in_paragraph(p, replacements)
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, replacements)


def spec_items_table(doc) -> "Table | None":
    """Таблица позиций спецификации в шаблоне (ищется по заголовку)."""
    for table in doc.tables:
        if not table.rows:
            continue
        header = " ".join(cell.text for cell in table.rows[0].cells)
        if "№ п/п" in header and "Наименование" in header:
            return table
    return None


def fill_spec_items(doc, items: list[dict]):
    """Заполняет таблицу позиций спецификации из шаблона.

    Строка-шаблон с плейсхолдерами {{item_*}} клонируется для каждой позиции,
    после чего исходная строка-шаблон удаляется. Форматирование ячеек
    сохраняется — текст заменяется внутри существующих runs.
    Цена — с копейками при наличии дробной части, кол-во может быть дробным.
    """
    table = spec_items_table(doc)
    if table is None or len(table.rows) < 2:
        return
    tpl_tr = table.rows[1]._tr
    for index, item in enumerate(items, start=1):
        qty = item.get("qty") if isinstance(item.get("qty"), Decimal) else parse_decimal(item.get("qty"))
        price = item.get("price") if isinstance(item.get("price"), Decimal) else parse_decimal(item.get("price"))
        amount = (qty * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        values = {
            PH.ITEM_NO: str(index),
            PH.ITEM_NAME: str(item.get("name", "")),
            PH.ITEM_QTY: format_quantity(qty),
            PH.ITEM_UNIT: str(item.get("unit", "")),
            PH.ITEM_PRICE: format_amount(price),
            PH.ITEM_SUM: format_amount(amount),
            PH.ITEM_NOTE: str(item.get("note", "") or ""),
        }
        new_tr = copy.deepcopy(tpl_tr)
        tpl_tr.addprevious(new_tr)
        for tc in new_tr.tc_lst:
            for paragraph in tc.xpath(".//w:p"):
                for run in paragraph.xpath(".//w:r"):
                    for token, value in values.items():
                        if token in run.text:
                            run.text = run.text.replace(token, value)
    tpl_tr.getparent().remove(tpl_tr)


def prices_valid_until_default() -> str:
    """Дата окончания действия цен — последний день текущего месяца."""
    today = datetime.now()
    last_day = calendar.monthrange(today.year, today.month)[1]
    return f"{last_day:02d}.{today.month:02d}.{today.year} г."


def mark_spec_final_page(doc):
    """Начинает итоговый блок спецификации (после таблицы позиций) с новой
    страницы, чтобы он всегда размещался на отдельном последнем листе."""
    for paragraph in doc.paragraphs:
        if PH.TOTAL_IN_WORDS in paragraph.text:
            paragraph.paragraph_format.page_break_before = True
            return


def docx_to_text(doc) -> str:
    """Извлекает текст из документа: абзацы, затем строки таблиц."""
    lines = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            lines.append("\t".join(cells))
    return "\n".join(lines).strip()


def docx_from_text(text: str) -> "Document":
    """Собирает документ Word из редактированного текста (по строке на абзац)."""
    doc = Document()
    for line in (text or "").splitlines():
        doc.add_paragraph(line)
    return doc


def build_replacements(data: dict, is_ip: bool) -> dict:
    """Словарь замен плейсхолдеров для шаблонов договора и доп. соглашения."""
    formatted_date = format_date(data.get(F.DAY, ""))

    if is_ip:
        name_ip = data.get(F.IP_NAME, "")
        return {
            PH.NUMBER: data.get(F.NUMBER, ""),
            PH.DATE: formatted_date,
            PH.MONTH: data.get(F.MONTH, ""),
            PH.IP_NAME: name_ip,
            PH.IP_NAME2: format_name_ip2(name_ip),
            PH.IP_NAME3: format_director_name(name_ip),
            PH.OGRNIP: data.get(F.OGRNIP, ""),
            PH.INN: data.get(F.INN, ""),
            PH.LEGAL_ADDRESS: data.get(F.LEGAL_ADDRESS, ""),
            PH.FACT_ADDRESS: data.get(F.FACT_ADDRESS, ""),
            PH.EMAIL: data.get(F.EMAIL, ""),
            PH.PHONE: data.get(F.PHONE, ""),
            PH.ACCOUNT: data.get(F.ACCOUNT, ""),
            PH.BANK: data.get(F.BANK, ""),
            PH.CORR_ACCOUNT: data.get(F.CORR_ACCOUNT, ""),
            PH.BIK: data.get(F.BIK, ""),
        }

    basis = data.get(F.BASIS, "")
    if basis == "Устав":
        basis_text = "устава"
    elif basis == "Доверенность":
        basis_text = "доверенности"
    else:
        basis_text = ""

    return {
        PH.NUMBER: data.get(F.NUMBER, ""),
        PH.DATE: formatted_date,
        PH.MONTH: data.get(F.MONTH, ""),
        PH.FULL_NAME: data.get(F.FULL_NAME, ""),
        PH.SHORT_NAME: data.get(F.SHORT_NAME, ""),
        PH.DIRECTOR: data.get(F.DIRECTOR, ""),
        PH.DIRECTOR_INITIALS: format_director_name(data.get(F.DIRECTOR, "")),
        PH.BASIS: basis_text,
        PH.OGRN: data.get(F.OGRN, ""),
        PH.INN: data.get(F.INN, ""),
        PH.KPP: data.get(F.KPP, ""),
        PH.LEGAL_ADDRESS: data.get(F.LEGAL_ADDRESS, ""),
        PH.FACT_ADDRESS: data.get(F.FACT_ADDRESS, ""),
        PH.EMAIL: data.get(F.EMAIL, ""),
        PH.PHONE: data.get(F.PHONE, ""),
        PH.ACCOUNT: data.get(F.ACCOUNT, ""),
        PH.BANK: data.get(F.BANK, ""),
        PH.CORR_ACCOUNT: data.get(F.CORR_ACCOUNT, ""),
        PH.BIK: data.get(F.BIK, ""),
    }
