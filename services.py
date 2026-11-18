"""Сервисный слой (п. 9.2): валидация, кэш шаблонов и сохранение документов.

Функции работают над DataStore и словарями данных и не зависят от интерфейса:
main.py собирает данные форм, показывает диалоги и уведомления, а сборка
.docx из шаблона и транзакция «файл + запись реестра» живут здесь. Это же
API используют headless-тесты (pytest) без запуска интерфейса.
"""
from __future__ import annotations

import copy
import hashlib
import os
import re
import sys
import zipfile
from decimal import Decimal

import applog
import storage
from docx import Document
from document_builder import (
    F, PH, SPEC_APPENDIX_NUMBER, SPEC_VAT_RATE,
    amount_in_words_ru, amount_to_words_ru, build_replacements, digits_only,
    docx_from_text, fill_spec_items, format_amount, format_date,
    format_name_ip2, mark_spec_final_page, number_to_words_ru, parse_decimal,
    prices_valid_until_default, replace_in_doc, spec_vat_amount,
    spec_vat_amount_from_items,
)

MAX_PATH_LENGTH = 240
MAX_FILENAME_LENGTH = 80
def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
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


# ------------------------------------------------------------------ #
#  Сборка и запись документов (перенос из ContractApp, п. 9.2)         #
# ------------------------------------------------------------------ #

def fit_document_path(directory: str, base: str, ext: str = ".docx") -> str:
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


def rec_is_ip(rec: dict, store=None) -> bool:
    """Признак ИП у записи реестра (учитывая старые записи без is_ip)."""
    extra = rec.get("extra") or {}
    if "is_ip" in extra:
        return bool(extra.get("is_ip"))
    party = store.get_counterparty(rec.get("party_id")) if store is not None else None
    return bool(party and party.get("type") == "ip")


def render_contract(data: dict, is_ip: bool) -> "Document":
    """Договор из шаблона: печатная форма сохраняется без искажений."""
    template_path = resource_path("dogovor_template_ip.docx" if is_ip else "dogovor_template.docx")
    doc = load_template_cached(template_path)
    replace_in_doc(doc, build_replacements(data, is_ip))
    return doc


def render_additional(data: dict, is_ip: bool, number: str,
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


def render_spec(rec: dict, store, spec_settings: dict | None = None) -> "Document":
    """Спецификация из шаблона (книжная ориентация).

    У ООО и ИП разные шаблоны: у ИП нет КПП, а вместо ОГРН — ОГРНИП.
    Общий spec_template.docx написан «в терминах ООО», поэтому для ИП
    используется отдельная копия spec_template_ip.docx (п. 1.1):
    строка «ОГРН {{OGRN}}» заменена на «ОГРНИП {{OGRNIP}}»,
    «ИНН {{INN}} / КПП {{KPP}}» — на «ИНН {{INN}}». Без этого в печатной
    форме ИП оставались пустые подписи «ОГРН» и «/ КПП».

    Ставка НДС и номер приложения берутся из spec_settings (п. 9.7);
    при отсутствии — дефолты SPEC_VAT_RATE / SPEC_APPENDIX_NUMBER.
    """
    spec_settings = spec_settings or {}
    try:
        vat_rate = int(spec_settings.get("spec_vat_rate") or SPEC_VAT_RATE)
    except (TypeError, ValueError):
        vat_rate = SPEC_VAT_RATE
    appendix_number = str(
        spec_settings.get("spec_appendix_number") or SPEC_APPENDIX_NUMBER
    )
    is_ip = rec_is_ip(rec, store)
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
        store.get_document(extra.get("contract_id"))
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
        PH.APPENDIX_NUMBER: appendix_number,
        PH.SPEC_NUMBER: rec.get("number") or "б/н",
        PH.SPEC_DAY: str(extra.get("spec_day", "")),
        PH.SPEC_MONTH: str(extra.get("spec_month", "")),
        PH.PRICES_VALID_UNTIL: prices_valid_until_default(),
        PH.TOTAL: format_amount(total),
        PH.TOTAL_IN_WORDS: amount_in_words_ru(total) if total else "",
        PH.VAT: format_amount(
            spec_vat_amount_from_items(items, vat_rate)
            if items else spec_vat_amount(total, vat_rate)
        ),
        PH.VAT_RATE: f"{vat_rate}%",
        PH.DELIVERY_TERM: str(extra.get("delivery_term", "")),
        PH.PAYMENT_TERMS: str(extra.get("payment_terms", "")),
        PH.DELIVERY_BASIS: str(extra.get("delivery_basis", "")),
    })
    mark_spec_final_page(doc)
    replace_in_doc(doc, replacements)
    fill_spec_items(doc, items)
    return doc


def build_record_docx(rec: dict, store) -> "Document":
    """Собирает .docx записи реестра по её типу (или из сохранённого текста)."""
    kind = rec.get("kind")
    extra = rec.get("extra") or {}
    if kind == storage.KIND_CONTRACT:
        return render_contract(rec.get("data") or {}, rec_is_ip(rec, store))
    if kind == storage.KIND_ADDITIONAL:
        return render_additional(
            rec.get("data") or {}, rec_is_ip(rec, store),
            rec.get("number") or "", extra.get("limit", ""), extra.get("term", ""),
        )
    if kind == storage.KIND_SPEC:
        return render_spec(rec, store)
    return docx_from_text(rec.get("text", ""))


def write_docx(doc, rec: dict, store, docs_dir: str) -> str:
    """Пишет .docx в папку контрагента атомарно и возвращает итоговый путь."""
    kind_label = storage.KIND_LABELS.get(rec.get("kind"), "Документ")
    party_name = store.party_name(rec.get("party_id"))
    number = rec.get("number") or "б/н"
    base = safe_filename(f"{kind_label} №{number} {party_name}", "Документ")
    path = fit_document_path(docs_dir, base, ".docx")
    counter = 2
    while os.path.exists(path):
        path = fit_document_path(docs_dir, f"{base} ({counter})", ".docx")
        counter += 1
    with applog.timed("Запись .docx", 500, "documents"):
        save_docx_atomic(doc, path)
    applog.get_logger("documents").info(
        "Документ записан: тип=%s, id=%s", rec.get("kind"), rec.get("id"),
    )
    return path


def save_document(store, draft: dict, docs_dir: str) -> dict:
    """Транзакция сохранения документа: .docx, затем запись в реестр.

    Сначала пишется файл, затем регистрируется запись (п. 1.6); при сбое
    записи реестра «осиротевший» файл удаляется и запись не добавляется.
    Возвращает {'ok', 'path', 'record', 'error'} — интерфейс только показывает
    результат и открывает файл.
    """
    doc = build_record_docx(draft, store)
    path = write_docx(doc, draft, store, docs_dir)
    rec_doc = store.add_document(
        draft["kind"], draft["party_id"], draft["number"], draft.get("date", ""),
        data=draft.get("data") or {}, extra=draft.get("extra") or {},
        text=draft.get("text", ""),
    )
    if store.last_error:
        try:
            store.documents.remove(rec_doc)
        except ValueError:
            pass
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except OSError:
            pass
        return {"ok": False, "path": path, "record": None,
                "error": store.last_error or "Ошибка записи реестра."}
    return {"ok": True, "path": path, "record": rec_doc, "error": None}
