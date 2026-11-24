"""GUI-free ядро реестра (этап M5) — модель строк и чистые функции.

Единый источник логики реестра для обеих сборок (аналог spec_table_core
и party_form_core, этапы M3/M4): RowSpec/DiffPlan и план дифа
(plan_diff/plan_display_count), константы колонок, сортировки, фильтра
типов и точек типов, подписи панели действий, чистые построители значений
строк контрагента/документа.

Tk ``registry_tree`` и ``registry_controller`` делегируют (имена
сохранены, тесты Tk без правок); Qt ``registry_page_qt`` строит
QAbstractTableModel из тех же RowSpec — дрейф «данные ↔ отображение»
между сборками исключён по построению. Модуль не импортирует tkinter
и PySide6 — тестируется без GUI.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

import agreements
import storage
from document_builder import F


#: Мягкий лимит отображения (R-05b): порция «Показать ещё». Значение из
#: прототип-замера (tools/bench_tree_prototype.py): популяция 10k — 86 мс
#: при пороге раздела Е 300 мс; 50k — 470 мс, поэтому объёмы до 10k
#: показываются целиком, большие — порциями по 10k с «Показать ещё».
#: Запрет раздела З: лимит 10 000 — значение прототип-замера, не
#: константа «на глаз»; не менять без нового замера.
DISPLAY_SOFT_LIMIT = 10_000

#: iid служебной строки «Показать ещё».
SHOW_MORE_IID = "show-more"

#: Префиксы iid по типу записи (party/doc ids не пересекаются).
IID_PARTY = "p:"
IID_DOC = "d:"
IID_TRASH_PARTY = "tp:"
IID_TRASH_DOC = "td:"
IID_UNKNOWN = "u:"
IID_SEC_UNKNOWN = "sec:unknown"
IID_SEC_TRASH = "sec:trash"

#: Колонки списка контрагентов (U1): ключ, заголовок, ширина, якорь, тянуть.
PARTY_COLUMNS = (
    ("name", "Контрагент", 340, "w", True),
    ("kind", "Тип", 92, "w", False),
    ("inn", "ИНН", 128, "w", False),
    ("id", "ID", 104, "w", False),
    ("address", "Адрес", 360, "w", True),
)

#: Колонки вида «Документы контрагента».
DOC_COLUMNS = (
    ("doc", "Документ", 330, "w", True),
    ("date", "Дата", 116, "w", False),
    ("expiry", "Срок", 350, "w", True),
)

#: Сортировка списка контрагентов (U2): подпись меню -> значение settings.
SORT_VALUE_BY_LABEL = {
    "По добавлению": "added",
    "По алфавиту": "name",
    "По ИНН": "inn",
}
SORT_LABEL_BY_VALUE = {value: label for label, value in SORT_VALUE_BY_LABEL.items()}

#: Фильтр типов в виде документов контрагента (UX-REG).
DOCS_TYPE_VALUE_BY_LABEL = {
    "Все типы": "all",
    "Договоры": storage.KIND_CONTRACT,
    "Доп. соглашения": storage.KIND_ADDITIONAL,
    "Спецификации": storage.KIND_SPEC,
}
DOCS_TYPE_LABEL_BY_VALUE = {value: label for label, value in DOCS_TYPE_VALUE_BY_LABEL.items()}

#: Цвета точек типов (V4-пилот): точка в нулевой колонке.
DOT_IP = "#28C840"
DOT_OOO = "#007AFF"
DOT_CONTRACT = "#007AFF"
DOT_ADDITIONAL = "#FEBC2E"
DOT_SPEC = "#28C840"
DOT_UNKNOWN = "#96969C"

#: Кнопки панели действий по типу выделенной записи (V3).
ACTION_LABELS = {
    "party": ("Редактировать", "Удалить"),
    "doc": ("Открыть Word", "Удалить"),
    "trash_party": ("Восстановить", "Удалить навсегда"),
    "trash_doc": ("Восстановить", "Удалить навсегда"),
    "unknown": ("Показать в папке", "Удалить"),
}


def action_labels_for(kind: str | None) -> tuple[str, str] | None:
    """Подписи кнопок панели действий для типа записи (чистая функция)."""
    return ACTION_LABELS.get(kind or "")


@dataclass
class RowSpec:
    """Строка модели реестра: iid, значения колонок, теги, точка типа."""

    iid: str
    values: tuple[str, ...]
    tags: tuple[str, ...] = ()
    dot: str | None = None  # цвет точки типа (V4); None — точки нет
    data: dict = field(default_factory=dict, repr=False)


@dataclass
class DiffPlan:
    """План диф-обновления (чистая функция plan_diff)."""

    to_delete: list[str]
    ops: list[tuple[int, str, str]]  # (позиция, iid, "insert"|"move")
    updates: list[RowSpec]


def plan_display_count(total: int, shown: int) -> tuple[int, bool]:
    """Сколько строк показывать и есть ли ещё (R-05b).

    ``shown`` — сколько строк уже разрешено показывать (копится
    кнопкой «Показать ещё»). Возвращает (количество строк, есть_ли_скрытые).
    Чистая функция — тестируется без Tk.
    """
    shown = max(shown, DISPLAY_SOFT_LIMIT)
    shown = min(shown, total) if total else 0
    return shown, total > shown


def plan_diff(current: Iterable[str], desired: list[RowSpec]) -> DiffPlan:
    """Диф модели строк без Tk: (удалить, операции, обновить).

    ``current`` — iid видимых строк в порядке отображения. Алгоритм:
    префикс, совпадающий с желаемой моделью, не трогается вовсе (критерий
    5.1 — «выбор/изменение не пересоздаёт строки»); правится только
    суффикс после первой расходимости — вставки и перемещения применяются
    в порядке возрастания позиции, поэтому каждая операция видит все
    строки до неё уже стоящими правильно.
    """
    current = list(current)
    # ВАЖНО: множество желаемых iid строится ОДИН раз. Внутри comprehension
    # оно пересобиралось бы на каждой итерации — O(n²) (на 10k это 5,8 с
    # вместо 2 мс, найдено профилированием прототипа).
    desired_iids = {row.iid for row in desired}
    to_delete = [iid for iid in current if iid not in desired_iids]
    deleted = set(to_delete)
    sim = [iid for iid in current if iid not in deleted]
    pos_of = {iid: idx for idx, iid in enumerate(sim)}

    ops: list[tuple[int, str, str]] = []  # (позиция, iid, "insert"|"move")
    k = 0
    while (k < len(sim) and k < len(desired)
           and sim[k] == desired[k].iid):
        k += 1
    for pos in range(k, len(desired)):
        iid = desired[pos].iid
        if iid in pos_of:
            if pos_of[iid] != pos:
                # Сдвиг в симуляции: как у tree.move — снять и вставить.
                sim.pop(pos_of[iid])
                sim.insert(pos, iid)
                lo, hi = min(pos_of[iid], pos), max(pos_of[iid], pos)
                for shift in range(lo, hi + 1):
                    pos_of[sim[shift]] = shift
                ops.append((pos, iid, "move"))
        else:
            sim.insert(pos, iid)
            for shift in range(pos, len(sim)):
                pos_of[sim[shift]] = shift
            ops.append((pos, iid, "insert"))
    return DiffPlan(to_delete=to_delete, ops=ops, updates=list(desired))


def party_row_values(rec: dict) -> tuple[str, str, str, str, str]:
    """Значения колонок строки контрагента (U1: name, тип, ИНН, ID, адрес)."""
    fields = rec.get("fields") or {}
    kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
    inn = str(fields.get(F.INN, "")).strip() or "—"
    address = str(fields.get(F.LEGAL_ADDRESS, "")).strip()
    if len(address) > 48:
        address = address[:47].rstrip() + "…"
    return (
        storage.display_name(rec), kind, inn, str(rec.get("id")), address,
    )


def party_row_dot(rec: dict) -> str:
    """Цвет точки типа контрагента (V4-пилот)."""
    return DOT_IP if rec.get("type") == "ip" else DOT_OOO


def doc_row_values(rec: dict, until: "str | None" = None) -> tuple[str, str, str]:
    """Значения колонок строки документа: документ, дата, срок (U1).

    ``until`` — срок из лёгкой карты хранилища (LAZY-1); без него срок
    берётся из тела в памяти записи (совместимость со старыми вызовами).
    """
    kind_label = storage.KIND_LABELS.get(rec.get("kind"), rec.get("kind", ""))
    number = rec.get("number") or "б/н"
    date_text = rec.get("date") or ""
    suffix = agreements.row_suffix(rec, until=until)
    if suffix.startswith("   ·   "):
        suffix = suffix[len("   ·   "):]
    return (f"{kind_label} № {number}", date_text or "—", suffix)


def doc_row_tags(rec: dict, until: "str | None" = None) -> tuple[str, ...]:
    """Теги строки документа по состоянию срока (2.18.0): красный/амберный."""
    if rec.get("kind") == storage.KIND_ADDITIONAL:
        state = agreements.collect_state(rec, until=until)
        if state == "expired":
            return ("expired",)
        if state == "soon":
            return ("soon",)
    return ()


def party_subtitle(rec: dict) -> str:
    """Подстрока контрагента в виде документов: тип, ИНН, ID (2.19.0)."""
    kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
    inn = str(rec.get("fields", {}).get(F.INN, "")).strip() or "—"
    return f"{kind}   ·   ИНН {inn}   ·   ID {rec.get('id')}"
