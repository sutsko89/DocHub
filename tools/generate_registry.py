#!/usr/bin/env python3
"""Генератор синтетических реестров для baseline-замеров (DIAG-1, раздел Е).

Создаёт детерминированный ``database.db`` через штатный API DataStore
(вне прод-кода: прод-код не меняется). Соотношение типовое — каждый
контрагент с несколькими документами (договор, доп. соглашение,
спецификации); тексты документов — синтетические по 2–5 КБ (реалистичный
вес колонки ``text``). Одинаковые аргументы дают одинаковые данные на
всех прогонах «до» и «после».

Использование:
    python tools/generate_registry.py --out <каталог> --records 10000

В <каталоге> появляется database.db (WAL). Запись реестра — одна
транзакция в конце (как при штатном save()).
"""

from __future__ import annotations

import argparse
import random
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import storage  # noqa: E402
from document_builder import F  # noqa: E402

SEED = 20261121  # детерминированность: одинаковые данные на всех прогонах

ITEM_NAMES = (
    "Болт М8x40 оцинк.", "Гайка М8", "Шайба М8", "Кабель ВВГнг 3x2.5",
    "Профиль ПН-2 50x40", "Лист ОЦ 0.5 мм", "Труба профильная 40x20",
    "Анкер клиновой 10x90", "Саморез 4.5x35", "Уголок 40x40x3",
)
ITEM_UNITS = ("шт.", "м", "кг", "уп.")
FILLER = (
    "поставки", "товара", "выполнения работ", "услуги", "монтажа",
    "сервисного обслуживания", "материалов", "оборудования",
)


def _synthetic_text(rng: random.Random, size_kb: float) -> str:
    """Синтетический текст заданного веса (2–5 КБ — реалистичный вес text)."""
    target = int(size_kb * 1024)
    parts: list[str] = []
    total = 0
    while total < target:
        chunk = " ".join(rng.choice(FILLER) for _ in range(10))
        line = f"Условие {chunk} — согласовано сторонами."
        parts.append(line)
        total += len(line) + 1
    return "\n".join(parts)


def _party_fields(index: int, is_ip: bool) -> dict:
    inn = ("12%010d" % index) if is_ip else ("10%09d" % index)
    name = f"ИП Смирнов-{index} И.И." if is_ip else f"ООО «Ромашка-{index}»"
    fields = {
        F.INN: inn,
        F.LEGAL_ADDRESS: f"г. Хабаровск, ул. Примерная, д. {index % 150 + 1}, оф. {index % 40 + 1}",
        F.ACCOUNT: "40702810%012d" % index,
        F.CORR_ACCOUNT: "30101810600000000851",
        F.BANK: "Дальневосточный банк ПАО Сбербанк",
        F.BIK: "040813608",
        F.EMAIL: f"client{index}@example.test",
        F.PHONE: f"+7 (4212) {index % 100:02d}-{(index // 100) % 100:02d}-00",
    }
    if is_ip:
        fields[F.IP_NAME] = name
        fields[F.OGRNIP] = "3%014d" % index
        fields[F.FACT_ADDRESS] = fields[F.LEGAL_ADDRESS]
    else:
        fields[F.FULL_NAME] = f"Общество с ограниченной ответственностью «Ромашка-{index}»"
        fields[F.SHORT_NAME] = name
        fields[F.DIRECTOR] = f"Смирнов-{index} И.И."
        fields[F.KPP] = "254001001"
        fields[F.OGRN] = "1%012d" % index
        fields[F.FACT_ADDRESS] = fields[F.LEGAL_ADDRESS]
        fields[F.BASIS] = "Устав"
    return fields


def _items(rng: random.Random, count: int) -> list[dict]:
    items = []
    for n in range(1, count + 1):
        qty = rng.choice(("1", "2", "5", "10", "12.5", "100"))
        price = rng.choice(("120", "340.50", "1 250", "8 900.99", "15 000"))
        items.append({
            "name": f"{rng.choice(ITEM_NAMES)} (партия {n})",
            "qty": qty,
            "unit": rng.choice(ITEM_UNITS),
            "price": price,
        })
    return items


def generate(out_dir: str, records: int, seed: int = SEED) -> dict:
    """Строит реестр на ``records`` записей (контрагенты + документы)."""
    rng = random.Random(seed)
    target = Path(out_dir).resolve()
    target.mkdir(parents=True, exist_ok=True)
    db_file = target / storage.DB_FILENAME

    # DataStore пишет строго в storage.db_path() — в скрипте-генераторе
    # (вне прод-кода) путь подменяется на целевой каталог.
    storage.db_path = lambda: str(db_file)  # type: ignore[assignment]

    started = time.perf_counter()
    store = storage.DataStore()
    # ~1 контрагент на 5 записей: договор + доп. соглашение + ~3 спецификации.
    parties = max(1, records // 5)
    documents_planned = max(0, records - parties)

    docs_created = 0
    for index in range(1, parties + 1):
        is_ip = index % 5 == 0
        fields = _party_fields(index, is_ip)
        party = store.add_counterparty("ip" if is_ip else "ooo", fields, save=False)
        if documents_planned <= 0:
            continue
        # Договор — основание для остальных документов.
        contract = store.add_document(
            storage.KIND_CONTRACT, party["id"], f"{index}", "15",
            data=dict(fields), extra={}, text="", save=False,
        )
        docs_created += 1
        # Доп. соглашение — каждый второй контрагент.
        if docs_created < documents_planned and index % 2 == 0:
            store.add_document(
                storage.KIND_ADDITIONAL, party["id"], f"ДС-{index}", "15",
                data=dict(fields),
                extra={"limit": "500 000", "term": "30",
                       "valid_until": f"31.12.{2026 + index % 3}", "contract_id": contract["id"]},
                text="", save=False,
            )
            docs_created += 1
        # Спецификации с позициями и текстом 2–5 КБ — до исчерпания плана.
        while docs_created < documents_planned:
            store.add_document(
                storage.KIND_SPEC, party["id"], f"{index}-{docs_created}", "20",
                data=dict(fields),
                extra={
                    "contract_id": contract["id"],
                    "items": _items(rng, rng.randint(5, 15)),
                    "total": 0,
                    "spec_day": "20", "spec_month": "июня",
                },
                text=_synthetic_text(rng, rng.uniform(2.0, 5.0)),
                save=False,
            )
            docs_created += 1

    save_started = time.perf_counter()
    ok = store.save()
    save_ms = (time.perf_counter() - save_started) * 1000
    total_ms = (time.perf_counter() - started) * 1000
    size = db_file.stat().st_size if db_file.exists() else 0
    return {
        "out": str(db_file),
        "parties": parties,
        "documents": docs_created,
        "records": parties + docs_created,
        "save_ok": ok,
        "save_ms": save_ms,
        "total_ms": total_ms,
        "db_bytes": size,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Синтетический реестр EW.DocHub для baseline-замеров (DIAG-1)",
    )
    parser.add_argument("--out", required=True, help="каталог для database.db")
    parser.add_argument("--records", type=int, required=True,
                        help="целевое число записей (контрагенты + документы)")
    args = parser.parse_args()
    result = generate(args.out, args.records)
    print(
        f"Записей: {result['records']} (контрагентов {result['parties']}, "
        f"документов {result['documents']})\n"
        f"База: {result['out']} ({result['db_bytes'] / (1024 * 1024):.1f} МБ)\n"
        f"save(): {result['save_ms']:.0f} мс (ok={result['save_ok']}), "
        f"всего: {result['total_ms'] / 1000:.1f} с"
    )
    return 0 if result["save_ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
