#!/usr/bin/env python3
"""Замер P-08 (UPD-UI, этап 5): сборка и запись .docx на UI-потоке.

Карточка P-08 — «сборка и запись .docx синхронно на UI-потоке»; решение
по замеру (roadmap, этап 5). Скрипт снимает ориентирующий замер полного
пути «собрать документ из шаблона + атомарная запись» для договора и
спецификации на синтетических данных:

    python tools/bench_docx_build.py [--runs 20] [--items 30]

Замер выполняется на текущей машине (в песочнице — Linux/Xvfb) и даёт
ОРИЕНТИР, а не Windows-значение: целевые контрольные замеры этапа 5 идут
по протоколу docs/protocol-section-e.md (канал perf «Запись .docx»).
Вывод: среднее/максимум мс на «сборка», «запись», «сборка+запись».
"""
from __future__ import annotations

import argparse
import os
import statistics
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import services  # noqa: E402
import storage  # noqa: E402
from document_builder import F  # noqa: E402

CONTRACT_FIELDS = {
    F.SHORT_NAME: "ООО Замер Плюс",
    F.FULL_NAME: "Общество с ограниченной ответственностью «Замер Плюс»",
    F.INN: "7701234567", F.KPP: "770101001", F.OGRN: "1037739010221",
    F.DIRECTOR: "Иванов Иван Иванович",
    F.LEGAL_ADDRESS: "г. Москва, ул. Замерная, д. 1, офис 100",
    F.BANK: "Тест-Банк", F.BIK: "044525225",
    F.ACCOUNT: "40702810100000001234",
    F.CORR_ACCOUNT: "30101810100000000225",
    F.NUMBER: "12", F.DAY: "23", F.MONTH: "ноября",
}


def spec_rec(items: int) -> dict:
    """Запись спецификации с N позициями (имя, кол-во, ед., цена)."""
    rows = [
        {"name": f"Позиция замера {i + 1}", "qty": str(i + 1),
         "unit": "шт", "price": f"{(i + 1) * 137}.50"}
        for i in range(items)
    ]
    return {
        "kind": storage.KIND_SPEC,
        "number": "5", "date": "23.11.2026",
        "data": {**CONTRACT_FIELDS, F.NUMBER: "5"},
        "extra": {"rows": rows, "total": sum((i + 1) * 137.5 for i in range(items))},
        "text": "",
    }


class FakeStore:
    """Минимум для render_spec: контрагент без ИП-признака."""

    def party_name(self, party_id):
        return "ООО Замер Плюс"

    def get_counterparty(self, party_id):
        return None

    def document_body(self, doc_id):
        return {"text": "", "data": {}, "extra": {}}


def measure(label: str, fn, runs: int) -> dict:
    times = []
    for _ in range(runs):
        t0 = time.perf_counter()
        fn()
        times.append((time.perf_counter() - t0) * 1000)
    return {
        "label": label,
        "avg": statistics.mean(times),
        "max": max(times),
        "min": min(times),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runs", type=int, default=20)
    parser.add_argument("--items", type=int, default=30,
                        help="позиций в спецификации")
    args = parser.parse_args()

    import document_builder  # после вставки корня в sys.path

    store = FakeStore()
    out_dir = tempfile.mkdtemp(prefix="dochub-p08-")

    contract_doc = services.render_contract(dict(CONTRACT_FIELDS), is_ip=False)
    spec_doc = services.render_spec(
        spec_rec(args.items), store, document_builder.SPEC_SETTINGS
        if hasattr(document_builder, "SPEC_SETTINGS") else None
    )

    results = [
        measure("договор: сборка",
                lambda: services.render_contract(dict(CONTRACT_FIELDS), False),
                args.runs),
        measure("договор: запись",
                lambda: services.save_docx_atomic(contract_doc, os.path.join(out_dir, "contract.docx")),
                args.runs),
        measure(f"спецификация ({args.items} поз.): сборка",
                lambda: services.render_spec(spec_rec(args.items), store, None),
                args.runs),
        measure("спецификация: запись",
                lambda: services.save_docx_atomic(spec_doc, os.path.join(out_dir, "spec.docx")),
                args.runs),
    ]

    print(f"P-08 ориентирующий замер (runs={args.runs}, "
          f"позиций={args.items}): сборка + атомарная запись .docx")
    for r in results:
        print(f"  {r['label']:<38} среднее {r['avg']:7.1f} мс "
              f"(мин {r['min']:.1f} / макс {r['max']:.1f})")
    build_sum = results[0]["avg"] + results[2]["avg"]
    write_sum = results[1]["avg"] + results[3]["avg"]
    print(f"  сборка+запись, суммарно среднее        "
          f"{build_sum + write_sum:7.1f} мс")

    import shutil
    shutil.rmtree(out_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
