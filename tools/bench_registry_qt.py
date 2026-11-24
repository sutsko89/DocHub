#!/usr/bin/env python3
"""Замер реестра Tk (Treeview) vs Qt (QTableView) на одной машине (этап M5).

Карточка TREE-1 / E-3: «первое отображение и прокрутка» на наборах
1k / 10k / 50k, сравнение с цифрами 2.25.2 Treeview (7,6 мс @1k /
413 мс @50k на Windows) и попарное сравнение Tk↔Qt на той же машине.
Абсолютные значения Linux не подставляются в обещания эффекта — на
Windows протокол снимается этим же скриптом (runbook M5).

Измерения (одинаковые синтетические данные, одинаковый порядок):

1. populate  — полная популяция N строк + первый проход раскладки
               (Tk: insert + update_idletasks; Qt: reset модели +
               processEvents с показанным виджетом);
2. diff      — «Показать ещё»: диф-добавление 10 000 строк к базе 10 000;
3. update100 — обновление значений 100 строк без пересоздания (R-04);
4. scroll    — 20 тиков прокрутки до конца и назад.

Запуск:
    python3 tools/bench_registry_qt.py              # Tk+Qt (нужен дисплей/Xvfb)
    python3 tools/bench_registry_qt.py --arms qt    # только Qt (offscreen)
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import registry_core as rcore
from document_builder import F


def make_rows(count: int) -> list:
    """Синтетические строки реестра (детерминированные, без записи на диск)."""
    rows = []
    for i in range(count):
        rec = {
            "id": i + 1,
            "type": "ip" if i % 5 == 0 else "ooo",
            "fields": {
                F.INN: f"7701{i:08d}",
                F.SHORT_NAME: f"ООО Компания-{i}",
                F.LEGAL_ADDRESS: f"г. Владивосток, ул. Рабочая, д. {i % 97 + 1}, оф. {i % 310}",
            },
        }
        rows.append(rcore.RowSpec(
            iid=f"p:{rec['id']}",
            values=rcore.party_row_values(rec),
            dot=rcore.party_row_dot(rec),
            data={"kind": "party", "rec": rec},
        ))
    return rows


def make_parties(count: int) -> list[dict]:
    """Те же записи в виде rec (для Tk-ветки значений строк)."""
    rows = make_rows(count)
    return [row.data["rec"] for row in rows]


# --------------------------------------------------------------------------- #
#  Tk-ветка
# --------------------------------------------------------------------------- #
def bench_tk(count: int, ticks: int = 20) -> dict:
    import tkinter as tk

    import registry_tree

    root = tk.Tk()
    root.withdraw()
    root.geometry("900x520+0+0")
    widget = registry_tree.RegistryTree(
        root, on_select=lambda *_a: None, on_activate=lambda *_a: None,
        columns=rcore.PARTY_COLUMNS,
    )
    widget.pack(fill="both", expand=True)
    rows = make_rows(count)

    result = {}
    start = time.perf_counter()
    widget.populate(rows)
    root.update_idletasks()
    result["populate_ms"] = (time.perf_counter() - start) * 1000.0

    # «Показать ещё»: диф-добавление к 10k ещё 10k.
    if count == 10_000:
        bigger = make_rows(20_000)
        start = time.perf_counter()
        widget.apply_diff(bigger)
        root.update_idletasks()
        result["diff_ms"] = (time.perf_counter() - start) * 1000.0

    # Обновление 100 строк без пересоздания.
    changed = [rcore.RowSpec(
        iid=row.iid, values=(f"ООО Правка-{i}",) + row.values[1:],
        tags=row.tags, dot=row.dot,
    ) for i, row in enumerate(rows[:100])]
    start = time.perf_counter()
    widget.apply_diff(changed + make_rows(count)[100:])
    root.update_idletasks()
    result["update100_ms"] = (time.perf_counter() - start) * 1000.0

    # Прокрутка: 20 тиков до конца/в начало.
    samples = []
    for i in range(ticks):
        target = 1.0 if i % 2 == 0 else 0.0
        start = time.perf_counter()
        widget.tree.yview_moveto(target)
        root.update_idletasks()
        samples.append((time.perf_counter() - start) * 1000.0)
    result["scroll_avg_ms"] = sum(samples) / len(samples)

    root.destroy()
    return result


# --------------------------------------------------------------------------- #
#  Qt-ветка
# --------------------------------------------------------------------------- #
def bench_qt(count: int, ticks: int = 20) -> dict:
    from PySide6.QtWidgets import QApplication

    import registry_page_qt

    qt_app = QApplication.instance() or QApplication([])
    holder = qt_app

    table = registry_page_qt.RegistryTableQt(
        on_select=lambda *_a: None, on_activate=lambda *_a: None,
        on_delete_key=lambda: None,
    )
    table.resize(880, 500)
    table.show()
    rows = make_rows(count)

    result = {}
    start = time.perf_counter()
    table.populate(rcore.PARTY_COLUMNS, rows)
    qt_app.processEvents()
    result["populate_ms"] = (time.perf_counter() - start) * 1000.0

    if count == 10_000:
        bigger = make_rows(20_000)
        start = time.perf_counter()
        table.apply_diff(bigger)
        qt_app.processEvents()
        result["diff_ms"] = (time.perf_counter() - start) * 1000.0
        table.apply_diff(rows)
        qt_app.processEvents()

    changed = [rcore.RowSpec(
        iid=row.iid, values=(f"ООО Правка-{i}",) + row.values[1:],
        tags=row.tags, dot=row.dot,
    ) for i, row in enumerate(rows[:100])]
    start = time.perf_counter()
    table.apply_diff(changed + make_rows(count)[100:])
    qt_app.processEvents()
    result["update100_ms"] = (time.perf_counter() - start) * 1000.0

    bar = table.verticalScrollBar()
    samples = []
    for i in range(ticks):
        target = bar.maximum() if i % 2 == 0 else bar.minimum()
        start = time.perf_counter()
        bar.setValue(target)
        qt_app.processEvents()
        samples.append((time.perf_counter() - start) * 1000.0)
    result["scroll_avg_ms"] = sum(samples) / len(samples)

    table.hide()
    table.deleteLater()
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--arms", default="tk,qt", help="tk,qt — какие ветки мерить")
    parser.add_argument("--ticks", type=int, default=20, help="тиков прокрутки")
    args = parser.parse_args()
    arms = [arm.strip() for arm in args.arms.split(",") if arm.strip()]

    sizes = (1_000, 10_000, 50_000)
    results: dict[str, dict] = {}
    for arm in arms:
        results[arm] = {}
        for size in sizes:
            results[arm][size] = bench_tk(size, args.ticks) if arm == "tk" \
                else bench_qt(size, args.ticks)

    print(f"Замер реестра (одна машина, одинаковые данные; тиков прокрутки: {args.ticks})")
    print(f"{'набор':>8} | {'в arm':>4} | {'populate, мс':>12} | {'diff, мс':>9} | "
          f"{'update100, мс':>13} | {'scroll, мс':>10}")
    for arm in arms:
        for size in sizes:
            r = results[arm][size]
            print(f"{size:>8} | {arm:>4} | {r['populate_ms']:>12.1f} | "
                  f"{r.get('diff_ms', float('nan')):>9.1f} | "
                  f"{r['update100_ms']:>13.1f} | {r['scroll_avg_ms']:>10.2f}")
    print("\nПримечание: diff/update100 меряются на базе 10k (в прочих строках — прочерк).")
    print("Windows-цифры снимаются этим же скриптом по runbook M5 (К-2).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
