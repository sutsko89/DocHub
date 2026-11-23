#!/usr/bin/env python3
"""Прототип-замер TREE-1: строки CTk-виджетов против ttk.Treeview (раздел Е).

Decision-прототип для карточки TREE-1 (этап 4 roadmap rev2): на
детерминированных наборах 1k/10k/50k измеряются три операции реестра.

1. Популяция текущим способом — CTkScrollableFrame + строка из CTk-виджетов
   (рамка, кнопка строки, две ghost-кнопки действий), как в refresh_full
   2.24.0. Это воспроизведение механизма затрат P-02 «пересоздание ~10
   виджетов на строку» (П-аудит); на 50k заведомо минуты, поэтому арм
   ограничен 1k/10k.
2. Популяция ttk.Treeview — те же записи, пять колонок значений, теги
   сроков. Порог раздела Е: первое отображение @10k <= 300 мс.
3. Точечные операции Treeview: item() на 1000 строк (обновление без
   пересоздания), вставка/удаление 100 строк (карта id->iid), yview-скролл
   (порог раздела Е: тики прокрутки <= 50 мс).

Использование:
    python tools/bench_tree_prototype.py --records 10000 --arms tree
    python tools/bench_tree_prototype.py --records 1000 --arms all

Результат печатается в stdout; абсолютные значения Linux/Xvfb не
подставляются в обещания эффекта — на Windows протокол снимается тем же
скриптом (раздел Е, «до/после» на одной машине и данных).
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from document_builder import F  # noqa: E402

ARMS_ALL = ("rows", "tree")


def make_parties(count: int) -> list[dict]:
    """Синтетические контрагенты без записи на диск (данные для списков)."""
    parties = []
    for i in range(count):
        is_ip = i % 4 == 0
        inn = f"{77010000000 + i * 137}"[:10] if not is_ip else f"{321300000000 + i * 137}"[:12]
        parties.append({
            "id": f"p-{i}",
            "type": "ip" if is_ip else "ooo",
            "fields": {
                F.INN: inn,
                F.SHORT_NAME: (f"ИП Предприниматель-{i}" if is_ip
                               else f"ООО Компания-{i} и партнёры"),
                F.LEGAL_ADDRESS: f"г. Владивосток, ул. Светланская, д. {i % 99 + 1}, оф. {i}",
            },
        })
    return parties


def party_row_values(rec: dict) -> list[str]:
    """Значения колонок строки контрагента (как в TREE-1: U1)."""
    fields = rec.get("fields") or {}
    kind = "ИП" if rec.get("type") == "ip" else "Юр. лицо"
    inn = str(fields.get(F.INN, "")).strip() or "—"
    address = str(fields.get(F.LEGAL_ADDRESS, "")).strip()
    if len(address) > 48:
        address = address[:47].rstrip() + "…"
    return [str(rec.get("fields", {}).get(F.SHORT_NAME, "")), kind, inn,
            str(rec.get("id")), address]


# --------------------------------------------------------------------------- #
#  Арм 1: текущая популяция строк из CTk-виджетов (механизм P-02)
# --------------------------------------------------------------------------- #
def bench_rows(parties: list[dict]) -> dict:
    import customtkinter as ctk
    import ui_macos as ui

    root = ctk.CTk()
    root.geometry("1200x800+0+0")
    try:
        root.update_idletasks()
        frame = ctk.CTkScrollableFrame(root, fg_color="transparent", corner_radius=0)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        root.update()  # окно должно быть отображено: CTk рисует во внутренние канвасы

        def one_pass() -> float:
            for child in frame.winfo_children():
                child.destroy()
            start = time.perf_counter()
            for rec in parties:
                row = ctk.CTkFrame(frame, fg_color="transparent")
                row.grid(row=rec["_i"] + 1, column=0, sticky="ew", pady=1)
                row.grid_columnconfigure(0, weight=1)
                btn = ctk.CTkButton(
                    row, text=f"{rec['fields'][F.SHORT_NAME]} · {rec['type']} · ИНН "
                              f"{rec['fields'][F.INN]} · ID {rec['id']}",
                    anchor="w", fg_color="transparent",
                    hover_color=ui.COLORS["sidebar_hover"],
                    text_color=ui.COLORS["text_secondary"], corner_radius=8,
                    height=30,
                )
                btn.grid(row=0, column=0, sticky="ew")
                ui.ghost_button(row, "Редактировать", lambda: None, width=110, height=26)
                ui.ghost_button(row, "Удалить", lambda: None, width=92, height=26)
            return (time.perf_counter() - start) * 1000.0

        first = one_pass()
        second = one_pass()
    finally:
        root.destroy()
    return {"populate_cold_ms": round(first, 1), "populate_warm_ms": round(second, 1)}


# --------------------------------------------------------------------------- #
#  Арм 2: ttk.Treeview
# --------------------------------------------------------------------------- #
def bench_tree(parties: list[dict]) -> dict:
    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.geometry("1200x800")
    try:
        root.withdraw()
        style = ttk.Style(root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("R.Treeview", rowheight=28, borderwidth=0)
        columns = ("name", "kind", "inn", "id", "address")
        tree = ttk.Treeview(root, columns=columns, show="headings",
                            selectmode="browse", style="R.Treeview")
        for key, title, width, anchor in (
            ("name", "Контрагент", 320, "w"), ("kind", "Тип", 80, "w"),
            ("inn", "ИНН", 120, "w"), ("id", "ID", 90, "w"),
            ("address", "Адрес", 380, "w"),
        ):
            tree.heading(key, text=title)
            tree.column(key, width=width, anchor=anchor, stretch=(key in ("name", "address")))
        tree.tag_configure("expired", foreground="#C92A2A")
        tree.tag_configure("soon", foreground="#B45309")
        tree.pack(fill="both", expand=True)
        root.update_idletasks()

        # Популяция: все строки одним проходом (значения + тег у части строк).
        def populate() -> float:
            start = time.perf_counter()
            for rec in parties:
                tags = ("expired",) if rec["_i"] % 97 == 0 else (
                    ("soon",) if rec["_i"] % 89 == 0 else ("",))
                tree.insert("", "end", iid=rec["id"], tags=tags,
                            values=party_row_values(rec))
            root.update_idletasks()
            return (time.perf_counter() - start) * 1000.0

        populate_ms = populate()

        # Точечные обновления item() на 1000 строк (как refresh_selection).
        sample = parties[:1000]
        start = time.perf_counter()
        for rec in sample:
            tree.item(rec["id"], values=party_row_values(rec))
        item_ms = (time.perf_counter() - start) * 1000.0

        # Точечные вставка/удаление 100 строк (как diff-обновление состава).
        extra = [f"tmp-{i}" for i in range(100)]
        start = time.perf_counter()
        for iid in extra:
            tree.insert("", "end", iid=iid, values=party_row_values(parties[0]))
        for iid in extra:
            tree.delete(iid)
        ins_del_ms = (time.perf_counter() - start) * 1000.0

        # Прокрутка: 100 шагов yview (тиков прокрутки), порог <= 50 мс на тик.
        root.update()
        tick_max = 0.0
        for step in range(100):
            frac = step / 100.0
            t0 = time.perf_counter()
            tree.yview_moveto(frac)
            root.update_idletasks()
            tick_max = max(tick_max, (time.perf_counter() - t0) * 1000.0)

        # Выделение: selection_set на 200 случайных строк.
        start = time.perf_counter()
        for rec in parties[:200]:
            tree.selection_set(rec["id"])
        sel_ms = (time.perf_counter() - start) * 1000.0
    finally:
        root.destroy()
    return {
        "populate_ms": round(populate_ms, 1),
        "item_1000_ms": round(item_ms, 1),
        "ins_del_100_ms": round(ins_del_ms, 1),
        "scroll_max_tick_ms": round(tick_max, 2),
        "select_200_ms": round(sel_ms, 1),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--records", type=int, default=10000)
    parser.add_argument("--arms", choices=ARMS_ALL + ("all",), default="tree")
    args = parser.parse_args()

    parties = make_parties(args.records)
    for i, rec in enumerate(parties):
        rec["_i"] = i

    print(f"records={args.records} arms={args.arms}")
    if args.arms in ("rows", "all"):
        # 50k в этом арме не измеряется: строки-виджеты на 50k — минуты (P-02).
        if args.records > 10000:
            print("rows: пропущен (арм строк-виджетов измеряется на 1k/10k)")
        else:
            print(f"rows : {bench_rows(parties)}")
    if args.arms in ("tree", "all"):
        print(f"tree : {bench_tree(parties)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
