#!/usr/bin/env python3
"""Baseline-замер записи базы STORE-1A и perf-отчёт (раздел Е roadmap rev2).

Инструмент «до/после» для карточки STORE-1A (R-01 инкрементальная запись,
R-05a счётчики статусной строки). Запускается на целевой Windows на ОДНОЙ
машине и ОДНИХ данных в двух состояниях кода:

  * «до»  — коммит 7465dfa (v2.23.0): save() перезаписывает базу целиком,
    счётчики статусной строки строятся по спискам в памяти (P-10);
  * «после» — текущий main (>= 2.24.0): save() пишет только dirty-строки,
    status_counts() считает COUNT(*) по покрывающим индексам.

Скрипт не зависит от состояния кода, в котором он запущен: отсутствие
status_counts()/save_full() обрабатывается откатом к старому механизму с
пометкой в выводе. Работает ТОЛЬКО с копией базы: источник не изменяется,
все записи идут во временный каталог %TEMP% и удаляются после прогона.

Использование:
    python tools/bench_store_1a.py --db <каталог>\\database.db --repeats 5
    python tools/bench_store_1a.py --report <каталог>\\logs\\perf.log

Критерий аудита (раздел Е): точечная save() <= 15 мс UI-времени @10k.
Данные готовит tools/generate_registry.py (детерминированные, одинаковые
на всех прогонах «до» и «после»).
"""

from __future__ import annotations

import argparse
import shutil
import statistics
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import storage  # noqa: E402
from document_builder import F  # noqa: E402


# --------------------------------------------------------------------------- #
#  Копия базы: источник не трогаем
# --------------------------------------------------------------------------- #
def copy_db(source: Path, workdir: Path) -> Path:
    """Копирует database.db (+ WAL/SHM, если остались) во временный каталог."""
    target = workdir / source.name
    for suffix in ("", "-wal", "-shm"):
        src = Path(str(source) + suffix)
        if src.exists():
            shutil.copy2(src, Path(str(target) + suffix))
    return target


def _stats(values: list[float]) -> dict:
    """Медиана/p95/min/max по списку значений в мс."""
    ordered = sorted(values)
    p95 = ordered[min(len(ordered) - 1, int(round(0.95 * (len(ordered) - 1))))]
    return {
        "n": len(ordered),
        "median": round(statistics.median(ordered), 2),
        "p95": round(p95, 2),
        "min": round(min(ordered), 2),
        "max": round(max(ordered), 2),
    }


def _fmt(stats: dict) -> str:
    return (f"n={stats['n']} медиана={stats['median']} мс p95={stats['p95']} мс "
            f"min={stats['min']} мс max={stats['max']} мс")


# --------------------------------------------------------------------------- #
#  Счётчики: новый механизм и реплика старого (P-10, списки в памяти)
# --------------------------------------------------------------------------- #
def in_memory_counts(store) -> dict:
    """Реплика старого способа: списки активных записей в памяти (P-10)."""
    deleted = storage.is_deleted
    active_parties = [r for r in store.counterparties if not deleted(r)]
    active_docs = [r for r in store.documents if not deleted(r)]
    return {
        "parties": len(active_parties),
        "documents": len(active_docs),
        "trash": len(store.counterparties) - len(active_parties)
                 + len(store.documents) - len(active_docs),
    }


# --------------------------------------------------------------------------- #
#  Основной прогон
# --------------------------------------------------------------------------- #
def run_bench(db_path: Path, repeats: int) -> int:
    if not db_path.exists():
        print(f"ОШИБКА: база не найдена: {db_path}")
        return 1

    workdir = Path(tempfile.mkdtemp(prefix="dochub-bench-"))
    timing: dict[str, list[float]] = {
        "status": [], "noop_save": [], "point_edit": [], "save_full": [],
    }
    try:
        working = copy_db(db_path, workdir)
        # DataStore пишет строго в storage.db_path() — в скрипте-замере
        # (вне прод-кода) путь подменяется на копию во временном каталоге.
        storage.db_path = lambda: str(working)  # type: ignore[assignment]

        store = storage.DataStore()
        parties = store.active_counterparties()
        docs = sum(len(store.documents_for(p["id"])) for p in parties[:20])
        print(f"База-копия: {working}")
        print(f"Записей в памяти: контрагентов {len(parties)}, "
              f"документов в первых 20 карточках: {docs}")
        if not parties:
            print("ОШИБКА: в базе нет активных контрагентов — нечего править")
            return 1

        has_counts = hasattr(store, "status_counts")
        has_full = hasattr(store, "save_full")
        print(f"Механизм счётчиков: {'status_counts() (COUNT(*), R-05a)' if has_counts
              else 'списки в памяти (реплика P-10, старый код)'}")
        print(f"Механизм save(): {'инкрементальный + save_full() доступен (STORE-1A)'
              if has_full else 'полная перезапись (старый код)'}")
        print(f"Повторов: {repeats}\n")

        for i in range(repeats):
            # 1. Счётчики статусной строки (R-05a / P-10).
            start = time.perf_counter()
            counts = store.status_counts() if has_counts else in_memory_counts(store)
            timing["status"].append((time.perf_counter() - start) * 1000.0)

            # 2. save() без изменений: R-01 — не пишет вовсе, ревизия не растёт.
            start = time.perf_counter()
            store.save()
            timing["noop_save"].append((time.perf_counter() - start) * 1000.0)

            # 3. Точечная правка контрагента (мутатор сам вызывает save()):
            #    в UI это ровно операция «Сохранить» карточки.
            rec = parties[i % len(parties)]
            fields = dict(rec.get("fields") or {})
            key = (F.PHONE if F.PHONE in fields else
                   F.EMAIL if F.EMAIL in fields else F.LEGAL_ADDRESS)
            fields[key] = str(fields.get(key, "")) + f"·{i}"
            start = time.perf_counter()
            if hasattr(store, "update_counterparty"):
                store.update_counterparty(rec["id"], fields)
            else:  # очень старый код: правка в памяти + полный save()
                rec["fields"] = fields
                store.save()
            timing["point_edit"].append((time.perf_counter() - start) * 1000.0)

            # 4. Полный путь (только новый код): перезапись всех строк.
            if has_full:
                start = time.perf_counter()
                store.save_full()
                timing["save_full"].append((time.perf_counter() - start) * 1000.0)

            print(f"Прогон {i + 1}/{repeats}: "
                  f"счётчики {timing['status'][-1]:.1f} мс, "
                  f"no-op save {timing['noop_save'][-1]:.1f} мс, "
                  f"точечная правка {timing['point_edit'][-1]:.1f} мс"
                  + (f", save_full {timing['save_full'][-1]:.1f} мс"
                     if has_full else ""))

        print("\n=== Итог (раздел Е, STORE-1A) ===")
        print(f"Счётчики статусной строки : {_fmt(_stats(timing['status']))}")
        print(f"save() без изменений      : {_fmt(_stats(timing['noop_save']))}")
        print(f"Точечная правка + save()  : {_fmt(_stats(timing['point_edit']))}")
        if timing["save_full"]:
            print(f"Полная перезапись save_full(): {_fmt(_stats(timing['save_full']))}")
        print("\nКритерий аудита: точечная правка (медиана) <= 15 мс @10k.")
        if has_counts:
            print(f"status_counts(): {counts}")
        return 0
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


# --------------------------------------------------------------------------- #
#  perf-отчёт (DIAG-1): агрегация logs/perf.log
# --------------------------------------------------------------------------- #
def run_report(log_path: Path, kind: str | None, source: str | None) -> int:
    if not log_path.exists():
        print(f"ОШИБКА: файл не найден: {log_path}")
        return 1
    stats: dict[tuple[str, str], list[float]] = {}
    total = 0
    skipped = 0
    for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
        # Формат: <дата> <время> <kind> <источник…> <значение_мс>. Источник
        # может содержать пробелы («Запись .docx», «Перерисовка реестра»):
        # первые три поля фиксированы, значение — последний токен.
        # Раньше парсер требовал ровно 5 токенов и молча пропускал
        # многословные источники — для E-4/P-08 это недопустимо.
        head = line.split(None, 3)
        if len(head) != 4:
            skipped += 1
            continue
        try:
            s, value_text = head[3].rsplit(None, 1)
            value = float(value_text)
        except ValueError:
            skipped += 1
            continue
        k = head[2]
        if kind and k != kind:
            continue
        if source and not s.startswith(source):
            continue
        stats.setdefault((k, s), []).append(value)
        total += 1
    if not stats:
        print("Подходящих записей нет (формат строки: дата время kind source значение_мс)")
        return 1
    print(f"perf.log: {log_path} — {total} записей"
          + (f" (пропущено неразобранных: {skipped})" if skipped else ""))
    for (k, s), values in sorted(stats.items()):
        print(f"{k:<8} {s:<24} {_fmt(_stats(values))}")
    print("\nСправка: tick — dt между событиями (pointer.poller — пулер жестов, "
          "registry.canvas — раскладка списка); duration — длительность операции; "
          "event — отметка момента (startup.built/startup.ready, значение 0).")
    print("Критерий раздела Е: p95 tick pointer.poller <= 25 мс (PERF-2).")
    print("Критерий этапа 5 (E-4): в первые 5 с после startup.built — нет "
          "duration-записей > 100 мс (главный поток).")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Baseline-замер STORE-1A и агрегатор perf.log (раздел Е)",
    )
    parser.add_argument("--db", help="база для замера (database.db или каталог с ней)")
    parser.add_argument("--repeats", type=int, default=5,
                        help="число повторов замера (по умолчанию 5)")
    parser.add_argument("--report", help="perf.log для агрегации вместо замера")
    parser.add_argument("--kind", choices=("tick", "duration"), help="фильтр kind")
    parser.add_argument("--source", help="префикс источника (pointer.poller, storage…)")
    args = parser.parse_args()

    if args.report:
        return run_report(Path(args.report), args.kind, args.source)
    if not args.db:
        parser.error("укажите --db <database.db> или --report <perf.log>")
    db_path = Path(args.db)
    if db_path.is_dir():
        db_path = db_path / storage.DB_FILENAME
    return run_bench(db_path, max(1, args.repeats))


if __name__ == "__main__":
    sys.exit(main())
