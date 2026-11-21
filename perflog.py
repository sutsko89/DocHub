"""Perf-журнал производительности (DIAG-1, этап 1 roadmap).

Включается env-флагом ``EWDOCHUB_PERF_LOG=1`` при запуске. По умолчанию
выключен: без флага поведение программы идентично прежнему — ни файлов,
ни записей, ни обработчиков.

Флаг читается один раз за сеанс (кэш в :func:`enabled`, сброс для тестов —
:func:`reset_for_tests`). Записи буферизуются и сбрасываются в
``logs/perf.log`` порциями, чтобы замер не искажал измеряемое: критичный
путь не содержит дисковых операций на каждый тик.

Формат строки: ``<дата-время> <kind> <source> <значение_мс>``. Персональных
данных нет — только латинские имена источников и числа.

Два класса метрик (roadmap, раздел Е; смешивать запрещено):

* ``tick`` — dt между последовательными событиями (тики пулера переноса/
  растягивания, раскладки канвы реестра). Выборка — все тики одного
  прогона; p50/p95/max считаются внутри каждого прогона отдельно.
* ``duration`` — длительности операций; дублирует существующие точки
  ``applog.timed`` (порог 500 мс в app.log не меняется), но в perf-лог
  попадает каждая операция, без порога.
"""

from __future__ import annotations

import atexit
import os
import time
from datetime import datetime

ENV_FLAG = "EWDOCHUB_PERF_LOG"
PERF_FILE_NAME = "perf.log"
FLUSH_EVERY = 256          # записей в буфере до сброса на диск

_enabled: bool | None = None
_buffer: list[str] = []


def perf_path() -> str:
    """Путь к perf-журналу (рядом с app.log, в профиле пользователя)."""
    import applog  # поздний импорт: applog не зависит от perflog

    return os.path.join(applog.log_dir(), PERF_FILE_NAME)


def enabled() -> bool:
    """Флаг ``EWDOCHUB_PERF_LOG`` читается один раз при первом обращении."""
    global _enabled
    if _enabled is None:
        _enabled = os.environ.get(ENV_FLAG, "").strip().lower() in ("1", "true", "yes", "on")
    return _enabled


def reset_for_tests() -> None:
    """Сбрасывает кэш флага и буфер (для автотестов)."""
    global _enabled
    _enabled = None
    _buffer.clear()


def _line(kind: str, source: str, value_ms: float) -> str:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    return f"{stamp} {kind} {source} {value_ms:.3f}"


def record(source: str, value_ms: float, kind: str = "duration") -> None:
    """Добавляет запись в буфер; сброс на диск — порциями (FLUSH_EVERY)."""
    if not enabled():
        return
    _buffer.append(_line(kind, source, value_ms))
    if len(_buffer) >= FLUSH_EVERY:
        flush()


def flush() -> None:
    """Пишет буфер в perf.log. Любая ошибка глушится: замер не роняет программу."""
    if not _buffer:
        return
    chunk = _buffer[:]
    _buffer.clear()
    try:
        import applog

        os.makedirs(applog.log_dir(), exist_ok=True)
        with open(perf_path(), "a", encoding="utf-8") as fh:
            fh.write("\n".join(chunk) + "\n")
    except OSError:
        pass


atexit.register(flush)


class TickMeter:
    """dt между последовательными вызовами :meth:`tick`.

    Первый вызов только фиксирует отсчёт. Когда флаг выключен, tick()
    сразу возвращается — стоимость замера на критичном пути нулевая.
    """

    __slots__ = ("_source", "_last")

    def __init__(self, source: str):
        self._source = source
        self._last: float | None = None

    def tick(self) -> None:
        if not enabled():
            return
        now = time.perf_counter()
        if self._last is not None:
            record(self._source, (now - self._last) * 1000.0, kind="tick")
        self._last = now


def bind_canvas_meter(widget, source: str):
    """Вешает тик-метр раскладок канвы на ``<Configure>`` (реестр: прокрутка).

    Обработчик ставится только при включённом флаге: без флага привязки
    нет вовсе. Возвращает метр (или None, если привязать не удалось).
    """
    if not enabled() or widget is None:
        return None
    meter = TickMeter(source)

    def _on_configure(_event=None) -> None:
        meter.tick()

    try:
        widget.bind("<Configure>", _on_configure, add="+")
    except Exception:
        return None
    return meter
