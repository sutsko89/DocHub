"""Сроки действия доп. соглашений (2.18.0) — сбор и уведомления.

Чистый модуль без Tk: по реестру находит активные доп. соглашения с
заполненным «действует до» (extra["valid_until_iso"]) и делит их на
просроченные («expired») и истекающие в течение месяца («soon»).

Используют:

* стартовое уведомление ContractApp (``notice_text`` + диалог
  «Открыть реестр» / «Позже», раз в сутки через settings.json);
* страница «Реестр»: красный текст строки просроченного, амберный —
  истекающего, плюс баннер сверху списка.

Записи без даты (созданные до 2.18.0) считаются бессрочными и никого
не тревожат: срок известен только для документов, созданных после
введения поля «Действует до».
"""

from __future__ import annotations

from datetime import date

import storage
from document_builder import (
    agreement_expiry_state, parse_valid_until, valid_until_display,
)


def _until_of(rec: dict, store=None, until=None) -> "str | None":
    """Срок «действует до» записи (LAZY-1).

    Источники по приоритету: явный ``until``; тело в памяти
    (rec["extra"] — свежие мутаторы и кэш document_body); лёгкая карта
    сроков хранилища ``store.expiry_until(id)`` для ленивых записей.
    """
    if until is not None:
        return until
    extra = rec.get("extra")
    if isinstance(extra, dict):
        return extra.get("valid_until_iso") or extra.get("valid_until")
    if store is not None:
        return store.expiry_until(rec.get("id"))
    return None


def with_dates(recs: list[dict], store=None) -> list[dict]:
    """Активные доп. соглашения, у которых срок известен и читается.

    Возвращает пары (запись, date) — записи без корректной даты не
    попадают ни в один список уведомлений.
    """
    result: list[tuple[dict, date]] = []
    for rec in recs:
        if rec.get("kind") != storage.KIND_ADDITIONAL or storage.is_deleted(rec):
            continue
        parsed = parse_valid_until(_until_of(rec, store=store))
        if parsed is not None:
            result.append((rec, parsed))
    return result


def collect(store, today: date | None = None,
            docs: "list[dict] | None" = None) -> dict[str, list[dict]]:
    """Делит доп. соглашения на «expired» и «soon».

    Сроки берутся из лёгкой карты хранилища — рендер реестра и рабочий
    поток старта не читают базу (LAZY-1). ``docs`` — снимок списка записей
    для рабочего потока (START-1): итерация по неизменяемому списку вместо
    живого хранилища.
    """
    buckets: dict[str, list[dict]] = {"expired": [], "soon": []}
    current = date.today() if today is None else today
    scored: list[tuple[date, str, dict]] = []
    for rec, until in with_dates(
            store.active_documents() if docs is None else docs, store=store):
        state = agreement_expiry_state(until, today=current)
        if state in buckets:
            scored.append((until, state, rec))
    scored.sort(key=lambda item: item[0])
    for until, state, rec in scored:
        buckets[state].append(rec)
    return buckets


def collect_state(rec: dict, today: date | None = None,
                  until: "str | None" = None) -> "str | None":
    """Состояние срока одной записи реестра: «expired»/«soon»/None (окраска)."""
    value = until if until is not None else _until_of(rec)
    return agreement_expiry_state(value, today=today)


def row_suffix(rec: dict, today: date | None = None,
               until: "str | None" = None) -> str:
    """Хвост строки документа в реестре: « · действует до 31.12.2026».

    Для просроченного — «истекло 20.10.2026», для истекающего —
    «истекает 15.12.2026» (чтобы причина выделения была видна сразу).
    Документы без срока возвращают пустую строку.
    """
    value = until if until is not None else _until_of(rec)
    display = valid_until_display(value)
    if not display:
        return ""
    state = collect_state(rec, today=today, until=until)
    if state == "expired":
        return f"   ·   истекло {display}"
    if state == "soon":
        return f"   ·   истекает {display}"
    return f"   ·   действует до {display}"


def notice_text(store, today: date | None = None,
                docs: "list[dict] | None" = None) -> str | None:
    """Текст стартового уведомления о сроках. None — тревожить не о чем.

    Просроченные перечисляются первыми: «истекло DD.MM.YYYY», истекающие —
    «истекает DD.MM.YYYY». Имена контрагентов добавляются, чтобы запись
    можно было опознать без открытия реестра. ``docs`` — снимок списка
    записей для рабочего потока старта (START-1).
    """
    buckets = collect(store, today=today, docs=docs)
    expired, soon = buckets["expired"], buckets["soon"]
    if not expired and not soon:
        return None
    lines: list[str] = []
    if expired:
        lines.append("Просроченные доп. соглашения:")
        for rec in expired:
            lines.append(f"• {describe(store, rec, today=today)}")
    if soon:
        if lines:
            lines.append("")
        lines.append(f"Истекает в течение месяца ({len(soon)}):")
        for rec in soon:
            lines.append(f"• {describe(store, rec, today=today)}")
    lines.append("")
    lines.append("Продлите срок или закройте соглашение с контрагентом.")
    return "\n".join(lines)


def describe(store, rec: dict, today: date | None = None) -> str:
    """Строка «Доп. соглашение № 3 от 01.10.2026 — ООО Ромашка · истекло …»."""
    number = rec.get("number") or "б/н"
    date_text = rec.get("date") or "б/д"
    party = store.party_name(rec.get("party_id"))
    value = _until_of(rec, store=store)
    display = valid_until_display(value)
    state = agreement_expiry_state(value, today=today)
    tail = "истекло " + display if state == "expired" else "истекает " + display
    name = f"{party} · " if party else ""
    return f"№ {number} от {date_text} — {name}{tail}"
