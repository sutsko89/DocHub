"""Тесты сроков действия доп. соглашений (2.18.0) — agreements.py.

Без Tk: классификация сроков (просрочено / истекает / норма), хвосты
строк реестра, текст стартового уведомления, игнорирование записей без
даты (созданных до 2.18.0).
"""
from __future__ import annotations

from datetime import date

import agreements
import storage

TODAY = date(2026, 11, 20)


def ensure_party(store):
    """Фикстура store пуста — добавляем контрагента лениво."""
    if not store.counterparties:
        store.add_counterparty("ooo", {
            "Сокращенное название компании": "ООО Ромашка"})


def add_additional(store, valid_iso, number="3", extra_tail=None):
    ensure_party(store)
    extra = {"limit": "100000", "term": "30"}
    if extra_tail:
        extra.update(extra_tail)
    if valid_iso is not None:
        extra["valid_until_iso"] = valid_iso
    return store.add_document(
        storage.KIND_ADDITIONAL, store.counterparties[0]["id"],
        number, "01.10.2026", {"Номер договора": "5"}, extra=extra,
    )


# --------------------------------------------------------------------------- #
#  collect: сортировка и buckets
# --------------------------------------------------------------------------- #
def test_collect_empty_registry_has_no_buckets(store):
    assert agreements.collect(store, today=TODAY) == {"expired": [], "soon": []}


def test_collect_expired_and_soon_sorted(store):
    late = add_additional(store, "2026-10-01", number="1")   # просрочено
    soon = add_additional(store, "2026-12-15", number="2")   # в месяце
    fresh = add_additional(store, "2027-06-30", number="4")  # далеко

    buckets = agreements.collect(store, today=TODAY)

    assert buckets["expired"] == [late]
    assert buckets["soon"] == [soon]
    # Свежий документ не попадает ни в один список.
    assert fresh not in buckets["expired"] + buckets["soon"]


def test_collect_soon_sorted_by_date(store):
    near = add_additional(store, "2026-11-25", number="1")
    far = add_additional(store, "2026-12-15", number="2")
    buckets = agreements.collect(store, today=TODAY)
    assert buckets["soon"] == [near, far], "раньше истекает — выше"


def test_collect_ignores_records_without_date(store):
    add_additional(store, None, number="5")  # запись до 2.18.0
    assert agreements.collect(store, today=TODAY) == {"expired": [], "soon": []}


def test_collect_ignores_contracts_and_trashed(store):
    ensure_party(store)
    store.add_document(storage.KIND_CONTRACT, store.counterparties[0]["id"],
                       "5", "01.10.2026", {},
                       extra={"valid_until_iso": "2026-01-01"})
    trashed = add_additional(store, "2026-01-01", number="6")
    trashed["deleted_at"] = "2026-11-01T00:00:00"

    assert agreements.collect(store, today=TODAY) == {"expired": [], "soon": []}


# --------------------------------------------------------------------------- #
#  collect_state / row_suffix
# --------------------------------------------------------------------------- #
def test_collect_state_expired_soon_none(store):
    expired = add_additional(store, "2026-01-01", number="1")
    soon = add_additional(store, "2026-12-15", number="2")
    fresh = add_additional(store, "2027-06-30", number="3")

    assert agreements.collect_state(expired, today=TODAY) == "expired"
    assert agreements.collect_state(soon, today=TODAY) == "soon"
    assert agreements.collect_state(fresh, today=TODAY) is None


def test_row_suffix_for_expired_soon_and_fresh(store):
    expired = add_additional(store, "2026-01-01", number="1")
    soon = add_additional(store, "2026-12-15", number="2")
    fresh = add_additional(store, "2027-06-30", number="3")

    assert "истекло 01.01.2026" in agreements.row_suffix(expired, today=TODAY)
    assert "истекает 15.12.2026" in agreements.row_suffix(soon, today=TODAY)
    assert "действует до 30.06.2027" in agreements.row_suffix(fresh, today=TODAY)


def test_row_suffix_without_date_is_empty(store):
    rec = add_additional(store, None, number="5")
    assert agreements.row_suffix(rec) == ""


# --------------------------------------------------------------------------- #
#  notice_text
# --------------------------------------------------------------------------- #
def test_notice_text_none_when_no_deadlines(store):
    add_additional(store, "2027-06-30", number="1")
    assert agreements.notice_text(store, today=TODAY) is None


def test_notice_text_lists_expired_first(store):
    add_additional(store, "2026-10-01", number="1")
    add_additional(store, "2026-12-15", number="2")

    text = agreements.notice_text(store, today=TODAY)

    assert "Просроченные" in text
    assert "№ 1" in text and "истекло 01.10.2026" in text
    assert "№ 2" in text and "истекает 15.12.2026" in text


def test_notice_text_mentions_party_name(store):
    add_additional(store, "2026-12-15", number="7")
    text = agreements.notice_text(store, today=TODAY)
    assert "ООО Ромашка" in text, "контрагент опознаваем без открытия реестра"
