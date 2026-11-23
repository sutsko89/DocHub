"""TREE-1/UX-REG — живые тесты RegistryTree (требуют дисплей/Xvfb).

Покрывает виджетную часть карточки TREE-1 (этап 4 roadmap):

  * R-03 — популяция Treeview: значения колонок, теги сроков
    «expired»/«soon», точки типов (V4) в колонке #0;
  * R-04 — диф-обновления: правка/добавление/удаление точечные,
    неизменённые строки сохраняют iid (критерий 5.1 «выбор/изменение
    не пересоздаёт строки»); перемещения дают итоговый порядок;
  * R-05b — план порций (plan_display_count) и активация строки
    «Показать ещё»;
  * U13 — клавиатурная навигация: флаг key_nav доходит до контроллера;
  * стилизация обеих тем и снятие слушателя темы вместе с виджетом (LEAK-1).

Чистые функции (plan_diff, plan_display_count) тестируются без Tk ниже.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

tk = pytest.importorskip("tkinter")

try:
    _probe = tk.Tk()
    _probe.withdraw()
    _probe.destroy()
    TK_AVAILABLE = True
except tk.TclError:  # pragma: no cover — нет дисплея в CI без Xvfb
    TK_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not TK_AVAILABLE, reason="нет графического дисплея (нужен Xvfb)")

import registry_tree as rt
import ui_macos as ui


@pytest.fixture()
def root():
    root = tk.Tk()
    root.geometry("900x520+0+0")
    yield root
    try:
        root.update_idletasks()
        root.destroy()
    except tk.TclError:
        pass


COLUMNS = (
    ("name", "Контрагент", 260, "w", True),
    ("kind", "Тип", 70, "w", False),
    ("inn", "ИНН", 110, "w", False),
    ("id", "ID", 80, "w", False),
    ("address", "Адрес", 220, "w", True),
)


def _row(iid: str, name: str, *, tags=(), dot=None):
    return rt.RowSpec(iid=iid, values=(name, "Тип", "ИНН", iid, "Адрес"),
                      tags=tags, dot=dot)


@pytest.fixture()
def tree(root):
    selections: list = []
    activations: list = []
    widget = rt.RegistryTree(
        root,
        on_select=lambda iid, key_nav, user=False: selections.append((iid, key_nav, user)),
        on_activate=lambda iid: activations.append(iid),
        columns=COLUMNS,
    )
    widget.pack(fill="both", expand=True)
    root.update()
    widget.selections = selections
    widget.activations = activations
    return widget


def make_rows(count: int, start: int = 0) -> list:
    return [_row(f"p:{i}", f"ООО Компания-{i}") for i in range(start, start + count)]


# --------------------------------------------------------------------------- #
#  Чистые функции (без Tk)
# --------------------------------------------------------------------------- #
def test_plan_display_count_within_limit():
    assert rt.plan_display_count(10_000, rt.DISPLAY_SOFT_LIMIT) == (10_000, False)


def test_plan_display_count_over_limit():
    shown, has_more = rt.plan_display_count(10_001, rt.DISPLAY_SOFT_LIMIT)
    assert (shown, has_more) == (rt.DISPLAY_SOFT_LIMIT, True)


def test_plan_display_count_extended_portion():
    shown, has_more = rt.plan_display_count(50_000, rt.DISPLAY_SOFT_LIMIT * 2)
    assert (shown, has_more) == (20_000, True)


def test_plan_display_count_all_shown_caption_state():
    shown, has_more = rt.plan_display_count(50_000, 50_000)
    assert (shown, has_more) == (50_000, False)


def test_plan_display_count_small_registry():
    assert rt.plan_display_count(5, 1) == (5, False)


def test_plan_diff_no_change():
    plan = rt.plan_diff(["a", "b"], [_row("a", "1"), _row("b", "2")])
    assert plan.to_delete == [] and plan.ops == []


def test_plan_diff_add_edit_delete_keeps_prefix():
    plan = rt.plan_diff(
        ["a", "b", "c"],
        [_row("a", "A2"), _row("c", "3"), _row("d", "4")],
    )
    assert plan.to_delete == ["b"]
    assert (2, "d", "insert") in plan.ops
    # Строки «a» и «c» не трогаются (префикс совпадает после удаления «b»)
    assert all(iid not in {"a", "c"} for _pos, iid, _op in plan.ops)


def test_plan_diff_move_single_op():
    plan = rt.plan_diff(
        ["a", "b", "c"],
        [_row("c", "3"), _row("a", "1"), _row("b", "2")],
    )
    assert plan.to_delete == []
    assert plan.ops == [(0, "c", "move")], "минимальная операция: сдвиг «c» в начало"


def test_plan_diff_append_many_is_insert_only():
    current = [f"p{i}" for i in range(1000)]
    desired = [_row(f"p{i}", "") for i in range(1000)] + \
              [_row(f"n{i}", "") for i in range(200)]
    plan = rt.plan_diff(current, desired)
    assert len(plan.ops) == 200
    assert all(op[2] == "insert" for op in plan.ops)


# --------------------------------------------------------------------------- #
#  Живой виджет
# --------------------------------------------------------------------------- #
def test_populate_sets_values_tags_and_dot(tree):
    rows = [
        _row("p:1", "ООО Один"),
        _row("d:2", "Доп. № 2", tags=("expired",)),
        _row("d:3", "Доп. № 3", tags=("soon",), dot="#FEBC2E"),
    ]
    tree.populate(rows)
    tree.master.update()

    assert tree.order() == ("p:1", "d:2", "d:3")
    assert tree.tree.item("p:1", "values")[0] == "ООО Один"
    assert tree.tree.item("d:2", "tags") == ("expired",)
    assert tree.tree.item("d:3", "image") != "", "точка типа (V4) в колонке #0"
    assert tree.row("d:3").tags == ("soon",)


def test_diff_updates_values_without_recreating_rows(tree):
    rows = make_rows(5)
    tree.populate(rows)
    before = set(tree.order())

    changed = [_row("p:0", "ООО Изменённая")] + make_rows(4, start=1)
    changed.append(_row("p:5", "ООО Новая"))
    tree.apply_diff(changed)
    tree.master.update()

    assert "p:0" in tree.order() and "p:5" in tree.order() and len(tree.order()) == 6
    assert before <= set(tree.order()), "неизменённые строки сохранили iid"
    assert tree.tree.item("p:0", "values")[0] == "ООО Изменённая"
    assert all(iid in before for iid in ("p:1", "p:2", "p:3", "p:4"))


def test_diff_deletes_missing_rows(tree):
    tree.populate(make_rows(4))
    tree.apply_diff(make_rows(2))
    tree.master.update()

    assert tree.order() == ("p:0", "p:1")


def test_diff_moves_to_desired_order(tree):
    tree.populate(make_rows(3))
    tree.apply_diff([_row("p:2", ""), _row("p:0", ""), _row("p:1", "")])
    tree.master.update()

    assert tree.order() == ("p:2", "p:0", "p:1")


def test_show_more_row_reports_activation(tree):
    tree.populate([_row(rt.SHOW_MORE_IID, "Показать ещё", tags=("showmore",))])
    tree.tree.selection_set(rt.SHOW_MORE_IID)
    tree._handle_select()
    tree._handle_activate(type("E", (), {"y": None})())

    assert tree.activations == [rt.SHOW_MORE_IID]


def test_echo_events_are_set_semantics(tree):
    rows = make_rows(3)
    tree.populate(rows)
    tree.set_selection("p:1")
    tree.master.update()  # <<TreeviewSelect>> асинхронное: эхо доставляется очередью

    assert tree.selected_iid() == "p:1"
    # Эхо дошло с user=False: контроллер применит SET и состояние не изменит.
    assert tree.selections[-1] == ("p:1", False, False)

    tree.set_selection(None)
    tree.master.update()
    # Эхо снятия выделения — тоже user=False: «запомненный» выбор контроллера
    # снимается только явным действием пользователя (click-toggle).
    assert tree.selections[-1] == (None, False, False)

    # Клик-toggle: программно эмулируем press по выделенной строке.
    tree.set_selection("p:2")
    tree.master.update()
    _bx, _by, _bw, _bh = tree.tree.bbox("p:2")
    tree._handle_press(type("E", (), {"y": _by + _bh // 2})())
    tree.master.update()  # after_idle отрабатывает
    assert (None, False, True) in tree.selections, \
        "повторный клик по выделенной строке — явное снятие (user=True)"
    assert tree.selected_iid() is None


def test_set_selection_ignores_missing_iid(tree):
    tree.populate(make_rows(2))
    tree.set_selection("p:999")
    assert tree.selected_iid() is None
    tree.set_selection("p:1")
    assert tree.selected_iid() == "p:1"


def test_key_nav_flag_reaches_controller(tree):
    tree.populate(make_rows(2))
    tree.set_selection("p:0")
    tree.master.update()  # эхо синхронизации доставлено и обработано как SET
    tree._mark_key_nav()
    tree._handle_select()

    assert tree.selections[-1] == ("p:0", True, False), "U13: стрелки идут с флагом key_nav"


def test_theme_change_restyles_and_listener_removed(tree):
    light_bg = tree._style.lookup(tree._style_name, "background")
    ui.set_appearance_mode("dark")
    dark_bg = tree._style.lookup(tree._style_name, "background")
    ui.set_appearance_mode("light")

    assert light_bg != dark_bg, "R-03: цвета списка следуют теме"

    listeners_before = len(ui._LISTENERS)
    tree.destroy()
    tree.master.update()
    assert len(ui._LISTENERS) == listeners_before - 1, \
        "LEAK-1: слушатель темы снят вместе с виджетом"


def test_columns_switch_for_docs_view(tree):
    tree.populate(make_rows(1))
    tree.set_columns(rt.PARTY_COLUMNS if hasattr(rt, "PARTY_COLUMNS") else COLUMNS)
    # В реальном коде колонки документов задаёт контроллер (DOC_COLUMNS);
    # здесь проверяем сам механизм смены набора колонок.
    doc_columns = (
        ("doc", "Документ", 300, "w", True),
        ("date", "Дата", 100, "w", False),
        ("expiry", "Срок", 300, "w", True),
    )
    tree.set_columns(doc_columns)
    tree.populate([rt.RowSpec("d:1", ("Договор № 5", "01.10.2026", "действует до 31.12.2026"))])
    tree.master.update()

    assert tree.tree.item("d:1", "values") == (
        "Договор № 5", "01.10.2026", "действует до 31.12.2026")
