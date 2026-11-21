"""UX-A (этап 2, параллельная ветка) — дешёвая UX-ветка форм.

Карточка UX-A: V1 контраст третичного текста (авто-проверка пар палитры),
U4 подсказки формата в полях, U5 прокрутка к ошибочному полю (хелпер —
Tk-часть в test_ux_a_tk.py), U6 один алерт на все ошибки позиций,
U7 Enter как «ОК» в нейтральных подтверждениях (деструктивные не
переводятся на «Enter = ОК»), U14 подсказка Ctrl+S, U3 поиск по ИНН
на доп. соглашении и спецификации.

Чистые части проверяются без GUI; живые виджеты — в test_ux_a_tk.py
(требуют дисплей/Xvfb).
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import ui_macos as ui
from counterparty_controller import FORMAT_HINTS
from document_builder import F
from document_save_controller import POSITION_ISSUE_LIMIT, position_issues_message
from settings_controller import HOTKEY_HINTS

REPO_ROOT = Path(__file__).resolve().parent.parent


# --------------------------------------------------------------------------- #
#  V1: контраст третичного текста — авто-проверка пар палитры
# --------------------------------------------------------------------------- #
def test_contrast_ratio_reference_values():
    # Чёрный на белом — максимум по шкале WCAG.
    assert ui.contrast_ratio("#000000", "#FFFFFF") == pytest.approx(21.0, abs=0.01)
    # Один и тот же цвет — минимальное соотношение 1:1.
    assert ui.contrast_ratio("#96969C", "#96969C") == pytest.approx(1.0)
    # Порядок аргументов не влияет.
    assert ui.contrast_ratio("#FFFFFF", "#000000") == pytest.approx(21.0, abs=0.01)
    # Подтверждённая аудитом пара: #96969C на белом ≈ 2.96:1.
    assert ui.contrast_ratio("#96969C", "#FFFFFF") == pytest.approx(2.96, abs=0.05)


def test_v1_tertiary_color_is_the_confirmed_value():
    # UX-аудит подтвердил #96969C как значение третичного текста светлой темы.
    assert ui.COLORS["text_tertiary"][0].lower() == "#96969c"


def test_v1_contrast_baseline_has_no_violations():
    """Контраст пар палитры проверен автоматически (критерий этапа 2, V1)."""
    problems = ui.contrast_baseline_violations()
    assert problems == [], f"пары палитры ниже порога: {problems}"


# --------------------------------------------------------------------------- #
#  U4: подсказки формата в полях
# --------------------------------------------------------------------------- #
def test_u4_format_hints_cover_strict_fields():
    # 5 полей с жёстким форматом (UX-аудит T7): ИНН, КПП, ОГРН/ОГРНИП,
    # расчётный и корреспондентский счёт, БИК.
    for key in (F.INN, F.KPP, F.OGRN, F.OGRNIP, F.ACCOUNT, F.CORR_ACCOUNT, F.BIK):
        assert key in FORMAT_HINTS, f"нет подсказки формата для {key}"
        assert FORMAT_HINTS[key], f"пустая подсказка для {key}"


def test_u4_format_hint_texts_match_field_formats():
    assert FORMAT_HINTS[F.INN] == "10 или 12 цифр"
    assert FORMAT_HINTS[F.KPP] == "9 цифр"
    assert FORMAT_HINTS[F.OGRN] == "13 цифр"
    assert FORMAT_HINTS[F.OGRNIP] == "15 цифр"
    assert FORMAT_HINTS[F.ACCOUNT] == "20 цифр"
    assert FORMAT_HINTS[F.CORR_ACCOUNT] == "20 цифр"
    assert FORMAT_HINTS[F.BIK] == "9 цифр"


def test_u4_hints_are_not_values_of_entry():
    """Подсказка — надпись-подложка, а не содержимое Entry (риск UX-A).

    В коде подсказка ставится только через set_placeholder/placeholder;
    внутреннее значение Entry подсказку не получает, поэтому get()
    пустого поля всегда возвращает пустую строку (проверка живых
    виджетов — в test_ux_a_tk.py). Здесь фиксируем контракт фабрики:
    подсказка передаётся отдельным аргументом placeholder.
    """
    source = (REPO_ROOT / "ui_macos.py").read_text(encoding="utf-8")
    assert "def set_placeholder" in source
    assert "placeholder: str | None = None" in source
    cp_source = (REPO_ROOT / "counterparty_controller.py").read_text(encoding="utf-8")
    # Подсказка не вставляется в значение поля — только параметром.
    assert "insert(0, FORMAT_HINTS" not in cp_source
    assert "set(FORMAT_HINTS" not in cp_source


# --------------------------------------------------------------------------- #
#  U6: один алерт на все ошибки позиций
# --------------------------------------------------------------------------- #
def test_u6_no_issues_returns_empty_string():
    items = [
        {"name": "Кабель", "qty": 5, "price": 120},
        {"name": "Крепёж", "qty": 1, "price": 0.5},
    ]
    assert position_issues_message(items) == ""


def test_u6_qty_and_price_issues_in_one_message():
    items = [
        {"name": "Кабель", "qty": 0, "price": 120},
        {"name": "Крепёж", "qty": 3, "price": 0},
        {"name": "Хомуты", "qty": -1, "price": -2},
    ]
    text = position_issues_message(items)
    assert "«Кабель» — количество" in text
    assert "«Крепёж» — цена" in text
    assert "«Хомуты» — количество" in text and "«Хомуты» — цена" in text
    # Раньше было три последовательных алерта (по одному на позицию) —
    # теперь одна сводка, а вызов алерта один (проверяется в save_spec).
    assert text.count("\n") == 3


def test_u6_nameless_item_is_named_honestly():
    text = position_issues_message([{"name": "  ", "qty": 0, "price": 0}])
    assert "позиция без названия" in text


def test_u6_long_price_list_is_truncated_with_counter():
    # 10 позиций × 2 ошибки = 20 строк; показываются первые 8, остальное —
    # счётчиком «…и ещё», чтобы алерт не занимал весь экран.
    items = [{"name": f"Позиция {i}", "qty": 0, "price": 0}
             for i in range(POSITION_ISSUE_LIMIT + 2)]
    text = position_issues_message(items)
    assert text.count("количество") == POSITION_ISSUE_LIMIT // 2
    assert "…и ещё 12" in text


def test_u6_empty_positions_list_is_not_an_issue():
    assert position_issues_message([]) == ""


# --------------------------------------------------------------------------- #
#  U7: Enter = «ОК» только в нейтральных подтверждениях
# --------------------------------------------------------------------------- #
CONFIRM_FILES = (
    "lifecycle_controller.py",
    "settings_controller.py",
    "document_context_controller.py",
    "updater_controller.py",
    "widgets.py",
    "main.py",
)


def _confirm_calls(path: Path) -> list[ast.Call]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    calls: list[ast.Call] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if isinstance(func, ast.Attribute) and func.attr == "confirm":
            # ui.confirm(...) — не self.app.ui.confirm и не что-то другое;
            # в этом репозитории confirm вызывается только из модуля ui_macos.
            calls.append(node)
    return calls


def _call_kwargs(call: ast.Call) -> dict[str, ast.expr]:
    return {kw.arg: kw.value for kw in call.keywords if kw.arg}


def _constant(node: ast.expr) -> str | None:
    return node.value if isinstance(node, ast.Constant) else None


# Титулы подтверждений, ведущих к потере данных: Enter обязан означать
# «Отмена» (destructive=True). Нейтральные (импорт-привязка, повтор
# номера) — Enter = «ОК».
DESTRUCTIVE_TITLES = {
    "Удалить документ",          # в корзину
    "Удалить контрагента",       # в корзину
    "Подтвердите удаление",      # массовое перемещение в корзину
    "Удалить навсегда",          # purge документа и контрагента
    "Очистить корзину",          # purge_all_trash
    "Удалить ключ DaData",       # сброс ключа
    "Несохранённые данные",      # закрытие программы с потерей правок
    "Несохранённые изменения",   # уход со страницы с потерей правок
    "Таблица позиций",           # очистка таблицы позиций
    "Обновление скачано",        # установка обновления закрывает программу
}


@pytest.mark.parametrize("filename", CONFIRM_FILES)
def test_u7_destructive_confirms_are_flagged(filename):
    found = 0
    for call in _confirm_calls(REPO_ROOT / filename):
        kwargs = _call_kwargs(call)
        title = None
        args = call.args
        # титул — второй позиционный аргумент (parent, title, message, ...)
        if len(args) >= 2:
            title = _constant(args[1])
        if title in DESTRUCTIVE_TITLES:
            found += 1
            assert "destructive" in kwargs, (
                f"{filename}: подтверждение «{title}» без destructive=True — "
                "Enter не должен нажимать деструктивную главную кнопку"
            )
    assert found > 0, f"{filename}: не найдено ни одного деструктивного подтверждения"


def test_u7_neutral_confirms_stay_neutral():
    """Привязка при импорте и повтор номера — нейтральные: без флага.

    import_controller и document_save_controller не имеют деструктивных
    подтверждений — там флаг не должен появляться вовсе.
    """
    neutral_titles = {"Похожее название уже есть", "Похожий контрагент уже есть",
                      "Повтор номера"}
    checked = 0
    for filename in CONFIRM_FILES + ("import_controller.py",
                                     "document_save_controller.py"):
        for call in _confirm_calls(REPO_ROOT / filename):
            kwargs = _call_kwargs(call)
            args = call.args
            title = _constant(args[1]) if len(args) >= 2 else None
            if title in neutral_titles:
                checked += 1
                assert "destructive" not in kwargs, (
                    f"{filename}: «{title}» — нейтральное подтверждение, "
                    "Enter должен нажимать главную кнопку"
                )
    assert checked == 3, f"найдено {checked} из 3 нейтральных подтверждений"


def test_u7_alert_return_binding_depends_on_destructive():
    """Чистая часть контракта U7: у Alert флаг читается в привязке Return.

    Живое поведение (нажатие Return меняет _result) — в test_ux_a_tk.py;
    здесь фиксируем, что привязка Return ставится только через флаг:
    деструктивный confirm передаёт destructive=True в конструктор Alert.
    """
    source = (REPO_ROOT / "ui_macos.py").read_text(encoding="utf-8")
    assert "self._destructive = bool(destructive)" in source
    assert "elif cancel_text:" in source  # нейтральный путь: Return → _ok


# --------------------------------------------------------------------------- #
#  U14: подсказка Ctrl+S
# --------------------------------------------------------------------------- #
def test_u14_hotkey_hints_card_lists_save_and_navigation():
    joined = {keys for keys, _description in HOTKEY_HINTS}
    assert any("Ctrl+S" in keys for keys in joined), "Ctrl+S отсутствует в подсказках"
    assert any("Ctrl+N" in keys for keys in joined), "Ctrl+N отсутствует в подсказках"
    assert all(description.startswith("—") for _, description in HOTKEY_HINTS), \
        "описания должны начинаться с тире (формат «клавиша — действие»)"


def test_u14_footer_button_carries_ctrl_s_suffix():
    """Кнопка футера получает суффикс «(Ctrl+S)» на каждой странице форм."""
    source = (REPO_ROOT / "main.py").read_text(encoding="utf-8")
    assert '"{label} (Ctrl+S)"' in source, "суффикс Ctrl+S не добавляется в _update_footer"
    assert "Сохранить договор (Ctrl+S)" in source, "исходная кнопка без суффикса"


# --------------------------------------------------------------------------- #
#  U3: поиск по ИНН на доп. соглашении и спецификации
# --------------------------------------------------------------------------- #
def test_u3_search_row_is_built_on_all_document_pages():
    page_source = (REPO_ROOT / "document_page_controller.py").read_text(encoding="utf-8")
    save_source = (REPO_ROOT / "document_save_controller.py").read_text(encoding="utf-8")
    assert 'build_party_search_row(card, "contract_search_var")' in page_source
    assert 'build_party_search_row(card, "addl_search_var")' in page_source, \
        "страница «Доп. соглашение» без поиска по ИНН"
    assert 'build_party_search_row(card, "spec_search_var")' in save_source, \
        "страница «Спецификация» без поиска по ИНН"


def test_u3_search_var_keeps_legacy_name_on_contract_page():
    """``contract_search_var`` сохранён — его читают smoke-тест и страницы."""
    source = (REPO_ROOT / "document_page_controller.py").read_text(encoding="utf-8")
    assert "contract_search_var" in source
