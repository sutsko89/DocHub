"""Интеграционные тесты сборки страниц ContractApp — с настоящим Tk.

Долг из README (закрыт в v2.12.0): сборщики виджетов в контроллерах
(catalog build_contract_page / build_additional_page / build_spec_page /
build_counterparty_page / build_registry_page / build_settings_page)
покрывались только smoke-тестом. Здесь приложение собирается целиком под
Xvfb (настоящий Tk + CustomTkinter) и проверяются ключевые привязки после
удаления фасадного слоя ContractApp:

  * все шесть страниц построены и зарегистрированы;
  * трассы переменных ведут в методы контроллеров напрямую
    (``_limit_var`` → DocumentPageController.format_limit: «1 500 000»
    и сумма прописью обновляются сразу после ``set``);
  * поиск по ИНН подставляет контрагента через DocumentContextController;
  * выбор контрагента переносится между страницами (п. 4.2);
  * таблица позиций спецификации создаётся (10 строк) и расширяется;
  * реестр рендерит карточки контрагентов;
  * статус ключа DaData показывается без подстановки значения (аудит 2.0.1);
  * кнопка футера переключается по страницам и привязана к контроллерам;
  * защита несохранённых данных (п. 4.4) видит правки формы.

Тесты требуют графический дисплей: без него модуль пропускается (Tk-сессию
для них даёт Xvfb — так же запускается и smoke-тест).
"""
from __future__ import annotations

import os
import time

import pytest

import document_builder
import main
import storage
from document_builder import F

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


@pytest.fixture(scope="module")
def app(tmp_path_factory):
    """Собирает ContractApp целиком с изолированным профилем пользователя.

    APPDATA указывает во временную папку: реестр (database.db) и настройки
    не трогают профиль разработчика. HOME тоже изолирован — папка
    «Документы» по умолчанию уходит во временный каталог. Модальные
    диалоги (ui.confirm/ui.alert) подменяются автостабами, чтобы случайная
    «грязная» страница не заблокировала прогон настоящим окном.
    """
    mp = pytest.MonkeyPatch()
    home = tmp_path_factory.mktemp("tk_home")
    mp.setenv("APPDATA", str(home / "appdata"))
    mp.setenv("HOME", str(home))
    mp.setattr(main.ui, "confirm", lambda *args, **kwargs: True)
    mp.setattr(main.ui, "alert", lambda *args, **kwargs: None)
    application = main.ContractApp()
    application.root.update()
    # Гасим стартовые таймеры ContractApp (after 250/600/900/1500 мс от
    # конструктора: онбординг «первый запуск» → страница «Контрагент»,
    # очистка .docx.tmp, автоочистка корзины, сверка реестра). Без этого
    # онбординг срабатывает внутри случайного теста, дёргает
    # _select_page("counterparty") и перезаписывает футер — флак
    # test_footer_save_button_follows_page (фикстура модульная, таймеры
    # просрочены к моменту любого update() в тестах). time.sleep не
    # возвращает управление раньше срока, поэтому после него все четыре
    # таймера гарантированно просрочены и снимаются одним update().
    time.sleep(1.8)
    application.root.update()
    # Онбординг уже дёрнул навигацию на «Контрагента» — возвращаем исходную
    # страницу «Договор», чтобы тесты начинались в прежнем детерминированном
    # состоянии (current_page == "contract"), но уже без стартовых таймеров.
    application._select_page("contract")
    application.root.update()
    yield application
    try:
        application.root.destroy()
    except Exception:
        pass
    mp.undo()


# --------------------------------------------------------------------------- #
#  Каркас окна и страниц
# --------------------------------------------------------------------------- #
def test_pages_built_and_registered(app):
    assert set(app.pages) == {
        "contract", "additional", "spec", "counterparty", "registry", "settings",
    }
    assert app.current_page == "contract"
    assert len(app.nav_items) == len(app.pages)
    assert set(app._form_baseline) == {
        "contract", "additional", "spec", "counterparty",
    }, "baseline всех форм (кроме реестра и настроек) снят после сборки"


def test_contract_page_widgets_created(app):
    assert app.contract_party_combo is not None
    for var in ("contract_party_var", "contract_number_var",
                "contract_day_var", "contract_month_var", "contract_search_var"):
        assert hasattr(app, var), f"нет {var}"
    assert app.contract_preview.get("1.0", "end").strip() != "", \
        "пустой предпросмотр договора при сборке страницы"


def test_counterparty_page_fields_created(app):
    for mode in ("ooo", "ip"):
        assert F.INN in app._cp_entries[mode], f"нет поля ИНН ({mode})"
    assert F.OGRN in app._cp_entries["ooo"], "нет поля ОГРН у юрлица"
    assert F.OGRNIP in app._cp_entries["ip"], "нет поля ОГРНИП у ИП"
    assert app._cp_type_var.get() == "Юр. лицо"


def test_settings_page_shows_key_status_without_value(app):
    status = app._dadata_key_status_var.get()
    assert status == "Ключ не задан.", \
        "статус без подстановки значения ключа (аудит 2.0.1)"


# --------------------------------------------------------------------------- #
#  Трассы переменных ведут в контроллеры напрямую (v2.12.0)
# --------------------------------------------------------------------------- #
def test_limit_trace_formats_and_updates_words(app):
    app._limit_var.set("1500000")
    assert app._limit_var.get() == "1 500 000", \
        "трасса _limit_var привязана к DocumentPageController.format_limit"
    words = app.limit_words_label.cget("text")
    assert "Сумма прописью:" in words
    assert "один миллион пятьсот тысяч рублей" in words

    app._limit_var.set("")
    assert app.limit_words_label.cget("text").endswith("—")


def test_term_trace_updates_words(app):
    app.term_var.set("45")
    assert app.words_label.cget("text") == "Срок прописью: сорок пять", \
        "трасса term_var привязана к DocumentPageController.update_words"
    app.term_var.set("")


def test_party_selection_transfers_across_pages(app):
    rec = app.store.add_counterparty("ooo", {
        F.INN: "7709001100",
        F.SHORT_NAME: "ООО «ТкВстрой»",
    }, save=False)
    display = app._doc_context_flow.party_display(rec)

    app.contract_party_var.set(display)
    assert app._current_party_id == rec["id"], \
        "трасса списка договора ведёт в DocumentContextController.on_party_selected"
    assert app.addl_party_var.get() == display, \
        "выбор перенесён на страницу доп. соглашения (п. 4.2)"

    app.contract_party_var.set("")


def test_inn_search_finds_party(app):
    rec = app.store.add_counterparty("ooo", {
        F.INN: "7709002200",
        F.SHORT_NAME: "ООО «ПоискИНН»",
    }, save=False)
    app.contract_search_var.set("7709002200")
    app._doc_page_flow.find_party_by_inn()

    assert app._current_party_id == rec["id"], "контрагент найден и подставлен"
    assert app.contract_search_var.get() == "", "поле поиска очищено"


def test_spec_table_created_with_rows(app):
    rows = app._doc_save_flow.spec_rows
    assert len(rows) == 10, "стартовое количество строк спецификации — 10"
    assert rows[0]["num"].cget("text") == "1"

    app._doc_save_flow.add_spec_row(name="Болт М12", qty="5", price="200", unit="шт.")
    rows = app._doc_save_flow.spec_rows
    assert len(rows) == 11
    assert rows[-1]["num"].cget("text") == "11", "нумерация строк актуальна"

    app._doc_save_flow.remove_spec_row()
    assert len(app._doc_save_flow.spec_rows) == 10


def test_registry_page_renders_party_rows(app):
    rec = app.store.add_counterparty("ooo", {
        F.INN: "7709003300",
        F.SHORT_NAME: "ООО «РеестрТк»",
    }, save=False)
    app._registry_flow.refresh()
    assert rec["id"] in app._registry_rows, "карточка контрагента отрисована в реестре"


def test_footer_save_button_follows_page(app):
    app._select_page("spec")
    assert app.current_page == "spec"
    app.root.update()
    assert app.save_button.cget("text") == "Сохранить спецификацию"

    app._select_page("registry")
    app.root.update()
    app._select_page("contract")
    app.root.update()
    assert app.save_button.cget("text") == "Сохранить договор"


def test_unsaved_data_guard_sees_changes(app):
    app.contract_number_var.set("777")
    assert app._doc_context_flow.is_form_dirty("contract"), \
        "правка номера делает страницу «грязной» (п. 4.4)"
    app._doc_context_flow.reset_form_baseline("contract")
    assert not app._doc_context_flow.is_form_dirty("contract")
    # вернуть пустую форму и зафиксировать baseline для последующих тестов
    app.contract_number_var.set("")
    app._doc_context_flow.reset_form_baseline("contract")


def test_dadata_guard_blocks_second_click(app):
    app._dadata_flow.busy = True
    app._dadata_flow.fill_from_inn()  # guard срабатывает до обращения к полям
    assert app._dadata_flow.busy is True
    assert app._dadata_flow.result_queue is None
    app._dadata_flow.busy = False


def test_on_dpichanged_recomputes_scale_and_geometry(app):
    """Перенос окна на монитор 150 % (п. 6.3): масштаб и геометрия обновлены.

    Обратный порядок относительно suggested rect: win_shell передаёт
    (x, y, ширина, высота). После проверки состояние откатывается —
    фикстура модульная, следующие тесты ждут прежний масштаб.
    """
    old_scale = app.scale
    old_geometry = app.root.geometry()
    try:
        app._on_dpichanged(144, (10, 20, 1210, 820))
        assert app.scale == 1.5
        assert main.ui.SCALE == 1.5
        geo = app.root.geometry()
        # Позиция из подсказанного прямоугольника обязана совпасть точно;
        # размер может подрезаться менеджером окружения (Xvfb без WM
        # подрезает запрошенный размер внутренней логикой CustomTkinter) —
        # важен сам факт применения новой геометрии.
        assert geo.endswith("+10+20"), geo
        assert geo != old_geometry
    finally:
        app.scale = old_scale
        main.ui.set_scale(old_scale)
        app.root.geometry(old_geometry)
        app.root.update()


def test_store_uses_isolated_profile(app):
    import storage as st

    assert os.environ["APPDATA"] in st.db_path(), \
        "реестр живёт во временном профиле, а не в данных разработчика"
    assert app.store.active_documents() == [], "чистый реестр на старте"
