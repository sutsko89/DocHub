r"""Каркасные тесты Qt-приложения (этап M2 миграции Tk → PySide6).

Проверяют контракты этапа M2 плана миграции на offscreen-платформе
(создание QApplication запрещено в тестах напрямую — используется fixture
``qapp`` pytest-qt; APPDATA изолируется стандартной fixture ``appdata``):

* окно со стандартной рамкой строится: заголовок совпадает с контрактом
  single_instance (FindWindow в activate_existing_window ищет по нему);
* панель разделов повторяет ключи Tk NAV_PAGES (дрейф навигации заметен
  сразу), страницы-плейсхолдеры сообщают «в разработке» и этап;
* тема переключается, пишется в settings.json под прежним ключом и
  переживает пересоздание каркаса; неизвестное значение — дефолт;
  слушатели темы снимаются по владельцу (LEAK-1/D8);
* DataStore принадлежит каркасу; счётчики статуса — из status_counts;
  битая база даёт отложенный алерт, а не падение (контракт Tk);
* QTimer-инфраструктура: таймер опроса тикает, останов останавливает;
  закрытие окна останавливает всё и пишет в журнал;
* perf-каналы startup.built/startup.ready пишутся при
  EWDOCHUB_PERF_LOG=1;
* диалоги qt_dialogs: деструктивное подтверждение — Enter = «Отмена»
  (U7, UX-A), choice возвращает ключ или None при закрытии.

Защита от второго экземпляра проверяется контрактом (заголовок/мьютексы):
сами Win32-мьютексы на не-Windows недоступны по конструкции
single_instance.py, реальная блокировка — шаг Windows-runbook (M2).
"""

from __future__ import annotations

import os
import sys
from types import SimpleNamespace

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest

import app_common
import applog
import perflog
import qt_dialogs
import qt_theme
import single_instance
import storage
from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import QLabel, QMessageBox

import main_qt  # noqa: E402 — после установки QT_QPA_PLATFORM

OOO_FIELDS = {
    "ИНН": "7701234567", "КПП": "770101001", "ОГРН": "1037739010221",
    "Полное название компании (Покупь)": "Общество с ограниченной "
    "ответственностью Ромашка",
    "Сокращенное название компании": "ООО Ромашка",
    "ФИО директора": "Иванов Иван Иванович", "Основание действий": "Устав",
    "Юридический адрес": "г. Москва, ул. Ленина, д. 1",
    "Фактический адрес": "г. Москва, ул. Ленина, д. 1",
    "Email": "info@romashka.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012",
    "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "ПАО Сбербанк", "БИК": "044525225",
}


@pytest.fixture()
def controller(qapp, appdata):
    """Каркас на изолированном профиле; таймеры останавливаются после теста."""
    ctrl = main_qt.AppController(qapp)
    yield ctrl
    ctrl._stop_timers()


# --------------------------------------------------------------------------- #
#  Окно, навигация, статус
# --------------------------------------------------------------------------- #
def test_window_title_matches_single_instance_contract(controller):
    # activate_existing_window ищет окно по этому заголовку (FindWindowW)
    assert controller.window.windowTitle() == single_instance.WINDOW_TITLE


def test_nav_sections_mirror_tk_pages(controller):
    """Ключи разделов Qt-панели — те же, что у Tk NAV_PAGES (без «Settings drift»)."""
    import main  # Tk-каркас: только чтение констант, окна не создаёт

    tk_keys = [key for key, _title, _icon in main.NAV_PAGES]
    qt_keys = [key for key, _title, _stage in main_qt.NAV_PAGES]
    assert qt_keys == tk_keys
    rows = [
        controller.nav.item(row).data(Qt.ItemDataRole.UserRole)
        for row in range(controller.nav.count())
    ]
    assert rows == tk_keys


def test_default_page_is_spec(controller):
    assert controller.current_page == main_qt.DEFAULT_PAGE == "spec"


def test_placeholder_pages_announce_stage(controller):
    """Страницы, ещё не ставшие живыми, сообщают «в разработке» и этап.

    Этап M6: «Спецификация» — живая страница CP1 (M3), «Добавить
    контрагента» — живая страница M4 (CounterpartyPageController),
    «Реестр» — живая страница M5 (RegistryPageController), «Договор
    поставки» и «Доп. соглашение» — живые страницы M6-a
    (DocumentPageQt); проверяются только будущие разделы. Живость
    закреплена в тестах своих этапов.
    """
    placeholder_stages = (
        (key, title, stage) for key, title, stage in main_qt.NAV_PAGES
        if key not in ("spec", "counterparty", "registry", "contract",
                       "additional")  # живые M3/M4/M5/M6-a
    )
    for key, title, stage in placeholder_stages:
        page = controller.pages[key]
        labels = page.findChildren(QLabel)
        texts = [label.text() for label in labels]
        assert any("в разработке" in text for text in texts), key
        assert any(stage in text for text in texts), key
        assert any(title in text for text in texts), key


def test_select_page_switches_stack_and_nav(controller):
    controller.select_page("registry")
    assert controller.current_page == "registry"
    assert controller.stack.currentWidget() is controller.pages["registry"]
    current = controller.nav.currentItem()
    assert current.data(Qt.ItemDataRole.UserRole) == "registry"
    controller.select_page("nonexistent")  # неизвестный ключ — без изменения
    assert controller.current_page == "registry"


def test_min_window_size_matches_tk_constants(controller):
    # Минимальный размер каркаса повторяет Tk-версию (BASE_*/MIN_* в main_qt)
    assert controller.window.minimumSize() == QSize(
        main_qt.MIN_WINDOW_WIDTH, main_qt.MIN_WINDOW_HEIGHT,
    )


def test_status_counts_reflect_store(controller):
    assert "Контрагентов: 0" in controller.status_label.text()
    controller.store.add_counterparty("ooo", dict(OOO_FIELDS))
    controller.refresh_status()
    assert "Контрагентов: 1" in controller.status_label.text()
    assert "Документов: 0" in controller.status_label.text()


# --------------------------------------------------------------------------- #
#  Профиль и настройки (совместимость с Tk-сборкой)
# --------------------------------------------------------------------------- #
def test_first_start_writes_only_database(controller, appdata):
    # Контракт первой установки (2.16.0): settings.json появляется только
    # при сохранении настроек, а не при старте каркаса.
    profile = appdata / "DogovorGenerator"
    names = {p.name for p in profile.iterdir()}
    assert "database.db" in names
    assert "settings.json" not in names
    assert "dadata_key.txt" not in names


def test_theme_toggle_persists_and_survives_restart(qapp, appdata):
    ctrl = main_qt.AppController(qapp)
    assert qt_theme.current_mode() == "light"
    assert ctrl.set_theme("dark") == "dark"
    assert app_common.load_settings().get("theme") == "dark"

    # «Перезапуск»: новый каркас на том же профиле читает сохранённую тему.
    ctrl2 = main_qt.AppController(qapp)
    assert qt_theme.current_mode() == "dark"
    assert ctrl2.set_theme("light") == "light"
    assert app_common.load_settings().get("theme") == "light"
    ctrl2._stop_timers()


def test_theme_unknown_value_falls_back_to_light(qapp, appdata):
    app_common.write_settings({"theme": "purple"})
    ctrl = main_qt.AppController(qapp)
    assert qt_theme.current_mode() == "light"
    ctrl._stop_timers()


def test_theme_listener_notified_and_owner_unsubscribes(controller, qapp, qtbot):
    seen: list[str] = []
    controller.register_theme_listener(seen.append, owner=controller.toast)
    controller.set_theme("dark")
    controller.set_theme("light")
    assert seen == ["dark", "light"]
    controller.unregister_theme_listener(seen.append)
    controller.set_theme("dark")
    assert seen == ["dark", "light"]  # после ручной отписки не получает событий

    # LEAK-1/D8: подписка с владельцем снимается сама при уничтожении.
    auto: list[str] = []
    controller.register_theme_listener(auto.append, owner=controller.toast)
    controller.toast.deleteLater()
    qtbot.wait(50)  # цикл событий обрабатывает DeferredDelete
    controller.set_theme("light")
    assert auto == []  # мёртвых замыканий нет


def test_theme_write_failure_alerts_but_stays_applied(controller, qapp, monkeypatch):
    shown: list[tuple] = []
    monkeypatch.setattr(qt_dialogs, "alert", lambda *a, **k: shown.append(a))
    # Запись настроек не удалась — write_settings возвращает False.
    monkeypatch.setattr(app_common, "write_settings", lambda settings: False)
    assert controller.set_theme("dark") == "dark"
    assert qt_theme.current_mode() == "dark"  # тема в сеансе применена
    assert shown and "Настройки не сохранены" in shown[0][1]


# --------------------------------------------------------------------------- #
#  DataStore: владение каркасом, битая база
# --------------------------------------------------------------------------- #
def test_store_load_warning_shows_deferred_alert(qapp, appdata, monkeypatch, qtbot):
    db = storage.db_path()
    os.makedirs(os.path.dirname(db), exist_ok=True)
    with open(db, "wb") as fh:
        fh.write(b"this is not a sqlite database")

    shown: list[tuple] = []
    monkeypatch.setattr(qt_dialogs, "alert", lambda *a, **k: shown.append(a))
    ctrl = main_qt.AppController(qapp)
    assert ctrl.store.load_warning  # каркас стартует с пустым реестром
    qtbot.waitUntil(lambda: bool(shown), timeout=3000)
    assert shown[0][1] == "Реестр загружен не полностью"
    ctrl._stop_timers()


# --------------------------------------------------------------------------- #
#  QTimer-инфраструктура и закрытие
# --------------------------------------------------------------------------- #
def test_poll_timer_ticks_and_stop_poll_timer_stops(controller, qtbot):
    ticks: list[int] = []
    controller.add_poll_timer("probe", 20, lambda: ticks.append(1))
    qtbot.waitUntil(lambda: len(ticks) >= 2, timeout=3000)
    controller.stop_poll_timer("probe")
    frozen = len(ticks)
    qtbot.wait(150)
    assert len(ticks) == frozen


def test_add_poll_timer_replaces_by_name(controller):
    first = controller.add_poll_timer("dup", 1000, lambda: None)
    second = controller.add_poll_timer("dup", 2000, lambda: None)
    assert first is not second
    assert first.objectName() == second.objectName()
    assert controller._timers["dup"] is second  # старый замещён, утечки нет
    assert not first.isActive()


def test_close_stops_timers_and_logs(controller, qapp, qtbot, monkeypatch):
    controller.window.show()
    qapp.processEvents()
    recorded: list[tuple] = []
    # applog — глобальный обработчик на процесс (путь фиксируется первым
    # профилем в прогоне); файл app.log проверяет smoke_qt в свежем
    # процессе, здесь — контракт записи закрытия через логгер каркаса.
    monkeypatch.setattr(
        controller, "log",
        SimpleNamespace(info=lambda *a, **k: recorded.append(a),
                        exception=lambda *a, **k: recorded.append(a)),
    )
    ticks: list[int] = []
    controller.add_poll_timer("probe", 20, lambda: ticks.append(1))
    qtbot.waitUntil(lambda: len(ticks) >= 1, timeout=3000)
    controller.window.close()
    qapp.processEvents()
    assert controller._timers == {}
    frozen = len(ticks)
    qtbot.wait(100)
    assert len(ticks) == frozen
    assert any("Каркас Qt закрыт" in str(a) for a in recorded)


def test_call_later_runs_on_event_loop(controller, qapp, qtbot):
    fired: list[int] = []
    controller.call_later(10, lambda: fired.append(1))
    qtbot.waitUntil(lambda: bool(fired), timeout=3000)


# --------------------------------------------------------------------------- #
#  Perf-каналы старта (E-4)
# --------------------------------------------------------------------------- #
def test_startup_perf_events_written(qapp, appdata, monkeypatch):
    monkeypatch.setenv(perflog.ENV_FLAG, "1")
    perflog.reset_for_tests()
    try:
        ctrl = main_qt.AppController(qapp)
        ctrl._emit_startup_perf()
        qapp.processEvents()  # singleShot(0) — ready на первом проходе цикла
        perflog.flush()
        with open(perflog.perf_path(), encoding="utf-8") as fh:
            content = fh.read()
        assert "startup.built" in content
        assert "startup.ready" in content
        assert "event" in content
    finally:
        perflog.reset_for_tests()


def test_perf_disabled_by_default(qapp, appdata, monkeypatch):
    monkeypatch.delenv(perflog.ENV_FLAG, raising=False)
    perflog.reset_for_tests()
    try:
        ctrl = main_qt.AppController(qapp)
        ctrl._emit_startup_perf()
        qapp.processEvents()
        perflog.flush()
        assert not os.path.exists(perflog.perf_path())
    finally:
        perflog.reset_for_tests()


# --------------------------------------------------------------------------- #
#  Тост
# --------------------------------------------------------------------------- #
def test_toast_shows_and_hides(controller, qapp):
    controller.window.show()
    qapp.processEvents()
    controller.notify("Сохранено")
    assert controller.toast.isVisible()
    assert not controller.toast_open.isVisible()
    assert not controller.toast_folder.isVisible()
    controller._hide_toast()
    assert not controller.toast.isVisible()


def test_toast_with_path_offers_actions(controller, qapp):
    controller.window.show()
    qapp.processEvents()
    controller.notify("Документ сохранён", path="C:/tmp/example.docx")
    assert controller.toast.isVisible()
    assert controller.toast_open.isVisible()
    assert controller.toast_folder.isVisible()
    # Повторный вызов без пути прячет кнопки и перезапускает таймер.
    controller.notify("Просто сообщение")
    assert not controller.toast_open.isVisible()
    assert controller._toast_path is None


# --------------------------------------------------------------------------- #
#  Диалоги qt_dialogs (замены Alert/Sheet)
# --------------------------------------------------------------------------- #
@pytest.fixture()
def stub_message_exec(monkeypatch):
    """Подменяет exec() и clickedButton(): «нажатие» — дефолтная кнопка.

    QMessageBox.exec — блокирующий вызов C++; в тестах Enter «нажимает»
    то, что задано setDefaultButton, что и проверяет семантику.
    """
    shown: list[QMessageBox] = []

    def _fake_exec(self):
        shown.append(self)
        return 0

    def _fake_clicked(self):
        return self.defaultButton()

    monkeypatch.setattr(QMessageBox, "exec", _fake_exec)
    monkeypatch.setattr(QMessageBox, "clickedButton", _fake_clicked)
    return shown


def test_alert_icon_by_kind(stub_message_exec):
    qt_dialogs.alert(None, "Заголовок", "Сообщение", "warning", ok_text="Понятно")
    box = stub_message_exec[0]
    assert box.icon() == QMessageBox.Icon.Warning
    assert box.text() == "Сообщение"
    assert box.defaultButton().text() == "Понятно"


def test_confirm_non_destructive_enter_confirms(stub_message_exec):
    assert qt_dialogs.confirm(None, "Вопрос", "Сохранить?", destructive=False) is True
    assert stub_message_exec[0].defaultButton().text() == "ОК"


def test_confirm_destructive_enter_cancels(stub_message_exec):
    # U7, UX-A: подтверждение потери данных не срабатывает от Enter
    assert qt_dialogs.confirm(
        None, "Удаление", "Удалить запись?", destructive=True,
    ) is False
    assert stub_message_exec[0].defaultButton().text() == "Отмена"


def test_choice_returns_primary_key(stub_message_exec):
    result = qt_dialogs.choice(
        None, "Выбор", "Что сделать?",
        [("bind", "Привязать", False), ("create", "Создать", True)],
    )
    assert result == "create"  # главная кнопка — дефолт для Enter


def test_choice_close_returns_none(monkeypatch):
    monkeypatch.setattr(QMessageBox, "exec", lambda self: 0)
    monkeypatch.setattr(QMessageBox, "clickedButton", lambda self: None)
    assert qt_dialogs.choice(
        None, "Выбор", "Что сделать?", [("a", "А", True), ("b", "Б", False)],
    ) is None


# --------------------------------------------------------------------------- #
#  Контракт второго экземпляра (реальная блокировка — Windows-runbook)
# --------------------------------------------------------------------------- #
def test_second_instance_contract_constants():
    # Заголовок Qt-окна — тот же, по которому FindWindowW активирует
    # уже запущенную копию; мьютексы — те же имена (Global + SID).
    assert main_qt.single_instance.WINDOW_TITLE == single_instance.WINDOW_TITLE
    assert single_instance.mutex_name().startswith("Global\\EWDocHub_SingleInstance")
    assert single_instance.INSTALLER_MUTEX_NAME == "Global\\EWDocHub_InstallerMutex"


def test_main_no_lock_returns_before_qapplication(monkeypatch):
    """main(): занятый мьютекс — выход до создания интерфейса (как в Tk)."""

    class _BusyLock:
        def acquire(self):
            return False

        def release(self):
            raise AssertionError("release не вызывается при отказе acquire")

    created: list[str] = []
    monkeypatch.setattr(main_qt.single_instance, "InstanceLock", _BusyLock)
    monkeypatch.setattr(
        main_qt.single_instance, "activate_existing_window", lambda: created.append("activate"),
    )
    monkeypatch.setattr(
        main_qt.single_instance, "show_already_running", lambda: created.append("message"),
    )
    assert main_qt.main() == 0
    assert created == ["activate", "message"]


# --------------------------------------------------------------------------- #
#  Тост-действия: честный отказ без файла (контракт main.open_file)
# --------------------------------------------------------------------------- #
def test_open_path_missing_file_alerts(controller, monkeypatch):
    shown: list[tuple] = []
    monkeypatch.setattr(qt_dialogs, "alert", lambda *a, **k: shown.append(a))
    assert controller._open_path("C:/no/such/file.docx") is False
    assert shown and shown[0][1] == "Файл не найден"
