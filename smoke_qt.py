"""Smoke Qt-сборки EW.DocHub (этапы M2–M5 миграции Tk → PySide6).

Проверяет критерии этапов M2 (каркас: старт → построение окна и разделов →
журналы → perf-события старта → тема → выход), M3 (контрольная точка CP1:
страница «Спецификация» — вставка блока из Excel с лимитом 500, итог,
предпросмотр из шаблона без записи на диск, сохранение через
services.save_document, защита несохранённых при закрытии), M4 (карточка
контрагента: состав полей ООО/ИП, валидация с подсветкой, «Заполнить по
ИНН» без ключа, окно правки карточки, Ctrl+S, защита страницы) и M5
(реестр TREE-1: модель строк = данным, поиск по ИНН на лету, панель
действий, корзина с восстановлением, drill «Документы — <имя>»). Запуск:

    python smoke_qt.py            # Windows — нативная рамка
    xvfb-run -a python smoke_qt.py  # Linux — Xvfb (или offscreen автоматически)

Профиль изолируется (APPDATA во временной папке), сеть отключается
(EWDOCHUB_NO_UPDATE_CHECK=1), perf-лог включается для проверки событий
startup.built/startup.ready. Формат вывода — как у smoke_test.py:
построчные «ok:/FAIL:», сводка неудач в конце, код возврата — числу
неудач (до 100).
"""
from __future__ import annotations

import json
import os
import sqlite3
import sys
import tempfile
from datetime import datetime, timedelta

# Изолированный профиль и заглушки окружения — до импортов приложения.
os.environ["APPDATA"] = os.path.join(tempfile.gettempdir(), "dogovor_smoke_qt_appdata")
os.environ["EWDOCHUB_NO_UPDATE_CHECK"] = "1"
# M2-критерий: EWDOCHUB_PERF_LOG=1 пишет события старта (startup.built/ready).
os.environ["EWDOCHUB_PERF_LOG"] = "1"
if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Windows-консоль может быть без полной поддержки кириллицы — как в smoke_test.
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import applog
import app_common
import main_qt as mq
import perflog
import qt_theme
import single_instance
import storage
from document_builder import F
from PySide6.QtCore import QCoreApplication, QEvent, QTimer
from PySide6.QtWidgets import QApplication, QDialog, QLabel, QLineEdit

FAILURES: list[str] = []
_TOTAL_CHECKS = 0
ALERTS: list[tuple] = []

# Чистый профиль на каждый запуск: база, журналы и настройки могли остаться
# от прошлого прогона (тема «перезапуска» проверяется на свежем файле).
for _path in (
    storage.db_path(),
    applog.log_path(),
    perflog.perf_path(),
    app_common.settings_path(),
):
    try:
        os.remove(_path)
    except OSError:
        pass


def check(condition, message):
    global _TOTAL_CHECKS
    _TOTAL_CHECKS += 1
    if condition:
        print(f"ok: {message}")
    else:
        FAILURES.append(message)
        print(f"FAIL: {message}")


# ---------- Журналы: тот же setup, что в main_qt.main() ----------
applog.setup_logging(f"{app_common.APP_VERSION}-{mq.UI_FLAVOR}")
applog.enable_crash_dump()  # путь crash.log, как включает точка входа Qt

# Модальные диалоги в smoke никто не закроет — подменяем (маршруты main_qt).
# choice: стартовые задачи (START-1) на свежем профиле показывают диалог
# сроков доп. соглашений — подмена не даёт смоку зависнуть; запись — для U9.
CHOICES: list[tuple] = []
_CHOICE_RESULT = None
mq.qt_dialogs.alert = lambda *args, **kwargs: ALERTS.append(args)
mq.qt_dialogs.confirm = lambda *args, **kwargs: True


def _fake_choice(*args, **kwargs):
    CHOICES.append(args)
    return _CHOICE_RESULT


mq.qt_dialogs.choice = _fake_choice


class _SyncFactory:
    """Фабрика потока, исполняющая воркер синхронно (подмена для смоука).

    Как RecordingFactory в test_import_controller: воркер отрабатывает
    внутри start(), очередь опрашивается tick() вручную — прогон
    детерминирован без реальных потоков.
    """

    def __call__(self, target=None, args=(), daemon=None, **_kw):
        self._target = target
        self._args = args
        return self

    def start(self) -> None:
        self._target(*self._args)

qt_app = QApplication.instance() or QApplication(sys.argv)

# ---------- 1. Старт каркаса ----------
check(app_common.APP_UI_FLAVOR == "tk", "app_common: флейвор Tk-сборки не изменён (M2 не трогает)")
check(mq.UI_FLAVOR == "qt", "каркас Qt: признак интерфейса qt")

ctrl = mq.AppController(qt_app)
check(ctrl.window.windowTitle() == single_instance.WINDOW_TITLE,
      "каркас: заголовок окна совпадает с контрактом single_instance (FindWindow)")
check(ctrl.nav.count() == len(mq.NAV_PAGES) == 6, "каркас: 6 разделов в панели")
check(ctrl.current_page == "spec", "каркас: раздел по умолчанию — «Спецификация» (CP1)")
check(ctrl.window.minimumSize().width() == mq.MIN_WINDOW_WIDTH
      and ctrl.window.minimumSize().height() == mq.MIN_WINDOW_HEIGHT,
      "каркас: минимальный размер окна 1120×620 (как Tk)")
check(ctrl.store is not None and hasattr(ctrl.store, "status_counts"),
      "каркас: DataStore подключён и принадлежит каркасу")
check(not os.path.exists(os.path.join(storage.data_dir(), "settings.json")),
      "профиль: settings.json при старте не создаётся (контракт первой установки)")

ctrl.window.show()
qt_app.processEvents()

# ---------- 2. Статус и хранилище ----------
check("Контрагентов: 0" in ctrl.status_label.text(), "статус: пустой реестр в футере")
_party = ctrl.store.add_counterparty("ooo", {
    "ИНН": "7701234567", "КПП": "770101001", "ОГРН": "1037739010221",
    "Полное название компании (Покупь)": "ООО Ромашка (smoke_qt)",
    "Сокращенное название компании": "ООО Ромашка",
    "ФИО директора": "Иванов Иван Иванович", "Основание действий": "Устав",
    "Юридический адрес": "г. Москва, ул. Ленина, д. 1",
    "Фактический адрес": "г. Москва, ул. Ленина, д. 1",
    "Email": "info@romashka.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012",
    "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "ПАО Сбербанк", "БИК": "044525225",
})
ctrl.refresh_status()
check("Контрагентов: 1" in ctrl.status_label.text(), "статус: счётчик контрагентов из store")
check(ctrl.select_page("registry") is None and ctrl.current_page == "registry",
      "навигация: переключение раздела")
ctrl.select_page("spec")

# ---------- 3. Тост ----------
ctrl.notify("Smoke-каркас: документ сохранён", path="/tmp/smoke_qt_example.docx")
check(ctrl.toast.isVisible(), "тост: показывается")
check(ctrl.toast_open.isVisible() and ctrl.toast_folder.isVisible(),
      "тост: действия «Открыть»/«Показать в папке» при path")
ctrl._hide_toast()
check(not ctrl.toast.isVisible(), "тост: скрыт по команде")

# ---------- 4. Диалоги подменены и не блокируют ----------
ctrl.alert(ctrl.window, "Проверка", "Сообщение алерта", "error")
check(len(ALERTS) == 1, "диалоги: подмена alert не блокирует прогон")

# ---------- 5. Таймеры опроса ----------
_ticks: list[int] = []
ctrl.add_poll_timer("smoke_probe", 15, lambda: _ticks.append(1))
check(ctrl._timers["smoke_probe"].isActive(), "таймеры: poll-таймер запущен")


def _spin_for(ms: int) -> None:
    loop_timer = QTimer()
    loop_timer.setSingleShot(True)
    loop_timer.timeout.connect(qt_app.quit)
    loop_timer.start(ms)
    qt_app.exec()


_spin_for(120)
check(len(_ticks) >= 3, f"таймеры: poll-таймер тикает ({len(_ticks)} тиков за 120 мс)")
ctrl.stop_poll_timer("smoke_probe")
_frozen = len(_ticks)
_spin_for(80)
check(len(_ticks) == _frozen, "таймеры: stop_poll_timer останавливает опрос")

# ---------- 6. Perf-каналы старта (EWDOCHUB_PERF_LOG=1 включён выше) ----------
ctrl._emit_startup_perf()
qt_app.processEvents()
perflog.flush()
check(os.path.exists(perflog.perf_path()), "perf: журнал создан при EWDOCHUB_PERF_LOG=1")
with open(perflog.perf_path(), encoding="utf-8") as fh:
    _perf = fh.read()
check("startup.built" in _perf, "perf: событие startup.built записано")
check("startup.ready" in _perf, "perf: событие startup.ready записано")

# ---------- 7. Тема: переключение и переживание перезапуска ----------
check(qt_theme.current_mode() == "light", "тема: стартовая светлая (пустой settings.json)")
ctrl.set_theme("dark")
check(qt_theme.current_mode() == "dark", "тема: переключение на тёмную")
with open(app_common.settings_path(), encoding="utf-8") as fh:
    _settings = json.load(fh)
check(_settings.get("theme") == "dark", "тема: ключ settings.json «theme» — прежний, dark")
ctrl2 = mq.AppController(qt_app)  # «перезапуск»: новый каркас на том же профиле
check(qt_theme.current_mode() == "dark", "тема: тёмная переживает пересоздание каркаса")
ctrl2.set_theme("light")
check(app_common.load_settings().get("theme") == "light", "тема: возврат на светлую сохранён")

# ---------- 8. Выход: закрытие окна завершает цикл событий ----------
QTimer.singleShot(150, ctrl.window.close)
_exit_code = ctrl.run()
check(_exit_code == 0, f"выход: QApplication.exec вернул 0 (получено {_exit_code})")
check(ctrl._timers == {}, "выход: таймеры опроса остановлены при закрытии")
check(not ctrl.window.isVisible(), "выход: окно закрыто")

# ---------- 9. Журналы ----------
check(os.path.exists(applog.log_path()), "журналы: app.log создан в профиле")
with open(applog.log_path(), encoding="utf-8") as fh:
    _log = fh.read()
check("Запуск EW.DocHub 2.28.2-qt" in _log,
      "журналы: строка старта с версией-флейвором «2.28.2-qt»")
check("Каркас Qt построен" in _log, "журналы: построение каркаса записано")
check("Каркас Qt закрыт" in _log, "журналы: закрытие каркаса записано")
check(os.path.exists(os.path.join(applog.log_dir(), applog.CRASH_FILE_NAME)),
      "журналы: crash.log создан (faulthandler включён)")

# ---------- 10. База: совместимость с Tk-сборкой ----------
check(os.path.exists(storage.db_path()), "хранилище: база SQLite создана")
with sqlite3.connect(storage.db_path()) as _conn:
    _meta = dict(_conn.execute("SELECT key, value FROM meta"))
check(int(_meta.get("schema_version", "0")) == storage.SQLITE_SCHEMA_VERSION,
      "хранилище: schema_version не изменён (Схема данных — запрет раздела З)")

# ---------- 11. Второй экземпляр: контракт блокировки ----------
check(single_instance.mutex_name().startswith("Global\\EWDocHub_SingleInstance"),
      "single instance: имя мьютекса данных прежнее")
check(single_instance.INSTALLER_MUTEX_NAME == "Global\\EWDocHub_InstallerMutex",
      "single instance: инсталляторный мьютекс прежний (AppMutex Inno Setup)")
if sys.platform == "win32":
    _lock1 = single_instance.InstanceLock()
    _lock2 = single_instance.InstanceLock()
    check(_lock1.acquire() is True, "single instance: первый экземпляр блокирует")
    check(_lock2.acquire() is False, "single instance: второй экземпляр в процессе отклонён")
    _lock1.release()

# ---------- 12. Спецификация CP1 (M3) ----------
# Живая страница заменяет плейсхолдер; диалоги подменены выше (confirm=True).
qapp = qt_app  # имя разделов CP1 — как в тестах pytest-qt
check(ctrl.pages["spec"] is ctrl.spec_flow.page,
      "CP1: страница «Спецификация» живая (плейсхолдер заменён)")
check(ctrl.stack.currentWidget() is ctrl.spec_flow.page,
      "CP1: активный раздел каркаса — «Спецификация»")
check(not ctrl.spec_flow.preview_text.toPlainText().startswith("Не удалось"),
      "CP1: предпросмотр собран без ошибки сборки")

# Контрагент раздела 2 + договор-основание — контекст страницы (п. 4.2).
_contract_rec = ctrl.store.add_document(
    storage.KIND_CONTRACT, _party["id"], "516", "10 января",
    data=dict(_party.get("fields", {})),
    extra={"day": "10", "month": "января", "is_ip": False},
    text="",
)
ctrl.spec_flow.refresh_parties()
qapp.processEvents()
check(ctrl.spec_flow.contract_combo.currentText() == "№ 516 от 10 января",
      "CP1/п. 4.2: договор-основание подставлен автоматически")

# PERF-1: вставка 600 строк из Excel даёт 500 и честное сообщение.
_confirms_m3: list[tuple] = []
mq.qt_dialogs.confirm = lambda *a, **k: _confirms_m3.append(a) or True
qapp.clipboard().setText("\n".join(
    f"Позиция {n}\t1\tшт\t10" for n in range(1, 601)
))
import time as _time
_paste_started = _time.perf_counter()
ctrl.spec_flow.table.paste_clipboard()
_paste_ms = (_time.perf_counter() - _paste_started) * 1000.0
print(f"perf: вставка 500 строк в Qt-таблицу — {_paste_ms:.1f} мс "
      "(порог выставляется после К-2; Windows-замер — runbook M3)")
check(ctrl.spec_flow.table.rowCount() == 500,
      "CP1/PERF-1: вставка 600 строк даёт ровно 500")
check("Вставлено 500" in ctrl.toast_label.text()
      and "отброшено 100" in ctrl.toast_label.text(),
      "CP1/PERF-1: сообщение «вставлено 500, отброшено 100» — тостом")

# U12: итог под таблицей — один spec_total с печатной формой.
qapp.clipboard().setText("Кабель ВВГнг 3х2.5\t120\tм\t85,5\n")
ctrl.spec_flow.table.paste_clipboard()
check(ctrl.spec_flow.totals_label.text() == "Позиций: 10 · Итого: 10260,00",
      "CP1/U12: итог под таблицей = spec_total позиций (10260,00)")

# Предпросмотр из шаблона + сохранение (транзакция сервисного слоя).
# Папка документов изолируется в профиль смоука (не трогаем «Документы»
# реального пользователя — профиль прогона, раздел Е).
_docs_dir = os.path.join(os.environ["APPDATA"], "documents_smoke")
os.makedirs(_docs_dir, exist_ok=True)
ctrl.spec_flow.documents_dir = lambda: _docs_dir  # noqa: B010 — подмена прогона
ctrl.spec_flow.number_entry.setText("1")
qapp.processEvents()
ctrl.spec_flow.flush_spec_preview()
check("СПЕЦИФИКАЦИЯ" in ctrl.spec_flow.preview_text.toPlainText(),
      "CP1: предпросмотр — печатная форма из шаблона spec_template.docx")
ctrl.spec_flow.save_spec()
_spec_docs = [d for d in ctrl.store.documents if d["kind"] == storage.KIND_SPEC]
check(len(_spec_docs) == 1, "CP1: сохранение — запись в реестре")
check(ctrl.store.get_document(_spec_docs[0]["id"]) is not None
      and os.path.isdir(ctrl.spec_flow.documents_dir()),
      "CP1: транзакция завершена — запись и папка документов на месте")
check("Спецификация № 1" in ctrl.toast_label.text(),
      "CP1: честное сообщение сохранения с номером")
check(ctrl.spec_flow.is_form_dirty() is False,
      "CP1: после сохранения форма «чистая» (baseline пересчитан)")

# Защита несохранённых при закрытии (п. 4.4) — на «перезапущенном» каркасе.
ctrl2.spec_flow.number_entry.setText("2")
check(ctrl2.spec_flow.is_form_dirty() is True,
      "CP1: правка номера — «грязная» форма")
_ctrl2_confirms: list[tuple] = []


def _recording_confirm(*args, **kwargs):
    _ctrl2_confirms.append(args)
    return True


_prev_confirm = mq.qt_dialogs.confirm
mq.qt_dialogs.confirm = _recording_confirm
ctrl2.window.show()
qapp.processEvents()
ctrl2.window.close()
qapp.processEvents()
mq.qt_dialogs.confirm = _prev_confirm
check(any("Несохранённые изменения" == str(a[1]) for a in _ctrl2_confirms),
      "CP1/п. 4.4: закрытие с «грязной» формой — подтверждение")
check(not ctrl2.window.isVisible(),
      "CP1/п. 4.4: окно закрыто после подтверждения")

# ---------- 13. Карточка контрагента M4 ----------
# Страница «Добавить контрагента» живая; диалоги подменены выше.
check(ctrl.pages["counterparty"] is ctrl.counterparty_flow.page,
      "M4: страница «Добавить контрагента» живая (плейсхолдер заменён)")
check(ctrl.counterparty_flow.current_mode() == "ooo"
      and "ОГРНИП" not in ctrl.counterparty_flow.entries["ooo"],
      "M4: состав полей ООО — Tk-паритет (без ОГРНИП)")
ctrl.counterparty_flow.on_type_change("ИП")
check(ctrl.counterparty_flow.current_mode() == "ip"
      and ctrl.counterparty_flow.stack.currentIndex() == 1,
      "M4: переключение типа карточки ООО/ИП")
check(F.FACT_ADDRESS not in ctrl.counterparty_flow.entries["ip"],
      "M4: у ИП фактический адрес повторяет юридический (отдельного поля нет)")
ctrl.counterparty_flow.on_type_change("Юр. лицо")

# Валидация: пустая карточка — подсветка + алерт, запись не создаётся.
ctrl.counterparty_flow.save_card()
check(any("Ошибки в данных" == str(a[1]) for a in ALERTS),
      "M4: сохранение пустой карточки — алерт «Ошибки в данных»")
check(ctrl.counterparty_flow.error_labels["ooo"][F.INN].text() == "Заполните поле",
      "M4: подсветка обязательного поля «Заполните поле»")

# U8 без ключа: запрос не стартует, сообщение понятное.
ctrl.counterparty_flow.entries["ooo"]["ИНН"].set_text("7701234567")
ctrl.dadata_flow.fill_from_inn()
check(any("Нет API-ключа" == str(a[1]) for a in ALERTS),
      "M4/U8: без ключа DaData — понятная ошибка, запрос не стартует")
check(ctrl.dadata_flow.busy is False, "M4/U8: busy сброшен после отказа ключа")

# Успешное сохранение карточки (реквизиты с валидными контрольными цифрами).
_m4_card = {
    "ИНН": "2540123459", "КПП": "254001001", "ОГРН": "1035001234564",
    "Полное название компании (Покупатель)": "ООО «Ландыш» (smoke_qt)",
    "Сокращенное название компании": "ООО «Ландыш»",
    "ФИО директора": "Смирнов Смирн Смирнович", "Основание действий": "Устав",
    "Юридический адрес": "690000, г. Владивосток, ул. Светланская, д. 3",
    "Фактический адрес": "690000, г. Владивосток, ул. Светланская, д. 3",
    "Email": "landysh@example.ru", "Телефон": "+7 (912) 345-67-89",
    "Расчётный счёт": "40702810123456789012",
    "Корреспондентский счёт": "30101810100000000225",
    "Название банка": "ПАО Сбербанк", "БИК": "044525225",
}
for _key, _value in _m4_card.items():
    _adapter = ctrl.counterparty_flow.entries["ooo"].get(_key)
    if _adapter is not None:
        _adapter.set_text(_value)
ctrl.counterparty_flow.save_card()
_m4_parties = ctrl.store.active_counterparties()
check(len(_m4_parties) == 2, "M4: карточка сохранена в реестр (2 контрагента)")
check(ctrl.current_party_id == _m4_parties[-1]["id"]
      and ctrl.current_page == "spec",
      "M4: после сохранения — контрагент текущий, раздел «Спецификация»")
check("добавлен в реестр" in ctrl.toast_label.text(),
      "M4: честное сообщение добавления")
check(ctrl.counterparty_flow.is_form_dirty() is False,
      "M4: после сохранения карточка «чистая» (baseline пересчитан)")

# Окно правки через кнопку на панели «Документ для» (M4-точка до M5).
ctrl.spec_flow.edit_current_party()
_edit_dialogs = ctrl.window.findChildren(QDialog)
check(len(_edit_dialogs) == 1, "M4: кнопка «Редактировать карточку» открывает окно")
_edit_dialogs[0].entries[F.DIRECTOR].set_text("Кузнецов Кузьма Кузьмич")
_edit_dialogs[0].save()
check(ctrl.store.get_counterparty(_m4_parties[-1]["id"])["fields"][F.DIRECTOR]
      == "Кузнецов Кузьма Кузьмич", "M4: правка карточки — реквизит обновлён")
check("обновлена" in ctrl.toast_label.text(),
      "M4: честное сообщение правки (документы не изменялись)")
_edit_dialogs[0].deleteLater()

# Защита страницы п. 4.4: Ctrl+S на «грязной» карточке (новый ИНН — без
# диалога дубликата, п. 4.1).
ctrl.select_page("counterparty")
ctrl.counterparty_flow.entries["ooo"][F.DIRECTOR].set_text("Одинцов Один Одинович")
ctrl.counterparty_flow.entries["ooo"][F.INN].set_text("7701234560")
ctrl.counterparty_flow.entries["ooo"][F.SHORT_NAME].set_text("ООО «Незабудка» (smoke_qt)")
check(ctrl.counterparty_flow.is_form_dirty() is True,
      "M4: правка карточки — страница «грязная»")
ctrl.save_current_page()
check(len(ctrl.store.active_counterparties()) == 3,
      "M4/Ctrl+S: сохранение карточки по Ctrl+S")
check(ctrl.counterparty_flow.is_form_dirty() is False,
      "M4/Ctrl+S: после сохранения страница «чистая»")


# ---------- 14. Реестр M5 (TREE-1) ----------
# Синхронная замена DeferredDelete: диалог M4 ещё не уничтожен к этому месту.
QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
qapp.processEvents()
check(ctrl.pages["registry"] is ctrl.registry_flow.page,
      "M5: страница «Реестр» живая (плейсхолдер заменён)")
ctrl.select_page("registry")
check(ctrl.registry_flow._registry_rendered is True,
      "M5: ленивый рендер — список построен при открытии страницы")
_model = ctrl.registry_flow.table.model
_m5_ids = [rec["id"] for rec in ctrl.store.active_counterparties()]
check(_model.order() == tuple(f"p:{i}" for i in _m5_ids),
      "M5: модель строк = данным (3 контрагента, iid «p:<id>»)")
check(ctrl.registry_flow._registry_caption.text() == f"Контрагенты ({len(_m5_ids)})",
      "M5: подпись списка «Контрагенты (N)»")

# Поиск по ИНН (цифры запроса — подстрока ИНН, Tk-контракт).
ctrl.registry_flow.search_edit.setText("234560")
ctrl.registry_flow.apply_filter()
check(len(_model.order()) == 1,
      "M5/U11: поиск по ИНН фильтрует на лету (placeholder, Enter — сразу)")
ctrl.registry_flow.search_edit.clear()
ctrl.registry_flow.apply_filter()

# Выделение и панель действий (V3).
ctrl.registry_flow.on_tree_select(f"p:{_m5_ids[0]}")
check(ctrl.registry_flow._registry_action_btn1.text() == "Редактировать"
      and ctrl.registry_flow._registry_action_btn2.text() == "Удалить",
      "M5/V3: панель действий контрагента — «Редактировать»/«Удалить»")

# «Редактировать» — окно правки M4 из реестра.
ctrl.registry_flow.run_action("party", ctrl.store.get_counterparty(_m5_ids[0]))
_m5_dialogs = ctrl.window.findChildren(QDialog)
check(len(_m5_dialogs) == 1, "M5: «Редактировать» открывает окно правки карточки")
_m5_dialogs[0].request_close()  # без изменений — закрывается сразу (DEF-1)
_m5_dialogs[0].deleteLater()
QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
check(len(ctrl.window.findChildren(QDialog)) == 0,
      "M5/DEF-1: окно правки без изменений закрывается без подтверждения")

# Delete через панель: подтверждение подменено (True) — мягкое удаление.
_m5_del = ctrl.store.get_counterparty(_m5_ids[1])
ctrl.registry_flow.run_action("party", _m5_del, second=True)
check("sec:trash" in _model.order() and f"tp:{_m5_ids[1]}" in _model.order(),
      "M5: «добавить 3, удалить 1» — секция «Корзина» совпадает с данными")
check(ctrl.registry_flow._registry_empty_trash_btn.isHidden() is False,
      "M5: кнопка «Очистить корзину» видна, когда корзина непуста")

# Восстановление из секции корзины.
_trashed = ctrl.store.trashed_counterparties()[0]
ctrl.registry_flow.on_tree_select(f"tp:{_trashed['id']}")
check(ctrl.registry_flow._registry_action_btn1.text() == "Восстановить",
      "M5/V3: панель записи корзины — «Восстановить»/«Удалить навсегда»")
ctrl.registry_flow.run_action("trash_party", _trashed)
check(_model.order() == tuple(f"p:{i}" for i in _m5_ids),
      "M5: «восстановить 1» — отображение совпадает с данными")

# Drill «Документы — <имя>» и возврат кнопкой «← Назад» (Escape — U13).
ctrl.registry_flow.on_tree_activate(f"p:{_m5_ids[0]}")
check(ctrl.registry_flow._registry_title_label.text() == "Документы — ООО Ромашка",
      "M5/2.19.0: drill — заголовок «Документы — <имя>»")
check([col[1] for col in _model.columns] == ["Документ", "Дата", "Срок"],
      "M5/2.19.0: drill меняет набор колонок")
ctrl.registry_flow._registry_back_btn.click()
check(ctrl.registry_flow._registry_title_label.text() == "Реестр",
      "M5/2.19.0: «← Назад» возвращает к списку контрагентов")

# Кросс-база: операции реестра видны Tk-пути DataStore (обязательный тест M3+).
import sqlite3 as _sq
with _sq.connect(storage.db_path()) as _conn:
    _meta5 = dict(_conn.execute("SELECT key, value FROM meta"))
check(int(_meta5.get("schema_version", "0")) == storage.SQLITE_SCHEMA_VERSION,
      "M5: schema_version не изменён операциями реестра (запрет З.1)")


# ---------- 15. Полный маршрут M6: договор → доп. соглашение → спецификация → импорт → корзина → сверка ----------
# Сценарии Г #4, #5, #13, #14, #15, #16 одной цепочкой на живых страницах.
# Папка документов — настройка профиля смоука (не «Документы» пользователя):
# путь через save_setting — тот же маршрут, что на странице «Настройки».
ctrl.save_setting("documents_dir", _docs_dir)

# --- Договор (Г #4): выбор по ИНН, сохранение, честное «сохранён», Word ---
ctrl.select_page("contract")
check(ctrl.pages["contract"] is ctrl.contract_flow.page,
      "M6/Г#4: страница «Договор поставки» живая (плейсхолдер заменён)")
ctrl.contract_flow.search_entry.setText("7701234567")
ctrl.contract_flow.find_party_by_inn(ctrl.contract_flow.search_entry)
check(ctrl.contract_flow.selected_party() is not None
      and ctrl.contract_flow.selected_party()["id"] == _m5_ids[0],
      "M6/Г#4/U3: поиск по ИНН подставил контрагента «ООО Ромашка»")
check(ctrl.additional_flow.party_combo.currentText()
      == ctrl.contract_flow.party_combo.currentText(),
      "M6/п. 4.2: выбор контрагента перенесён на страницу доп. соглашения")
ctrl.contract_flow.number_entry.setText("601")
ctrl.contract_flow.day_entry.setText("10")
qapp.processEvents()
ctrl.contract_flow.preview_timer.stop()
ctrl.contract_flow.render_preview()
check(len(ctrl.contract_flow.preview_text.toPlainText()) > 500,
      "M6/Г#4: предпросмотр договора — печатная форма из шаблона")
ctrl.contract_flow.save_contract()
_m6_contracts = [d for d in ctrl.store.documents
                 if d["kind"] == storage.KIND_CONTRACT and d["number"] == "601"]
check(len(_m6_contracts) == 1, "M6/Г#4: договор № 601 — запись в реестре")
check("сохранён" in ctrl.toast_label.text() and "601" in ctrl.toast_label.text(),
      "M6/Г#4: честное «сохранён» с номером договора")
_m6_party_dir = os.path.join(_docs_dir, "ООО Ромашка")
check(any(os.path.exists(os.path.join(_m6_party_dir, name))
          for name in (os.listdir(_m6_party_dir) if os.path.isdir(_m6_party_dir) else [])),
      "M6/Г#4: файл договора записан в папку контрагента")

# --- Доп. соглашение (Г #5): договор-основание автоподставлен, «Действует до» ---
ctrl.select_page("additional")
check(ctrl.pages["additional"] is ctrl.additional_flow.page,
      "M6/Г#5: страница «Доп. соглашение» живая (плейсхолдер заменён)")
ctrl.additional_flow.refresh_parties()
check(ctrl.additional_flow.contract_combo.currentText() == "№ 516 от 10 января",
      "M6/п. 5.1: договор-основание подставлен автоматически (первый из реестра)")
check(ctrl.additional_flow.valid_entry.text() ==
      f"31.12.{datetime.now().year}",
      "M6/Г#5: «Действует до» по умолчанию — 31 декабря текущего года (2.18.0)")
ctrl.additional_flow.limit_entry.setText("10000")
check(ctrl.additional_flow.limit_entry.text() == "10 000",
      "M6: лимит форматируется с разрядами")
ctrl.additional_flow.term_entry.setText("10")
check("десять" in ctrl.additional_flow.words_label.text().lower(),
      "M6: срок прописью под полем отсрочки")
check("десять тысяч" in ctrl.additional_flow.limit_words_label.text().lower(),
      "M6: сумма прописью под полем лимита")
ctrl.additional_flow.valid_entry.setText("31.03.2027")
ctrl.additional_flow.save_additional()
_m6_addls = [d for d in ctrl.store.documents
             if d["kind"] == storage.KIND_ADDITIONAL]
check(len(_m6_addls) == 1, "M6/Г#5: доп. соглашение — запись в реестре")
check(_m6_addls[0]["extra"]["valid_until"] == "31.03.2027"
      and _m6_addls[0]["extra"]["valid_until_iso"] == "2027-03-31",
      "M6/Г#5/2.18.0: «Действует до» в extra — вид и ISO")
check("сохранено" in ctrl.toast_label.text(),
      "M6/Г#5: честное «сохранено» доп. соглашения")
check(ctrl.additional_flow.is_form_dirty() is False,
      "M6: после сохранения доп. соглашения форма «чистая»")

# --- Спецификация: третья запись маршрута к тому же контрагенту/договору ---
ctrl.select_page("spec")
ctrl.spec_flow.refresh_parties()
ctrl.spec_flow.number_entry.setText("2")
ctrl.spec_flow.flush_spec_preview()
ctrl.spec_flow.save_spec()
check(len([d for d in ctrl.store.documents
           if d["kind"] == storage.KIND_SPEC]) == 2,
      "M6/маршрут: вторая спецификация сохранена (маршрут полн)")

# --- Импорт пакета (Г #13): U9-сводка, реестр, unknown, отчёт с «Открыть реестр» ---
from docx import Document as _SmokeDoc  # noqa: E402 — как в smoke_test.py
import document_importer as _di  # noqa: E402
import import_qt as _iq  # noqa: E402

_import_src = os.path.join(os.environ["APPDATA"], "import_source")
os.makedirs(_import_src, exist_ok=True)
_import_main = os.path.join(_import_src, "Договор поставки № 888.docx")
_d = _SmokeDoc()
_d.add_paragraph("ДОГОВОР ПОСТАВКИ № 888")
_d.add_paragraph("«11» января 2026 г.")
_d.add_paragraph("ООО «Ромашка-Импорт», именуемое в дальнейшем «Покупатель»,")
_d.add_paragraph("ИНН 2540123456, КПП 254001001")
_d.save(_import_main)
_m6_no_inn = []
for _n in (1, 2):
    # «Старые» файлы без ИНН: тип опознан (договор), реквизитов нет —
    # ok=True, recognized=False → контур U9 (файл требует решения).
    _p = os.path.join(_import_src, f"Договор поставки без ИНН {_n}.docx")
    _d = _SmokeDoc()
    _d.add_paragraph("ДОГОВОР ПОСТАВКИ")
    _d.add_paragraph("ООО «Старый пакет», именуемое в дальнейшем «Покупатель»")
    _d.save(_p)
    _m6_no_inn.append(_p)

import_qt_reports: list[tuple] = []
ctrl.import_flow.show_report = (
    lambda entries, summary: import_qt_reports.append((entries, summary)))
CHOICES.clear()
_CHOICE_RESULT = "create_by_name"
ctrl.import_flow.start([_import_main, *_m6_no_inn],
                       thread_factory=_SyncFactory())
ctrl.import_flow.tick()
qapp.processEvents()
_CHOICE_RESULT = None
check(len(CHOICES) >= 1
      and "файлов без ИНН" in str(CHOICES[-1][2]),
      "M6/Г#13/U9: серия файлов без ИНН — ОДИН диалог сводки")
check(len(CHOICES) == 1,
      "M6/U9: 3 файла — один пакетный диалог, без серии на каждый файл")
check(len(ctrl.store.active_counterparties()) == 5,
      "M6/Г#13: «Ромашка-Импорт» (по ИНН) + «Старый пакет» (по названию)")
check(len([d for d in ctrl.store.documents
           if d["kind"] == storage.KIND_CONTRACT]) >= 2
      and len(import_qt_reports) == 1,
      "M6/Г#13: пакет зарегистрирован, отчёт показан")
_entries_m6, _summary_m6 = import_qt_reports[0]
check("Добавлено документов: 3" in _summary_m6,
      "M6/Г#13: сводка импорта честная (3 добавлено, 2 новых контрагента)")
check(len(_entries_m6) == 3, "M6/п. 3.5: по каждому файлу строка отчёта")
_check_unknown = os.path.isdir(os.path.join(_docs_dir, "unknown"))
check(_check_unknown is False or os.listdir(os.path.join(_docs_dir, "unknown")) == [],
      "M6/п. 3.5: отклонённых файлов нет — все распознаны")

# Отчёт: диалог «Открыть реестр» (U10) — оболочка строится, кнопка ведёт в реестр.
_m6_report = _iq.ImportReportDialog(
    ctrl.window, _entries_m6, _summary_m6, storage.KIND_LABELS, ctrl.import_flow.open_registry)
check(_m6_report.open_registry_btn.text() == "Открыть реестр",
      "M6/U10: кнопка «Открыть реестр» в отчёте импорта")
_m6_report.deleteLater()
QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)

# --- Корзина (Г #14): мягкое удаление импортированного договора и возврат ---
_ctrl_registry_model = ctrl.registry_flow.table.model
_m6_doc_to_trash = next(d for d in ctrl.store.documents
                        if d["kind"] == storage.KIND_CONTRACT
                        and d["number"] == "888")
ctrl.select_page("registry")
ctrl.registry_flow.refresh()
ctrl.registry_flow.run_action("doc", _m6_doc_to_trash, second=True)
check(ctrl.store.get_document(_m6_doc_to_trash["id"], include_deleted=True)
      .get("deleted_at") is not None,
      "M6/Г#14: документ в корзине (мягкое удаление)")
ctrl.registry_flow.run_action("trash_doc", _m6_doc_to_trash)
check(ctrl.store.get_document(_m6_doc_to_trash["id"]).get("deleted_at") is None,
      "M6/Г#14: документ восстановлен из корзины")
# TRASH-WARN: просроченная запись удаляется со сводкой в журнале.
_m6_trash_doc = ctrl.store.get_document(_m6_doc_to_trash["id"])
_m6_trash_doc["deleted_at"] = (
    datetime.now() - timedelta(days=storage.TRASH_RETENTION_DAYS + 5)
).isoformat(timespec="seconds")
_m6_warn_before = len(_log) if isinstance(_log, str) else 0
ctrl.cleanup_expired_trash(planned=1)
check(ctrl.store.get_document(_m6_doc_to_trash["id"]) is None,
      "M6/TRASH-WARN: просроченная запись корзины удалена")
check("Автоочистка корзины" in ctrl.toast_label.text(),
      "M6/TRASH-WARN: итог автоочистки в немодальном уведомлении")

# --- Сверка «реестр ↔ папка» (Г #15): сирота → баннер → «Восстановить в реестр» ---
_m6_orphan_dir = os.path.join(_docs_dir, "ООО Ромашка")
os.makedirs(_m6_orphan_dir, exist_ok=True)
_m6_orphan = os.path.join(_m6_orphan_dir, "Договор № 990 от июня.docx")
with open(_m6_orphan, "wb") as _fh:
    _fh.write(b"PK\x03\x04")
_CHOICE_RESULT = "restore"
CHOICES.clear()
ctrl.registry_flow.scan_consistency(orphans=[_m6_orphan])
_m6_banner_texts = [label.text() for label in ctrl.pages["registry"].findChildren(QLabel)]
check(any("файлов вне реестра: 1" in t for t in _m6_banner_texts),
      "M6/Г#15/DEF-3: баннер сирот над списком реестра")
ctrl.registry_flow.open_reconcile_dialog()
qapp.processEvents()
QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
_CHOICE_RESULT = None
check(len(CHOICES) == 1 and "Файлы вне реестра" == str(CHOICES[0][1]),
      "M6/Г#15: диалог разбора сирот показан")
check(any(d["number"] == "990" and d["extra"].get("restored")
          for d in ctrl.store.documents),
      "M6/Г#15: сирота восстановлена в реестр (тип и номер по имени файла)")
_m6_banner_texts = [label.text() for label in ctrl.pages["registry"].findChildren(QLabel)]
check(not any("файлов вне реестра" in t for t in _m6_banner_texts),
      "M6/Г#15: баннер исчез после разбора")

# --- Настройки (Г #16): тема, папка, ключ DaData (профиль смоука) ---
# Радиоиндикаторы могут отставать от фактической темы (смена была через
# другой каркас) — переключение идёт через страницу: light → dark.
ctrl.select_page("settings")
check(ctrl.pages["settings"] is ctrl.settings_flow.page,
      "M6/Г#16: страница «Настройки» живая (плейсхолдер заменён)")
ctrl.settings_flow._theme_light.setChecked(True)
ctrl.settings_flow._theme_dark.setChecked(True)
check(qt_theme.current_mode() == "dark"
      and app_common.load_settings().get("theme") == "dark",
      "M6/Г#16: смена темы со страницы сохраняется в settings.json")
check(ctrl.settings_flow._dadata_key_entry.echoMode()
      == QLineEdit.EchoMode.Password,
      "M6/Г#16: поле ключа DaData скрыт ввод (парольный режим)")
check(ctrl.settings_flow._documents_dir() == _docs_dir,
      "M6/Г#16: путь папки документов со страницы настроек")

# Пинг UI при фоновой сверке уже покрыт тестами startup_qt; здесь —
# контроль, что живые таймеры каркаса продолжают тикать после маршрута.
_ticks_m6: list[int] = []
ctrl.add_poll_timer("m6_probe", 15, lambda: _ticks_m6.append(1))
_spin_for(100)
check(len(_ticks_m6) >= 2,
      f"M6: таймеры каркаса отзывчивы после полного маршрута ({len(_ticks_m6)} тиков)")
ctrl.stop_poll_timer("m6_probe")


# ---------- Сводка ----------
try:
    perflog.reset_for_tests()
except Exception:
    pass

print("SMOKE QT OK" if not FAILURES else "SMOKE QT FAILED")
if FAILURES:
    print(f"\n=== СВОДКА: {len(FAILURES)} из {_TOTAL_CHECKS} проверок не прошли ===")
    for _message in FAILURES:
        print(f"FAIL: {_message}")
    sys.exit(min(len(FAILURES), 100))
print(f"SMOKE QT OK: {_TOTAL_CHECKS} проверок")
