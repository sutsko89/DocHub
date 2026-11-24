"""Smoke Qt-сборки EW.DocHub (этапы M2–M4 миграции Tk → PySide6).

Проверяет критерии этапов M2 (каркас: старт → построение окна и разделов →
журналы → perf-события старта → тема → выход), M3 (контрольная точка CP1:
страница «Спецификация» — вставка блока из Excel с лимитом 500, итог,
предпросмотр из шаблона без записи на диск, сохранение через
services.save_document, защита несохранённых при закрытии) и M4 (карточка
контрагента: состав полей ООО/ИП, валидация с подсветкой, «Заполнить по
ИНН» без ключа, окно правки карточки, Ctrl+S, защита страницы). Запуск:

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
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QDialog

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
mq.qt_dialogs.alert = lambda *args, **kwargs: ALERTS.append(args)
mq.qt_dialogs.confirm = lambda *args, **kwargs: True

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
