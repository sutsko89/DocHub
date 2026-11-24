"""Smoke-скелет Qt-каркаса EW.DocHub (этап M2 миграции Tk → PySide6).

Проверяет критерии этапа M2 плана миграции на живом каркасе:
старт → построение окна и разделов → журналы → perf-события старта →
тема (переключение + перезапуск) → выход (закрытие окна завершает цикл
событий). Запуск:

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
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

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
