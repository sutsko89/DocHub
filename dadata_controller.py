"""Автозаполнение карточки контрагента по ИНН — второй контроллер ContractApp.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет жизненным циклом запроса DaData:
guard повторных кликов, busy-анимация, фоновый HTTP-воркер, опрос очереди
результата из главного потока Tk и подстановка реквизитов в поля формы.

Исправления аудита 2.0.1, сохранённые здесь:
  * повторные клики «Заполнить по ИНН» при активном запросе игнорируются
    (guard busy) — раньше два параллельных воркера гонялись за полями формы;
  * результат приходит через ``queue.Queue`` и опрашивается из главного
    потока — рабочий поток не трогает Tk и не вызывает ``root.after()``
    (гарантий потокобезопасности Tcl нет).

U8-полностью (UX-B, этап 5): «Заполнить по ИНН» работает не только на
странице «Добавить контрагента», но и в окне редактирования карточки
(edit_counterparty, DEF-1). Контроллер адресуется через :class:`DadataTarget`
— снимок адресата запроса (поля формы, режим «ip»/«ooo», строка статуса,
родительский экземпляр для диалогов). Ссылку на приложение контроллер
сохраняет по-прежнему: таймеры after ставятся на главное окно (оно живёт
дольше окна правки), а все обращения к полям и статусу идут через цель.

Контракт владения состоянием: подстановка реквизитов выполняется ТОЛЬКО
в главном потоке (result_tick из after); рабочий поток кладёт результат
в очередь. Окно правки могло закрыться, пока запрос летел по сети, —
потеря TclError при обращении к разрушенным виджета/переменной означает
тихий отказ подстановки, диалог уже некому показывать.
"""

from __future__ import annotations

import queue
import threading
from collections.abc import Callable
from typing import Any

import tkinter as tk

import dadata_client as dadata
import ui_macos as ui
from document_builder import F


class DadataTarget:
    """Адресат «Заполнить по ИНН»: поля формы, режим, статус, окно-владелец.

    ``entries`` — словарь полей РЕШАЮЩЕГО режима (у страницы приложения
    это ``app._cp_entries[mode]``, у окна правки — локальный ``entries``);
    ``mode`` — фиксированный режим подстановки ("ip"/"ooo"); ``status_var``
    — строка прогресса/результата; ``root`` — окно, к которому привязаны
    диалоги (у окна правки диалоги модальны к нему, а не к главному окну).
    """

    def __init__(self, root: Any, entries: dict, mode: str, status_var: Any):
        self.root = root
        self.entries = entries
        self.mode = mode
        self.status_var = status_var

    @classmethod
    def from_app(cls, app: Any) -> "DadataTarget":
        """Цель «страница добавления»: режим читается из _cp_type_var."""
        mode = "ip" if app._cp_type_var.get() == "ИП" else "ooo"
        return cls(app.root, app._cp_entries[mode], mode, app._cp_status_var)


class DadataController:
    """Жизненный цикл «Заполнить по ИНН»: guard → HTTP → подстановка."""

    def __init__(self, app: Any) -> None:
        self.app = app
        # Состояние активного запроса.
        self.busy = False
        self.busy_after = None        # id таймера анимации для after_cancel
        self.busy_phase = 0
        self.result_queue: "queue.Queue | None" = None
        self._target: "DadataTarget | None" = None

    # ------------------------------------------------------------------ #
    #  Запрос
    # ------------------------------------------------------------------ #
    def fill_from_inn(self, thread_factory: Callable[..., Any] | None = None,
                      target: "DadataTarget | None" = None) -> None:
        """Запрашивает реквизиты по ИНН из активной карточки (ООО или ИП).

        ``target=None`` — страница добавления контрагента (прежнее
        поведение); окно правки передаёт собственную цель. Guard повторных
        кликов стоит в начале: пока запрос выполняется, новый воркер не
        запускается (аудит 2.0.1). ``thread_factory`` разрешается в момент
        вызова (по умолчанию — ``threading.Thread``): smoke-тест подменяет
        ``main.threading.Thread`` синхронным исполнителем, и подмена должна
        действовать в момент запуска воркера.
        """
        if self.busy:
            return
        self._target = target or DadataTarget.from_app(self.app)
        mode = self._target.mode
        widget = self._target.entries.get(F.INN)
        inn = widget.get().strip() if widget is not None else ""
        expected = 12 if mode == "ip" else 10
        if not inn.isdigit() or len(inn) != expected:
            ui.alert(
                self._alert_parent(), "Проверьте ИНН",
                f"ИНН должен содержать {expected} цифр.", "error",
            )
            self._target = None
            return

        api_key = dadata.get_api_key()
        if not api_key:
            ui.alert(
                self._alert_parent(), "Нет API-ключа",
                "Ключ DaData не задан. Укажите его на странице «Настройки» "
                "или заполните реквизиты вручную.", "error",
            )
            self._target = None
            return

        self.busy = True
        self.busy_phase = 0
        # Очередь результата + опрос из главного потока (как у пакетного
        # импорта): рабочий поток не трогает Tk и не вызывает root.after()
        # (аудит 2.0.1 — гарантий потокобезопасности Tcl нет).
        result_queue: queue.Queue = queue.Queue()
        self.result_queue = result_queue
        self.busy_tick()
        factory = thread_factory or threading.Thread
        factory(
            target=self.fetch_worker,
            args=(inn, api_key, mode, result_queue),
            daemon=True,
        ).start()
        self.app.root.after(80, self.result_tick)

    # ------------------------------------------------------------------ #
    #  Анимация ожидания
    # ------------------------------------------------------------------ #
    def busy_tick(self) -> None:
        """Кадр анимации «Загружаю реквизиты…» (раз в 300 мс)."""
        if not self.busy:
            return
        dots = "." * (self.busy_phase % 4)
        var = getattr(self._active_target(), "status_var", None)
        try:
            if var is not None:
                var.set(f"Загружаю реквизиты из ЕГРЮЛ/ЕГРИП{dots}")
        except tk.TclError:
            # Окно-владелец закрылось: анимация не нужна, цикл остановится
            # в busy_stop() из result_tick/party_*.
            self.busy_stop()
            return
        self.busy_phase += 1
        self.busy_after = self.app.root.after(300, self.busy_tick)

    def busy_stop(self) -> None:
        """Останавливает анимацию и цикл опроса результата."""
        self.busy = False
        if self.busy_after is not None:
            try:
                self.app.root.after_cancel(self.busy_after)
            except Exception:
                pass
            self.busy_after = None
        self.result_queue = None  # цикл опроса останавливается
        self._target = None

    # ------------------------------------------------------------------ #
    #  Фоновый воркер и опрос результата
    # ------------------------------------------------------------------ #
    def fetch_worker(self, inn: str, api_key: str, mode: str,
                     result_queue: "queue.Queue") -> None:
        """Рабочий поток DaData: только HTTP и очередь — никаких вызовов Tk."""
        try:
            info = dadata.fetch_party(inn, api_key)
        except Exception as exc:
            result_queue.put(("error", str(exc)))
            return
        result_queue.put(("ok", mode, info))

    def result_tick(self) -> None:
        """Опрашивает очередь результата DaData из главного потока Tk."""
        result_queue = self.result_queue
        if result_queue is None:
            return  # запрос завершён или окно закрывается
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            self.app.root.after(80, self.result_tick)
            return
        except (tk.TclError, AttributeError):
            return
        kind = item[0]
        if kind == "error":
            self.party_failed(item[1])
        elif kind == "ok":
            self.party_loaded(item[1], item[2])

    # ------------------------------------------------------------------ #
    #  Подстановка реквизитов
    # ------------------------------------------------------------------ #
    def _active_target(self, mode: str | None = None) -> "DadataTarget":
        """Активная цель запроса; без цели — страница добавления.

        Прямые вызовы party_loaded/party_failed/busy_tick (тесты, фасады)
        исторически адресуют страницу приложения. Если при этом передан
        явный ``mode`` (подстановка знает режим из результата запроса),
        поля берутся по нему, а не по типу, выбранному на странице.
        """
        if self._target is not None:
            return self._target
        if mode is not None:
            return DadataTarget(
                self.app.root, self.app._cp_entries[mode], mode,
                self.app._cp_status_var,
            )
        return DadataTarget.from_app(self.app)

    def _alert_parent(self) -> Any:
        """Окно для диалогов: цель (окно правки), если жива; иначе главное.

        Fake-объекты тестов без Tk не имеют winfo_exists — цель считается
        живой (в тестах target.root и app.root совпадают либо проверяется
        сам факт вызова alert с parent).
        """
        target = self._target
        win = getattr(target, "root", None) if target is not None else None
        if win is not None and hasattr(win, "winfo_exists"):
            try:
                if not bool(win.winfo_exists()):
                    return self.app.root
            except tk.TclError:
                return self.app.root
        return win if win is not None else self.app.root

    def party_failed(self, message: str) -> None:
        """Запрос не удался: стоп-анимация, понятная ошибка пользователю."""
        parent = self._alert_parent()
        status_var = getattr(self._active_target(), "status_var", None)
        self.busy_stop()
        self._reset_status(status_var)
        ui.alert(parent, "Не удалось заполнить реквизиты", message, "error")

    def party_loaded(self, mode: str, info: dict) -> None:
        """Карточка получена: реквизиты подставляются в поля активного режима."""
        parent = self._alert_parent()
        target = self._active_target(mode)
        status_var = getattr(target, "status_var", None)
        self.busy_stop()
        filled: list[str] = []

        def put(key: str, value: str):
            if not value:
                return
            widget = target.entries.get(key)
            if widget is None:
                return
            widget.delete(0, tk.END)
            widget.insert(0, value)
            filled.append(key)

        try:
            put(F.INN, info["inn"])
            put(F.LEGAL_ADDRESS, info["legal_address"])
            put(F.BANK, info["bank_name"])
            put(F.BIK, info["bik"])
            put(F.CORR_ACCOUNT, info["correspondent_account"])

            if mode == "ooo":
                put(F.FULL_NAME, info["full_name"])
                put(F.SHORT_NAME, info["short_name"])
                put(F.DIRECTOR, info["director"])
                put(F.KPP, info["kpp"])
                put(F.OGRN, info["ogrn"])
                fact = target.entries.get(F.FACT_ADDRESS)
                if fact is not None and not fact.get().strip():
                    put(F.FACT_ADDRESS, info["legal_address"])
            else:
                put(F.IP_NAME, info["director"])
                put(F.OGRNIP, info["ogrn"])
        except tk.TclError:
            # Окно правки закрылось, пока запрос был в полёте (U8): тихий
            # отказ — диалог уже некому показывать, состояние формы не важно.
            return

        if not filled:
            ui.alert(
                parent, "Данных нет",
                "Сервис вернул карточку без данных для подстановки.", "error",
            )
            return

        self._reset_status(status_var, "Реквизиты заполнены из ЕГРЮЛ/ЕГРИП")
        # V5/UX-B: реквизиты подставлены программно — нажатий клавиш не было,
        # точку «грязной» формы обновляем явно (если приложение умеет).
        updater = getattr(self.app, "_update_dirty_dots", None)
        if updater is not None:
            try:
                updater()
            except Exception:
                pass

    @staticmethod
    def _reset_status(status_var: Any, text: str = "") -> None:
        """Очистка/запись строки статуса после завершения запроса.

        Окно-владелец могло закрыться: TclError тишина — статус уже некому
        показывать.
        """
        if status_var is None:
            return
        try:
            status_var.set(text)
        except (tk.TclError, RuntimeError):
            pass
