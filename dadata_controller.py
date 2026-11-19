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

Контроллер получает ссылку на приложение и обращается к его состоянию
(``_cp_entries``, ``_cp_type_var``, ``_cp_status_var``, ``root``) через
``self.app``; вызовы, которые подменяет smoke-тест, идут через фасады
ContractApp с прежними именами.
"""

from __future__ import annotations

import queue
from collections.abc import Callable
from typing import Any

import tkinter as tk

import dadata_client as dadata
import ui_macos as ui
from document_builder import F


class DadataController:
    """Жизненный цикл «Заполнить по ИНН»: guard → HTTP → подстановка."""

    def __init__(self, app: Any) -> None:
        self.app = app
        # Состояние активного запроса.
        self.busy = False
        self.busy_after = None        # id таймера анимации для after_cancel
        self.busy_phase = 0
        self.result_queue: "queue.Queue | None" = None

    # ------------------------------------------------------------------ #
    #  Запрос
    # ------------------------------------------------------------------ #
    def fill_from_inn(self, thread_factory: Callable[..., Any]) -> None:
        """Запрашивает реквизиты по ИНН из активной карточки (ООО или ИП).

        Guard повторных кликов стоит в начале: пока запрос выполняется,
        новый воркер не запускается (аудит 2.0.1). ``thread_factory``
        передаётся вызывающим кодом на каждый запуск: smoke-тест подменяет
        ``main.threading.Thread`` синхронным исполнителем, и подмена должна
        действовать в момент вызова.
        """
        if self.busy:
            return
        mode = "ip" if self.app._cp_type_var.get() == "ИП" else "ooo"
        widget = self.app._cp_entries[mode].get(F.INN)
        inn = widget.get().strip() if widget is not None else ""
        expected = 12 if mode == "ip" else 10
        if not inn.isdigit() or len(inn) != expected:
            ui.alert(
                self.app.root, "Проверьте ИНН",
                f"ИНН должен содержать {expected} цифр.", "error",
            )
            return

        api_key = dadata.get_api_key()
        if not api_key:
            ui.alert(
                self.app.root, "Нет API-ключа",
                "Ключ DaData не задан. Укажите его на странице «Настройки» "
                "или заполните реквизиты вручную.", "error",
            )
            return

        self.busy = True
        self.busy_phase = 0
        # Очередь результата + опрос из главного потока (как у пакетного
        # импорта): рабочий поток не трогает Tk и не вызывает root.after()
        # (аудит 2.0.1 — гарантий потокобезопасности Tcl нет).
        result_queue: queue.Queue = queue.Queue()
        self.result_queue = result_queue
        self.busy_tick()
        thread_factory(
            target=self.app._cp_fetch_worker,
            args=(inn, api_key, mode, result_queue),
            daemon=True,
        ).start()
        self.app.root.after(80, self.app._cp_result_tick)

    # ------------------------------------------------------------------ #
    #  Анимация ожидания
    # ------------------------------------------------------------------ #
    def busy_tick(self) -> None:
        """Кадр анимации «Загружаю реквизиты…» (раз в 300 мс)."""
        if not self.busy:
            return
        dots = "." * (self.busy_phase % 4)
        self.app._cp_status_var.set(f"Загружаю реквизиты из ЕГРЮЛ/ЕГРИП{dots}")
        self.busy_phase += 1
        self.busy_after = self.app.root.after(300, self.app._cp_busy_tick)

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
            self.app.root.after(80, self.app._cp_result_tick)
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
    def party_failed(self, message: str) -> None:
        """Запрос не удался: стоп-анимация, понятная ошибка пользователю."""
        self.busy_stop()
        self.app._cp_status_var.set("")
        ui.alert(self.app.root, "Не удалось заполнить реквизиты", message, "error")

    def party_loaded(self, mode: str, info: dict) -> None:
        """Карточка получена: реквизиты подставляются в поля активного режима."""
        self.busy_stop()
        self.app._cp_status_var.set("")
        filled: list[str] = []

        def put(key: str, value: str):
            if not value:
                return
            widget = self.app._cp_entries[mode].get(key)
            if widget is None:
                return
            widget.delete(0, tk.END)
            widget.insert(0, value)
            filled.append(key)

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
            fact = self.app._cp_entries[mode].get(F.FACT_ADDRESS)
            if fact is not None and not fact.get().strip():
                put(F.FACT_ADDRESS, info["legal_address"])
        else:
            put(F.IP_NAME, info["director"])
            put(F.OGRNIP, info["ogrn"])

        if not filled:
            ui.alert(
                self.app.root, "Данных нет",
                "Сервис вернул карточку без данных для подстановки.", "error",
            )
            return

        self.app._cp_status_var.set("Реквизиты заполнены из ЕГРЮЛ/ЕГРИП")
