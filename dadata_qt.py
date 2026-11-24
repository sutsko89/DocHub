"""«Заполнить по ИНН» на Qt — жизненный цикл запроса DaData (этап M4, U8-полностью).

Qt-перенос ``dadata_controller.DadataController``: guard повторных кликов,
busy-анимация, фоновый HTTP-воркер, опрос очереди результата из главного
потока и подстановка реквизитов в поля формы. Сеть — без изменений
``dadata_client.fetch_party`` (SEC-1 allowlist не расширяется, TLS/таймауты
не трогаются); рабочий поток кладёт результат в ``queue.Queue``, опрос —
именованные poll-таймеры каркаса (QTimer, не QThread — «переписать заодно»
запрещено планом, риск 6).

Контракты Tk-версии, переносимые дословно:

* guard busy: повторные клики при активном запросе игнорируются
  (аудит 2.0.1) — раньше два параллельных воркера гонялись за полями;
* ИНН валидируется по режиму карточки: 10 цифр для ООО, 12 для ИП;
* без API-ключа — понятная ошибка, запрос не запускается;
* подстановка реквизитов выполняется ТОЛЬКО в главном потоке (result_tick
  из poll-таймера каркаса); рабочий поток не трогает UI;
* тихий отказ, когда окно-владелец закрылось, пока запрос летел (U8):
  RuntimeError уничтоженного QLineEdit/окна означает, что диалог уже
  некому показывать, — состояние формы не важно;
* статус запроса (анимация/результат) живёт на цели (:class:`DadataTargetQt`),
  а не на контроллере — как Tk DadataTarget: у страницы добавления и окна
  правки свои строки статуса.

Снимок адресата (:class:`DadataTargetQt`) — аналог Tk DadataTarget: поля
формы (адаптеры с ``.get()``/``.set_text()``), фиксированный режим
"ip"/"ooo", окно-владелец диалогов. Фабрика потока передаётся на каждый
вызов (``thread_factory``) — подмена в тестах действует в момент запуска
воркера, как в Tk.
"""

from __future__ import annotations

import queue
import threading

import dadata_client as dadata
from document_builder import F

# Кадр busy-анимации и опрос результата — интервалы Tk-версии.
BUSY_INTERVAL_MS = 300
RESULT_INTERVAL_MS = 80

BUSY_BASE_TEXT = "Загружаю реквизиты из ЕГРЮЛ/ЕГРИП"


class DadataTargetQt:
    """Адресат «Заполнить по ИНН»: поля формы, режим, статус, окно-владелец.

    ``entries`` — словарь полей РЕШАЮЩЕГО режима (адаптеры QtEntry /
    QtOptionEntry); ``mode`` — фиксированный режим подстановки
    ("ip"/"ooo"); ``status_setter`` — callable(str) для строки статуса
    (уничтоженное окно — RuntimeError глушится: статус уже некому
    показывать); ``root`` — виджет-владелец диалогов (у окна правки
    диалоги модальны к нему, а не к главному окну).
    """

    def __init__(self, root, entries: dict, mode: str, status_setter=None) -> None:
        self.root = root
        self.entries = entries
        self.mode = mode
        self.status_setter = status_setter

    def alive(self) -> bool:
        """Жив ли адресат: виджет-владелец не уничтожен.

        Аналог Tk ``winfo_exists``: закрытое (скрытое), но не уничтоженное
        окно живо — как в Tk. Уничтоженный C++-объект Qt поднимает
        RuntimeError при любом доступе — это и есть признак смерти.
        """
        try:
            _ = self.root.metaObject()
            return True
        except RuntimeError:
            return False

    def set_status(self, text: str) -> None:
        if self.status_setter is None:
            return
        try:
            self.status_setter(text)
        except RuntimeError:
            pass  # окно-владелец уничтожено — статус уже некому показывать


class DadataFlowQt:
    """Жизненный цикл «Заполнить по ИНН»: guard → HTTP → подстановка."""

    def __init__(self, app) -> None:
        self.app = app
        # Состояние активного запроса (как Tk DadataController).
        self.busy = False
        self.busy_phase = 0
        self.result_queue: "queue.Queue | None" = None
        self._target: "DadataTargetQt | None" = None

    # ------------------------------------------------------------------ #
    #  Запрос
    # ------------------------------------------------------------------ #
    def fill_from_inn(self, thread_factory=None,
                      target: "DadataTargetQt | None" = None) -> None:
        """Запрашивает реквизиты по ИНН из активной карточки (ООО или ИП).

        ``target=None`` — страница добавления контрагента (цель строится
        из её состояния); окно правки передаёт собственную цель (снимок).
        Guard повторных кликов стоит в начале: пока запрос выполняется,
        новый воркер не запускается. ``thread_factory`` разрешается в
        момент вызова (по умолчанию ``threading.Thread``): подмена в
        тестах должна действовать в момент запуска воркера.
        """
        if self.busy:
            return
        self._target = target or self.app.counterparty_flow.dadata_target()
        target = self._target
        mode = target.mode
        widget = target.entries.get(F.INN)
        inn = widget.get().strip() if widget is not None else ""
        expected = 12 if mode == "ip" else 10
        if not inn.isdigit() or len(inn) != expected:
            self.app.alert(
                self._alert_parent(), "Проверьте ИНН",
                f"ИНН должен содержать {expected} цифр.", "error",
            )
            self._target = None
            return

        api_key = dadata.get_api_key()
        if not api_key:
            self.app.alert(
                self._alert_parent(), "Нет API-ключа",
                "Ключ DaData не задан. Укажите его на странице «Настройки» "
                "или заполните реквизиты вручную.", "error",
            )
            self._target = None
            return

        self.busy = True
        self.busy_phase = 0
        # Очередь результата + опрос из главного потока (контракт раздела
        # Д плана): рабочий поток не трогает UI. Именованные poll-таймеры
        # каркаса: останов при закрытии окна — _stop_timers/abort().
        result_queue: queue.Queue = queue.Queue()
        self.result_queue = result_queue
        self.busy_tick()
        self.app.add_poll_timer("dadata_busy", BUSY_INTERVAL_MS, self.busy_tick)
        self.app.add_poll_timer("dadata_result", RESULT_INTERVAL_MS, self.result_tick)
        factory = thread_factory or threading.Thread
        factory(
            target=self.fetch_worker,
            args=(inn, api_key, mode, result_queue),
            daemon=True,
        ).start()

    def abort(self) -> None:
        """Полный сброс активного запроса (закрытие окна каркаса)."""
        self.busy_stop()

    # ------------------------------------------------------------------ #
    #  Анимация ожидания
    # ------------------------------------------------------------------ #
    def busy_tick(self) -> None:
        """Кадр анимации «Загружаю реквизиты…» (раз в 300 мс)."""
        if not self.busy:
            return
        if self._target is None:
            self.busy_stop()
            return
        dots = "." * (self.busy_phase % 4)
        self._target.set_status(f"{BUSY_BASE_TEXT}{dots}")
        self.busy_phase += 1

    def busy_stop(self) -> None:
        """Останавливает анимацию и цикл опроса результата."""
        self.busy = False
        try:
            self.app.stop_poll_timer("dadata_busy")
            self.app.stop_poll_timer("dadata_result")
        except RuntimeError:
            pass  # таймеры уже уничтожены вместе с окном
        self.result_queue = None  # цикл опроса останавливается
        self._target = None

    # ------------------------------------------------------------------ #
    #  Фоновый воркер и опрос результата
    # ------------------------------------------------------------------ #
    def fetch_worker(self, inn: str, api_key: str, mode: str,
                     result_queue: "queue.Queue") -> None:
        """Рабочий поток DaData: только HTTP и очередь — никаких вызовов UI."""
        try:
            info = dadata.fetch_party(inn, api_key)
        except Exception as exc:
            result_queue.put(("error", str(exc)))
            return
        result_queue.put(("ok", mode, info))

    def result_tick(self) -> None:
        """Опрашивает очередь результата DaData из главного потока."""
        result_queue = self.result_queue
        if result_queue is None:
            return  # запрос завершён или окно закрывается
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            return
        kind = item[0]
        if kind == "error":
            self.party_failed(item[1])
        elif kind == "ok":
            self.party_loaded(item[1], item[2])

    # ------------------------------------------------------------------ #
    #  Подстановка реквизитов
    # ------------------------------------------------------------------ #
    def _active_target(self, mode: str | None = None) -> "DadataTargetQt":
        """Активная цель запроса; без цели — страница добавления.

        Прямые вызовы party_loaded/party_failed/busy_tick (тесты) без
        цели адресуют страницу приложения. Если при этом передан явный
        ``mode`` (подстановка знает режим из результата запроса), поля
        берутся по нему, а не по типу, выбранному на странице.
        """
        if self._target is not None:
            return self._target
        if mode is not None:
            return self.app.counterparty_flow.dadata_target(mode)
        return self.app.counterparty_flow.dadata_target()

    def _alert_parent(self):
        """Окно для диалогов: цель (окно правки), если жива; иначе главное."""
        target = self._target
        win = getattr(target, "root", None) if target is not None else None
        if win is not None and not target.alive():
            return self.app.window
        return win if win is not None else self.app.window

    def party_failed(self, message: str) -> None:
        """Запрос не удался: стоп-анимация, понятная ошибка пользователю."""
        target = self._target
        parent = self._alert_parent()
        self.busy_stop()
        if target is not None:
            target.set_status("")
        self.app.alert(parent, "Не удалось заполнить реквизиты", message, "error")

    def party_loaded(self, mode: str, info: dict) -> None:
        """Карточка получена: реквизиты подставляются в поля активного режима."""
        target = self._active_target(mode)
        parent = self._alert_parent()
        self.busy_stop()
        filled: list[str] = []

        def put(key: str, value: str) -> None:
            if not value:
                return
            adapter = target.entries.get(key)
            if adapter is None:
                return
            adapter.set_text(value)
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
        except RuntimeError:
            # Окно правки закрылось, пока запрос был в полёте (U8): тихий
            # отказ — диалог уже некому показывать, состояние формы не важно.
            return

        if not filled:
            self.app.alert(
                parent, "Данных нет",
                "Сервис вернул карточку без данных для подстановки.", "error",
            )
            return

        target.set_status("Реквизиты заполнены из ЕГРЮЛ/ЕГРИП")
        # V5/UX-B: реквизиты подставлены программно — нажатий клавиш не было.
        # «Грязное» состояние страницы честно отражает несохранённые
        # реквизиты (baseline не сбрасывается). Обновление точек навигации
        # из Tk не переносится: в Qt-каркасе M4 точек нет (заметки этапа).
