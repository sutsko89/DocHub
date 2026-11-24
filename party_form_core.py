"""GUI-free ядро формы карточки контрагента (этап M4 миграции Tk → PySide6).

Один источник логики «чтение/снимок/сбой» для обеих сборок — аналог
``spec_table_core`` для таблицы позиций (M3): дрейф поведения между
Tk- и Qt-карточками исключён по построению. Виджеты обеих сборок
адресуются через протокол ``.get()``:

* Tk-виджеты (``ui.entry``/``CTkOptionMenu``) имеют ``.get()`` нативно;
* Qt-страница оборачивает ``QLineEdit`` в адаптер с ``.get()``
  (counterparty_page_qt.QtEntry) — та же функция ядра читает обе сборки.

Контракты, перенесённые дословно из counterparty_controller (аудит 2.0.1):

* сбой чтения поля не превращается в пустую строку (D4): сбойные ключи
  не попадают в ``data`` — вызывающий код решает: блокировать сохранение
  (обязательные поля) или явно спросить пользователя (необязательные);
* снимок полей помечает сбой маркером, не пустотой: «поле сломалось»
  никогда не выглядит как «пользователь стёр значение»;
* подстановка прежних значений сбойных необязательных полей — только
  после явного согласия пользователя; молчаливой замены нет.

Модуль не импортирует Tk и Qt — safe для Qt-сборки (EWDochub_qt.spec
исключает Tk-стек) и для Tk-сборки.
"""

from __future__ import annotations

from document_builder import F

# Маркер в снимке полей: сбой чтения виджета отличается от пустой строки,
# поэтому «поле сломалось» никогда не выглядит как «пользователь стёр значение».
READ_ERROR_MARKER = "\x00поле-недоступно\x00"

# U4 (UX-A): подсказки формата в полях с жёстким форматом. В Tk показываются
# подложкой в пустом поле (ui.TextField.set_placeholder) — не в значении:
# get() пустого поля всегда даёт пустую строку, подсказка не попадает
# в карточку и не «путается» с введёнными данными. В Qt — то же самое
# через QLineEdit.setPlaceholderText. Ключи — константы F (состав полей
# «не трогаем», план M4) — дрейфа между сборками нет.
FORMAT_HINTS: dict[str, str] = {
    F.INN: "10 или 12 цифр",
    F.KPP: "9 цифр",
    F.OGRN: "13 цифр",
    F.OGRNIP: "15 цифр",
    F.ACCOUNT: "20 цифр",
    F.CORR_ACCOUNT: "20 цифр",
    F.BIK: "9 цифр",
}

# Один источник маски допустимых символов телефона (сервис валидирует им же,
# services.validate_phone; Qt-валидатор живого ввода берёт ту же строку).
PHONE_ALLOWED_PATTERN = r"^[\d\+\-\(\)\s\/\.]*$"


def read_entry_values(entries, log=None) -> tuple[dict, list[str]]:
    """Читает значения полей; сбой чтения не превращается в пустую строку (D4).

    Возвращает ``(data, failed)``: сбойные ключи в ``data`` не попадают —
    вызывающий код обязан решить: блокировать сохранение (обязательные
    поля) или явно спросить пользователя (необязательные). Поведение
    совпадает с ``collect()`` страницы карточки: тип исключения — в журнал.
    """
    data: dict = {}
    failed: list[str] = []
    for key, widget in entries.items():
        try:
            data[key] = widget.get().strip()
        except Exception as exc:
            failed.append(key)
            if log is not None:
                log.warning(
                    "Не удалось прочитать поле карточки «%s»: %s",
                    key, type(exc).__name__,
                )
    return data, failed


def form_signature(entries) -> dict:
    """Снимок значений полей: сбой чтения помечается маркером, не пустотой."""
    signature: dict = {}
    for key, widget in entries.items():
        try:
            signature[key] = widget.get().strip()
        except Exception:
            signature[key] = READ_ERROR_MARKER
    return signature


def split_read_failures(failed: list[str], required_keys) -> tuple[list[str], list[str]]:
    """Делит сбойные поля на обязательные (блокируют сохранение) и необязательные.

    Пустое поле ввода — не сбой: блокируется только сохранение, когда
    сломано чтение виджета обязательного поля. Сбой в необязательном поле
    даёт пользователю явный выбор — молчаливой замены нет (аудит 2.0.1).
    """
    blocking = [key for key in failed if key in required_keys]
    optional = [key for key in failed if key not in required_keys]
    return blocking, optional


def merge_saved_values(data: dict, optional: list[str], saved: dict) -> dict:
    """Подставляет прежние значения сбойных необязательных полей.

    Вызывается только после явного согласия пользователя («Сохранить
    с прежними значениями») — молчаливой замены нет. Ключи, которых нет
    среди сбойных, не трогаются.
    """
    for key in optional:
        if key not in data:
            data[key] = str(saved.get(key, "")).strip()
    return data
