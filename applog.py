"""Логирование приложения: файл с ротацией и без персональных данных.

Правила (см. «План исправлений», п. 2.3):
    * пишем только технические события: старт, версия, ошибки чтения/записи,
      длительность операций;
    * не пишем реквизиты (ИНН, КПП, ОГРН, счета), ФИО, ключ DaData, полный
      текст документов и полные пути к файлам документов.

Логгер настраивается один раз в :func:`setup_logging`; остальные модули
получают его через :func:`get_logger` и не зависят от порядка импорта.
"""

from __future__ import annotations

import logging
import os
import platform
import sys
import time
from contextlib import contextmanager
from logging.handlers import RotatingFileHandler

SETTINGS_DIR_NAME = "DogovorGenerator"
LOGGER_NAME = "ewdochub"
LOG_FILE_NAME = "app.log"
MAX_BYTES = 2 * 1024 * 1024   # 2 МБ на файл
BACKUP_COUNT = 3              # app.log + 3 архивных копии

_configured = False


def log_dir() -> str:
    """Каталог логов в профиле пользователя."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, SETTINGS_DIR_NAME, "logs")


def log_path() -> str:
    return os.path.join(log_dir(), LOG_FILE_NAME)


def setup_logging(version: str = "", level: int = logging.INFO) -> logging.Logger:
    """Настраивает файловый лог с ротацией. Безопасно вызывать повторно."""
    global _configured
    logger = logging.getLogger(LOGGER_NAME)
    if _configured:
        return logger

    logger.setLevel(level)
    logger.propagate = False
    formatter = logging.Formatter(
        "%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler: logging.Handler
    try:
        os.makedirs(log_dir(), exist_ok=True)
        handler = RotatingFileHandler(
            log_path(), maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT,
            encoding="utf-8",
        )
    except OSError:
        # Лог недоступен (read-only профиль, занятый файл) — не мешаем работе.
        handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    _configured = True

    logger.info(
        "Запуск EW.DocHub %s; Python %s; %s",
        version or "(версия не задана)",
        platform.python_version(),
        platform.platform(),
    )
    return logger


def get_logger(name: str = "") -> logging.Logger:
    """Возвращает дочерний логгер приложения (настраивает лог при первом вызове)."""
    if not _configured:
        setup_logging()
    return logging.getLogger(LOGGER_NAME + (f".{name}" if name else ""))


def redact(value: str, visible: int = 2) -> str:
    """Маскирует значение для логов: видны только первые ``visible`` символов."""
    text = str(value or "")
    if not text:
        return ""
    if len(text) <= visible:
        return "*" * len(text)
    return text[:visible] + "*" * (len(text) - visible)


@contextmanager
def timed(operation: str, threshold_ms: int = 500, logger_name: str = ""):
    """Контекстный менеджер: пишет в лог только «медленные» операции.

    Используется для метрик из п. 2.3/5.10: операции дольше 500 мс попадают
    в лог, остальные не создают шума.
    """
    started = time.perf_counter()
    error: BaseException | None = None
    try:
        yield
    except BaseException as exc:  # noqa: BLE001 — логируем и пробрасываем дальше
        error = exc
        raise
    finally:
        elapsed_ms = (time.perf_counter() - started) * 1000
        log = get_logger(logger_name)
        if error is not None:
            log.warning(
                "Операция «%s» прервана ошибкой за %.0f мс: %s",
                operation, elapsed_ms, type(error).__name__,
            )
        elif elapsed_ms >= threshold_ms:
            log.info("Операция «%s» заняла %.0f мс", operation, elapsed_ms)
