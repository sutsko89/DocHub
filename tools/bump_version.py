#!/usr/bin/env python3
"""Единый бамп версии EW.DocHub (аудит 2.0.1, P0 «версии разъехались»).

Версия живёт в двух местах, но синхронизируется одним скриптом:
  * app_common.py  — APP_VERSION (единый источник; EWDochub.spec читает его
                     напрямую при сборке; с этапа M1 миграции источник
                     здесь — раньше был main.py);
  * installer.iss  — MyAppVersion.

Скрипт меняет оба места, обновляет README (заголовок и имена setup.exe)
и проверяет, что spec не держит собственной копии версии.

Использование:
    python tools/bump_version.py 2.1.1
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

TARGETS = [
    (
        REPO_ROOT / "app_common.py",
        r'APP_VERSION = "[\d.]+"',
        'APP_VERSION = "{version}"',
    ),
    (
        REPO_ROOT / "installer.iss",
        r'#define MyAppVersion "[\d.]+"',
        '#define MyAppVersion "{version}"',
    ),
]

# Файлы, где собственной копии версии быть НЕ должно (проверка).
FORBIDDEN_LITERALS = [
    (
        REPO_ROOT / "EWDochub.spec",
        r'APP_VERSION\s*=\s*"\d+\.\d+\.\d+"',
        "EWDochub.spec: версия должна читаться из app_common.py, а не дублироваться",
    ),
]

README_SETUP_RE = re.compile(r"EW\.DocHub-[\d.]+-setup")
VERSION_RE = re.compile(r"^\d+\.\d+\.\d+$")


def main() -> int:
    if len(sys.argv) != 2 or not VERSION_RE.match(sys.argv[1]):
        print("Использование: python tools/bump_version.py <X.Y.Z>")
        return 2
    version = sys.argv[1]

    for path, pattern, replacement in TARGETS:
        text = path.read_text(encoding="utf-8")
        new_text, count = re.subn(pattern, replacement.format(version=version), text)
        if count != 1:
            print(f"ОШИБКА: {path.name}: шаблон найден {count} раз (ожидался 1)")
            return 1
        path.write_text(new_text, encoding="utf-8")
        print(f"обновлено: {path.name} → {version}")

    # README: заголовок и имена setup.exe — для информации, не источник правды.
    readme = REPO_ROOT / "README.md"
    text = readme.read_text(encoding="utf-8")
    text = re.sub(r"# EW\.DocHub \(версия [\d.]+\)",
                  f"# EW.DocHub (версия {version})", text)
    text = README_SETUP_RE.sub(f"EW.DocHub-{version}-setup", text)
    readme.write_text(text, encoding="utf-8")
    print(f"обновлено: README.md (заголовок, имена setup.exe) → {version}")

    ok = True
    for path, pattern, message in FORBIDDEN_LITERALS:
        if re.search(pattern, path.read_text(encoding="utf-8")):
            print(f"ПРЕДУПРЕЖДЕНИЕ: {message}")
            ok = False

    print("Проверка сборочных файлов:")
    print("  python -m pytest tests/ && python smoke_test.py")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
