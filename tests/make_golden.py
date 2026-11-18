"""Генератор эталонов golden-тестов (п. 7.1).

Запуск из корня репозитория:
    python tests/make_golden.py

Пересоздаёт tests/golden/*.json по текущему коду. Осознанная операция:
изменение шаблона .docx или подстановок должно попадать в отдельный коммит
вместе с обновлёнными эталонами.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from render_helpers import render_six_forms  # noqa: E402

GOLDEN_DIR = Path(__file__).resolve().parent / "golden"


def main() -> int:
    GOLDEN_DIR.mkdir(exist_ok=True)
    signatures = render_six_forms()
    for name, signature in signatures.items():
        path = GOLDEN_DIR / f"{name}.json"
        path.write_text(
            json.dumps(signature, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print(f"эталон: {path.relative_to(GOLDEN_DIR.parent.parent)} "
              f"({len(signature['text'])} символов, таблиц: {len(signature['tables'])})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
