"""HYG-1 (этап 2, параллельная ветка) — гигиена поставки.

Карточка HYG-1, Т-аудит P2-14 и P2-16:

  * P2-14 — GitHub Actions пинятся по полному SHA коммита, а не по
    major-тегу («actions/checkout@v4»): тег перемещаем, и компрометация
    репозитория экшена или подмена тега доставит чужой код в CI/релиз,
    который собирает и публикует установщик с правами contents:write;
  * P2-16 — в корне репозитория есть LICENSE-файл: без него GitHub не
    определяет лицензию, а установщик и README ссылаются на неё.

Тест читает workflow-файлы как текст — без YAML-парсера, чтобы проверять
именно пины в строках ``uses:``.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
WORKFLOWS = sorted((REPO_ROOT / ".github" / "workflows").glob("*.yml"))

SHA_RE = re.compile(r"uses:\s*([^\s@]+)@([0-9a-f]{40})\b")


def _uses_lines(path: Path) -> list[str]:
    lines = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = raw.strip()
        if stripped.startswith("- uses:") or stripped.startswith("uses:"):
            lines.append(stripped)
    return lines


def test_workflows_exist():
    assert WORKFLOWS, "workflow-файлы не найдены"


@pytest.mark.parametrize("workflow", WORKFLOWS, ids=lambda p: p.name)
def test_p2_14_actions_pinned_to_full_sha(workflow):
    """Каждый внешний экшен пиннут по 40-значному SHA (P2-14)."""
    for line in _uses_lines(workflow):
        body = line.split("uses:", 1)[1].strip()
        if body.startswith("./"):  # локальные экшены репозитория не пинятся
            continue
        match = SHA_RE.search(line)
        assert match, (
            f"{workflow.name}: экшен не пиннут по SHA — «{body}». "
            "Major-теги (@vN) запрещены карточкой HYG-1: обновляйте пин "
            "осознанно, новой ревизией SHA с комментарием версии."
        )


@pytest.mark.parametrize("workflow", WORKFLOWS, ids=lambda p: p.name)
def test_p2_14_pinned_sha_carry_version_comments(workflow):
    """Рядом с SHA — комментарий версии: обновление пина прослеживается."""
    for line in _uses_lines(workflow):
        body = line.split("uses:", 1)[1].strip()
        if body.startswith("./"):
            continue
        if "@" in body:
            after_at = body.split("@", 1)[1]
            sha = after_at.split()[0]
            assert "#" in after_at, (
                f"{workflow.name}: у пина {sha[:12]} нет комментария версии — "
                "добавьте «# vX.Y.Z», чтобы SHA можно было сопоставить релизу"
            )


def test_p2_16_license_file_in_repo_root():
    license_file = REPO_ROOT / "LICENSE"
    assert license_file.exists(), "LICENSE отсутствует в корне (P2-16)"
    text = license_file.read_text(encoding="utf-8")
    assert "как есть" in text.lower(), "LICENSE не отражает условия поставки"
    assert "license.txt" in text, "LICENSE не ссылается на перечень сторонних компонентов"
