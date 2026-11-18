"""Фикстуры pytest-набора EW.DocHub.

Тесты идут без интерфейса: APPDATA изолируется в tmp-каталог, используются
сервисный слой (services.py) и чистый конвейер печатных форм
(document_builder.py). GUI не создаётся — Xvfb нужен только smoke_test.py.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


@pytest.fixture()
def appdata(tmp_path, monkeypatch):
    """Изолированный «профиль пользователя» (%APPDATA%) на один тест."""
    path = tmp_path / "appdata"
    path.mkdir()
    monkeypatch.setenv("APPDATA", str(path))
    return path


@pytest.fixture()
def store(appdata):
    import storage

    return storage.DataStore()


@pytest.fixture()
def ooo_fields():
    from render_helpers import OOO_FIELDS

    return dict(OOO_FIELDS)


@pytest.fixture()
def ip_fields():
    from render_helpers import IP_FIELDS

    return dict(IP_FIELDS)
