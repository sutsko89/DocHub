r"""Именованные мьютексы single_instance (аудит 2.0.1, P2 «Local\ не защищает RDP»).

Проверяются чистые функции имён и referенс из installer.iss; реальный Win32
мьютекс работает только на Windows (покрывается smoke-тестом на раннере).
"""
from __future__ import annotations

import sys

import pytest

import single_instance


def test_mutex_name_is_global():
    name = single_instance.mutex_name()
    assert name.startswith("Global\\"), "Local\\ не защищает вторую сессию (RDP)"


def test_mutex_name_contains_product():
    assert "EWDocHub" in single_instance.mutex_name()


def test_mutex_name_includes_sid_when_available(monkeypatch):
    monkeypatch.setattr(single_instance, "_user_sid", lambda: "S-1-5-21-1004663348")
    name = single_instance.mutex_name()
    assert "S-1-5-21-1004663348" in name
    assert name.startswith("Global\\")


def test_user_sid_is_empty_off_windows():
    if sys.platform == "win32":
        pytest.skip("на Windows SID настоящий и не пуст")
    assert single_instance._user_sid() == ""


@pytest.mark.skipif(sys.platform == "win32",
                    reason="на Windows тест создавал бы реальные мьютексы сессии")
def test_acquire_release_cycle_off_windows():
    lock = single_instance.InstanceLock()
    assert lock.acquire() is True
    lock.release()


def test_installer_mutex_constant():
    assert single_instance.INSTALLER_MUTEX_NAME == "Global\\EWDocHub_InstallerMutex"


def test_installer_iss_references_installer_mutex():
    with open("installer.iss", encoding="utf-8") as fh:
        source = fh.read()
    assert "AppMutex=Global\\EWDocHub_InstallerMutex;Local\\EWDocHub_InstallerMutex" in source


def test_local_fallback_names_cover_global_refusal():
    candidates = single_instance.InstanceLock._mutex_candidates()
    assert "Local\\EWDocHub_InstallerMutex" in candidates
    assert single_instance.INSTALLER_MUTEX_NAME in candidates
