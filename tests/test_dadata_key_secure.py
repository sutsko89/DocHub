"""Шифрование ключа DaData (DPAPI) и ветки TLS-ошибок (аудит 2.0.1).

Формат хранения и миграция старого открытого файла проверяются с
эмулированным DPAPI — тесты идут на любой платформе, включая Linux CI.
Круговое шифрование на настоящем Windows-железе дополнительно закрывается
smoke-тестом (save/load равенство).
"""
from __future__ import annotations

import base64
import json
import ssl
import urllib.error

import pytest

import dadata_client as dadata


# --------------------------------------------------------------------------- #
#  Базовый контракт без DPAPI (как в dev/CI на Linux)
# --------------------------------------------------------------------------- #
def test_roundtrip_without_dpapi(appdata):
    assert dadata.save_api_key("plain-key-42")
    assert dadata.load_api_key() == "plain-key-42"


def test_stored_without_dpapi_is_not_encrypted(appdata):
    dadata.save_api_key("plain-key-42")
    stored = open(dadata.key_storage_path(), encoding="utf-8").read()
    assert not stored.startswith(dadata.KEY_ENC_PREFIX)


# --------------------------------------------------------------------------- #
#  Формат и миграция с эмулированным DPAPI
# --------------------------------------------------------------------------- #
@pytest.fixture()
def fake_dpapi(monkeypatch):
    """«Windows»: _dpapi_* работают на любой платформе теста."""

    def protect(text: str) -> str:
        blob = base64.b64encode(("enc:" + text).encode("utf-8")).decode("ascii")
        return dadata.KEY_ENC_PREFIX + blob

    def unprotect(stored: str):
        if not stored.startswith(dadata.KEY_ENC_PREFIX):
            return None
        try:
            raw = base64.b64decode(
                stored[len(dadata.KEY_ENC_PREFIX):], validate=True,
            ).decode("utf-8")
        except Exception:
            return None
        return raw[4:] if raw.startswith("enc:") else None

    monkeypatch.setattr(dadata, "_dpapi_available", lambda: True)
    monkeypatch.setattr(dadata, "_dpapi_protect", protect)
    monkeypatch.setattr(dadata, "_dpapi_unprotect", unprotect)


def test_saved_key_is_encrypted_with_dpapi(appdata, fake_dpapi):
    assert dadata.save_api_key("secret-key")
    stored = open(dadata.key_storage_path(), encoding="utf-8").read()
    assert stored.startswith(dadata.KEY_ENC_PREFIX)  # открытого ключа на диске нет
    assert "secret-key" not in stored
    assert dadata.load_api_key() == "secret-key"     # круговое шифрование


def test_legacy_plaintext_file_migrates_to_dpapi(appdata, fake_dpapi):
    path = dadata.key_storage_path()
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("old-plain-key\n")  # как писала версия 2.0.1
    assert dadata.load_api_key() == "old-plain-key"
    migrated = open(path, encoding="utf-8").read()
    assert migrated.startswith(dadata.KEY_ENC_PREFIX)  # файл перезаписан шифром


def test_corrupted_encrypted_value_is_not_fatal(appdata, fake_dpapi):
    with open(dadata.key_storage_path(), "w", encoding="utf-8") as fh:
        fh.write(dadata.KEY_ENC_PREFIX + "!!!не-base64!!!")
    assert dadata.load_api_key() == ""  # не падает, ключ считается незаданным


def test_encrypt_failure_falls_back_to_plaintext(appdata, fake_dpapi, monkeypatch):
    # Шифрование сломалось (например, сбой DPAPI) — ключ не теряем.
    monkeypatch.setattr(dadata, "_dpapi_protect", lambda text: None)
    assert dadata.save_api_key("fallback-key")
    stored = open(dadata.key_storage_path(), encoding="utf-8").read()
    assert stored == "fallback-key"
    assert dadata.load_api_key() == "fallback-key"


def test_validation_rules_preserved(appdata):
    assert not dadata.save_api_key("")
    assert not dadata.save_api_key("ключ с пробелом")
    assert not dadata.save_api_key(None)


# --------------------------------------------------------------------------- #
#  TLS-ошибки не маскируются под «нет интернета» (аудит 2.0.1, P3)
# --------------------------------------------------------------------------- #
def _patch_urlopen(monkeypatch, exc: Exception) -> None:
    def raise_exc(_request, timeout=None):
        raise exc

    monkeypatch.setattr(dadata.urllib.request, "urlopen", raise_exc)


def test_cert_error_wrapped_in_urlerror_is_reported_as_tls(appdata, monkeypatch):
    cert_exc = ssl.SSLCertVerificationError(1, "certificate verify failed")
    _patch_urlopen(monkeypatch, urllib.error.URLError(cert_exc))
    with pytest.raises(dadata.DadataError, match="сертификата"):
        dadata.fetch_party("7701234567", "key")


def test_raw_ssl_error_is_reported_as_tls(appdata, monkeypatch):
    _patch_urlopen(monkeypatch, ssl.SSLError("handshake failure"))
    with pytest.raises(dadata.DadataError, match="сертификата"):
        dadata.fetch_party("7701234567", "key")


def test_regular_connection_error_still_reports_no_internet(appdata, monkeypatch):
    _patch_urlopen(monkeypatch, urllib.error.URLError("connection refused"))
    with pytest.raises(dadata.DadataError, match="Нет соединения"):
        dadata.fetch_party("7701234567", "key")
