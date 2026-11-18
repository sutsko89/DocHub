"""DaData с моком urlopen (п. 7.6): коды ответов, таймаут, разбор, ключ.

Сетевых обращений нет: ``urllib.request.urlopen`` подменяется заглушкой,
поведение сверяется с dadata_client.fetch_party и _parse.
"""
from __future__ import annotations

import io
import json
import urllib.error

import pytest

import dadata_client as dadata
from dadata_client import DadataError, _parse, fetch_party

API_URL = dadata.API_URL


SAMPLE = {"suggestions": [{"data": {
    "name": {
        "full_with_opf": "Общество с ограниченной ответственностью «Ромашка»",
        "short_with_opf": "ООО «Ромашка»",
    },
    "inn": "2540123456", "kpp": "254001001", "ogrn": "1035001234567",
    "address": {"value": "690000, г. Владивосток, ул. Светланская, д. 1"},
    "management": {"name": "Петров Пётр Петрович"},
    "bank": {"name": "ПАО Банк", "bic": "044525225",
             "correspondent_account": "30101810100000000225"},
}}]}


class _FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        return False


def _patch_urlopen(monkeypatch, handler):
    monkeypatch.setattr(dadata.urllib.request, "urlopen", handler)


# ---------- Успешный ответ ----------

def test_200_response_parsed(monkeypatch):
    def handler(request, timeout=None):
        assert json.loads(request.data.decode("utf-8")) == {"query": "2540123456"}
        assert request.get_header("Authorization") == "Token KEY-123"
        return _FakeResponse(json.dumps(SAMPLE).encode("utf-8"))

    _patch_urlopen(monkeypatch, handler)
    parsed = fetch_party("2540123456", "KEY-123")
    assert parsed["inn"] == "2540123456"
    assert parsed["short_name"] == "ООО «Ромашка»"
    assert parsed["director"] == "Петров Пётр Петрович"
    assert parsed["legal_address"].startswith("690000")
    assert parsed["bik"] == "044525225"
    assert parsed["correspondent_account"] == "30101810100000000225"


def test_empty_suggestions_means_not_found(monkeypatch):
    def handler(request, timeout=None):
        return _FakeResponse(json.dumps({"suggestions": []}).encode("utf-8"))

    _patch_urlopen(monkeypatch, handler)
    with pytest.raises(DadataError) as exc_info:
        fetch_party("2540123456", "KEY")
    assert "не найдена" in str(exc_info.value)


# ---------- Коды ошибок ----------

@pytest.mark.parametrize("code,fragment", [(401, "API-ключ"), (429, "лимит")])
def test_http_error_codes(monkeypatch, code, fragment):
    def handler(request, timeout=None):
        raise urllib.error.HTTPError(API_URL, code, "err", None, None)

    _patch_urlopen(monkeypatch, handler)
    with pytest.raises(DadataError) as exc_info:
        fetch_party("2540123456", "KEY")
    assert fragment in str(exc_info.value)


def test_http_error_500_generic_message(monkeypatch):
    def handler(request, timeout=None):
        raise urllib.error.HTTPError(API_URL, 500, "oops", None, None)

    _patch_urlopen(monkeypatch, handler)
    with pytest.raises(DadataError) as exc_info:
        fetch_party("2540123456", "KEY")
    assert "500" in str(exc_info.value)


def test_urlerror_means_no_connection(monkeypatch):
    def handler(request, timeout=None):
        raise urllib.error.URLError("connection refused")

    _patch_urlopen(monkeypatch, handler)
    with pytest.raises(DadataError) as exc_info:
        fetch_party("2540123456", "KEY")
    assert "соединения" in str(exc_info.value)


def test_timeout_has_own_message(monkeypatch):
    def handler(request, timeout=None):
        raise TimeoutError()

    _patch_urlopen(monkeypatch, handler)
    with pytest.raises(DadataError) as exc_info:
        fetch_party("2540123456", "KEY")
    assert "не отвечает" in str(exc_info.value)


# ---------- Разбор ответа ----------

def test_parse_ip_fio_as_director():
    parsed = _parse({
        "fio": {"surname": "Иванов", "name": "Иван", "patronymic": "Иванович"},
        "inn": "254012345678",
    })
    assert parsed["director"] == "Иванов Иван Иванович"
    assert parsed["inn"] == "254012345678"


def test_parse_missing_fields_are_empty():
    parsed = _parse({})
    assert parsed["inn"] == "" and parsed["short_name"] == ""


def test_parse_response_inn_may_differ_from_query():
    """ИНН берётся из ответа — сопоставление с запрошенным делает UI."""
    data = json.loads(json.dumps(SAMPLE))
    data["suggestions"][0]["data"]["inn"] = "7701234567"
    parsed = _parse(data["suggestions"][0]["data"])
    assert parsed["inn"] == "7701234567"


# ---------- Хранение ключа (п. 1.3: ключа нет в коде) ----------

def test_no_key_in_clean_profile(appdata):
    assert dadata.load_api_key() == ""


def test_key_saved_and_cleared(appdata):
    assert dadata.save_api_key("test-key-42") is True
    assert dadata.load_api_key() == "test-key-42"
    assert dadata.clear_api_key() is True
    assert dadata.load_api_key() == ""
