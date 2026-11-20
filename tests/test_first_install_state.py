r"""Контракт поставки 2.16.0: первая установка и обновление.

Проверка релиза v2.15.0 на реальном ПК подняла два требования пользователя:

1. «При первой установке приложения на ПК реестр должен быть чистым,
   ключа DaData быть не должно» — установщик возит только сборку PyInstaller,
   ни одной пользовательской базы/ключа в неё не входит; всё создаётся в
   %APPDATA%\DogovorGenerator во время работы. Тесты фиксируют состояние
   свежего профиля как регрессионный барьер: пустой реестр, отсутствие
   файла ключа и settings.json.
2. «При обновлении данные остаются» — обновление поверх прошлой версии не
   затрагивает профиль: реестр, настройки и ключ DaData переживают
   установку новой версии (пользователь не теряет контрагентов и документы).
   Полная очистка профиля — только явное решение пользователя при удалении
   программы (вопрос деинсталлятора, installer.iss [Code]).
"""
from __future__ import annotations

from pathlib import Path

import dadata_client as dadata
import main
import storage


# --------------------------------------------------------------------------- #
#  1. Первая установка: свежий профиль чист
# --------------------------------------------------------------------------- #
def test_fresh_install_registry_is_empty(appdata):
    store = storage.DataStore()
    assert store.active_counterparties() == []
    assert store.trashed_counterparties() == []
    assert store.active_documents() == []
    assert store.trashed_documents() == []
    assert store.trash_count() == 0


def test_fresh_install_without_dadata_key(appdata):
    assert dadata.load_api_key() == ""
    # файл ключа не создаётся, пока пользователь не введёт его в «Настройках»
    assert not Path(dadata.key_storage_path()).exists()
    # legacy-каталог прошлой версии тоже не создаётся
    assert not Path(dadata.legacy_key_storage_path()).exists()


def test_fresh_install_without_settings_file(appdata):
    # settings.json появляется при первом сохранении настроек, а не при старте
    assert not Path(main.settings_path()).exists()
    assert main.load_settings() == {}


def test_fresh_install_creates_only_db(appdata):
    profile = appdata / "DogovorGenerator"
    storage.DataStore()
    # из файлов профиля — только база данных; ключей и настроек нет
    names = {p.name for p in profile.iterdir()}
    assert "database.db" in names
    assert "dadata_key.txt" not in names
    assert "settings.json" not in names


# --------------------------------------------------------------------------- #
#  2. Обновление поверх: данные пользователя сохраняются
# --------------------------------------------------------------------------- #
def test_update_keeps_registry_settings_and_key(appdata):
    store = storage.DataStore()
    party = store.add_counterparty(
        "ooo", {"name": "ООО Ромашка", "inn": "7701234567"}
    )
    store.add_document("contract", party["id"], "12", "20.11.2026", {"sum": "150000"})
    assert dadata.save_api_key("test-key-42")
    assert main.write_settings({"theme": "dark", "auto_update_check": False})

    # «Установка новой версии поверх»: тот же профиль, новое хранилище
    reinstalled = storage.DataStore()
    parties = reinstalled.active_counterparties()
    assert len(parties) == 1
    assert parties[0]["fields"]["name"] == "ООО Ромашка"
    assert len(reinstalled.active_documents()) == 1
    assert dadata.load_api_key() == "test-key-42"
    assert main.load_settings() == {"theme": "dark", "auto_update_check": False}
