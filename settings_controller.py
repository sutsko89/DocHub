"""Страница «Настройки», ключ DaData и папка документов — восьмой контроллер.

Шаг декомпозиции «бог-класса» ContractApp (аудит 2.0.1, P2 «по одному
контроллеру за итерацию»). Класс владеет страницей «Настройки» (тема,
папка документов, ключ DaData), сохранением/удалением ключа DaData
(аудит 2.0.1: сохранённый ключ не подставляется в поле — значение не видно
и не копируется) и всей логикой папки документов: папка по умолчанию
«Документы/EW.DocHub» (п. 6.2), честное решение при недоступной папке
«повторить / выбрать другую / по умолчанию» (п. 2.8), фоновый вариант без
Tk-диалогов для импорта и проверка выхода пути за пределы папки (п. 2.7).

Контроллер получает ссылку на приложение и обращается к общим службам
через него (``self.app.store`` не нужен — здесь только настройки):
``self.app.root``, ``self.app._settings``, ``self.app._persist_settings`` …
Состояние остаётся на приложении с прежними именами
(``_documents_dir_var``, ``_dadata_key_entry``, ``_dadata_key_status_var``,
счётчики ``_docs_dir_retry_count`` / флаги ``_docs_dir_fallback_active`` /
``_docs_dir_dialog_active``): их читают smoke-тест и другие контроллеры —
контроллер владеет кодом, а не данными страницы.

Маршруты подмен и диалоговые примитивы:

* ``ui.alert`` / ``ui.confirm`` — атрибуты общего модуля ui_macos:
  подмена ``m.ui.alert`` видна контроллеру напрямую;
* подтверждение удаления ключа — через ``self.app._choice_dialog`` /
  ``ui.confirm`` (instance-атрибут приложения и атрибут модуля);
* открытие папки кнопкой — через фасад ``self.app._open_file``: его тело
  резолвит модульное ``main.open_file`` на каждый вызов, а smoke-тест
  подменяет именно модульное имя;
* смена темы — через фасад ``self.app._on_theme_change`` (тема меняет
  всё окно, не только страницу настроек);
* немедленное сохранение выбранной папки — через
  ``self.app._persist_settings``: иначе при аварийном завершении выбор
  терялся бы до закрытия программы.

Зависимости от модуля main (``main.SETTINGS_DIR_NAME``, ширины разметки
``main.ENTRY_WIDTH`` / ``main.LEFT_PANEL_WIDTH``) берутся поздним импортом
внутри методов: модуль main импортирует этот контроллер, поэтому
верхнеуровневый импорт дал бы цикл, а к моменту любого вызова метода main
уже полностью инициализирован. Папка «Документы» берётся напрямую из
``win_shell.documents_known_folder`` (п. 9.3). Остальные зависимости —
``ui_macos``, ``win_shell``, ``dadata_client``, ``tkinter`` — импортируются
напрямую.

Поведение один в один повторяет код, вынесенный из main.py версии 2.8.0.
"""

from __future__ import annotations

import os
from tkinter import filedialog
from typing import Any

import customtkinter as ctk
import tkinter as tk

import dadata_client as dadata
import ui_macos as ui
import win_shell


class SettingsController:
    """Страница «Настройки», ключ DaData и папка документов.

    Состояние страницы (виджеты, переменные, счётчики ретраев папки)
    живёт на приложении — контроллер получает его через ``self.app.*``
    на каждый вызов.
    """

    # Сколько раз можно выбрать «Повторить» при недоступной папке
    # документов, прежде чем путь принудительно уводится в папку
    # по умолчанию (защита от бесконечного диалога, п. 2.8).
    MAX_DIR_RETRIES = 2

    def __init__(self, app: Any) -> None:
        self.app = app

    # ------------------------------------------------------------------ #
    #  Страница «Настройки»
    # ------------------------------------------------------------------ #
    def build_settings_page(self):
        """Страница «Настройки»: тема, папка документов, ключ DaData."""
        app = self.app
        import main  # поздний импорт: ширины разметки main.py (см. докстринг)

        _page, left, _right, _textbox = app._make_page(
            "settings", "Настройки",
            "Тема оформления, папка документов и ключ DaData.",
        )

        # -- Тема -------------------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=0, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Тема оформления")
        theme_toggle = ui.ThemeToggle(
            card.body,
            command=app._on_theme_change,
            mode="dark" if ctk.get_appearance_mode().lower() == "dark" else "light",
        )
        card.wide(theme_toggle, pady=(int(2 * app.scale), int(6 * app.scale)))
        # Переключатель — по ширине двух иконок, не растягиваем на всю карточку.
        theme_toggle.grid_configure(sticky="w")

        # -- Папка документов -------------------------------------------- #
        card = ui.Card(left, labels_above=True)
        card.grid(row=1, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("Папка для сохранения документов")
        app._documents_dir_var = tk.StringVar(value=app._settings_flow.documents_dir())
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=app._documents_dir_var,
                font=ui.font(12), text_color=ui.COLORS["text"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 48) * app.scale),
            ),
            pady=(int(2 * app.scale), int(6 * app.scale)),
        )
        buttons = ctk.CTkFrame(card.body, fg_color="transparent")
        card.wide(buttons)
        ui.secondary_button(
            buttons, "Выбрать папку…", app._settings_flow.choose_documents_dir,
            width=int(170 * app.scale), height=int(30 * app.scale),
        ).pack(side="left")
        ui.secondary_button(
            buttons, "Открыть папку", lambda: app._open_file(app._settings_flow.documents_dir()),
            width=int(150 * app.scale), height=int(30 * app.scale),
        ).pack(side="left", padx=(int(10 * app.scale), 0))
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Здесь сохраняются договоры и доп. соглашения, созданные в программе. "
                     "Для каждого контрагента заводится своя папка с его документами.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 48) * app.scale),
            ),
            pady=(int(4 * app.scale), int(2 * app.scale)),
        )

        # Добавление документов перенесено в раздел «Реестр» (п. 4.5).

        # -- DaData ------------------------------------------------------ #
        card = ui.Card(left, labels_above=True)
        card.grid(row=2, column=0, sticky="new", padx=int(4 * app.scale), pady=(0, int(12 * app.scale)))
        card.section("DaData: автозаполнение по ИНН")
        card.wide(
            ctk.CTkLabel(
                card.body,
                text="Ключ не входит в поставку — укажите свой ключ с dadata.ru. "
                     "Без ключа доступен только ручной ввод реквизитов. При "
                     "автозаполнении сервису отправляется только ИНН — публичный "
                     "идентификатор организации/ИП; реквизиты карточки и документы "
                     "никуда не отправляются.",
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 48) * app.scale),
            ),
            pady=(int(2 * app.scale), int(6 * app.scale)),
        )
        app._dadata_key_entry = ui.entry(card.body, width=int(main.ENTRY_WIDTH * app.scale))
        try:
            app._dadata_key_entry.entry.configure(show="•")
        except Exception:
            pass
        # Аудит 2.0.1: сохранённый ключ НЕ подставляется в поле — реальное
        # значение не показывается и не копируется по Ctrl+C. Поле остаётся
        # пустым, пока пользователь не решит сменить ключ; состояние — строкой ниже.
        app._dadata_key_status_var = tk.StringVar(
            value=app._settings_flow.dadata_key_status_text())
        card.wide(
            ctk.CTkLabel(
                card.body, textvariable=app._dadata_key_status_var,
                font=ui.font(11), text_color=ui.COLORS["text_tertiary"],
                anchor="w", justify="left",
                wraplength=int((main.LEFT_PANEL_WIDTH - 48) * app.scale),
            ),
            pady=(int(2 * app.scale), int(4 * app.scale)),
        )
        card.row(
            "API-ключ", app._dadata_key_entry, stretch=True,
            hint="Ключ хранится в профиле пользователя отдельно от реестра.",
        )
        key_buttons = ctk.CTkFrame(card.body, fg_color="transparent")
        card.wide(key_buttons)
        ui.secondary_button(
            key_buttons, "Сохранить ключ", app._settings_flow.save_dadata_key,
            width=int(160 * app.scale), height=int(30 * app.scale),
        ).pack(side="left")
        ui.ghost_button(
            key_buttons, "Удалить ключ", app._settings_flow.clear_dadata_key,
            width=int(130 * app.scale), height=int(28 * app.scale),
        ).pack(side="left", padx=(int(10 * app.scale), 0))

    # ------------------------------------------------------------------ #
    #  Ключ DaData
    # ------------------------------------------------------------------ #
    def dadata_key_status_text(self) -> str:
        """Статус сохранённого ключа — без подстановки самого значения."""
        if dadata.load_api_key():
            if dadata.dpapi_available():
                return "Ключ сохранён и зашифрован (DPAPI, профиль пользователя)."
            return "Ключ сохранён в профиле пользователя."
        return "Ключ не задан."

    def save_dadata_key(self):
        """Сохраняет введённый пользователем ключ DaData."""
        app = self.app
        key = app._dadata_key_entry.get().strip()
        if not key:
            ui.alert(app.root, "Ключ не указан",
                     "Введите API-ключ DaData или оставьте поле пустым.", "error")
            return
        if dadata.save_api_key(key):
            try:
                app._dadata_key_entry.delete(0, "end")
            except Exception:
                pass
            app._dadata_key_status_var.set(app._settings_flow.dadata_key_status_text())
            ui.alert(
                app.root, "Ключ сохранён",
                "Автозаполнение по ИНН будет использовать этот ключ."
                + (" Значение зашифровано (DPAPI)." if dadata.dpapi_available() else ""),
                "success",
            )
        else:
            ui.alert(
                app.root, "Не удалось сохранить ключ",
                "Проверьте права на запись в профиль пользователя:\n"
                f"{dadata.key_storage_path()}", "error",
            )

    def clear_dadata_key(self):
        """Удаляет сохранённый ключ DaData."""
        app = self.app
        if not ui.confirm(
            app.root, "Удалить ключ DaData",
            "Удалить сохранённый ключ? Автозаполнение по ИНН перестанет работать.",
            ok_text="Удалить",
        ):
            return
        dadata.clear_api_key()
        try:
            app._dadata_key_entry.delete(0, "end")
        except Exception:
            pass
        app._dadata_key_status_var.set(app._settings_flow.dadata_key_status_text())
        ui.alert(app.root, "Ключ удалён",
                 "Ключ DaData удалён из профиля пользователя.", "success")

    # ------------------------------------------------------------------ #
    #  Папка документов: выбор и немедленное сохранение
    # ------------------------------------------------------------------ #
    def choose_documents_dir(self):
        """Выбор папки документов с немедленным сохранением настройки."""
        app = self.app
        current = self.documents_dir()
        chosen = filedialog.askdirectory(
            parent=app.root, initialdir=current,
            title="Папка для сохранения договоров и документов",
        )
        if not chosen:
            return
        app._settings["documents_dir"] = chosen
        app._documents_dir_var.set(chosen)
        # Сохраняем сразу, а не только при закрытии: иначе при аварийном
        # завершении выбранная папка теряется.
        if app._persist_settings():
            ui.alert(
                app.root, "Папка сохранена",
                f"Документы будут сохраняться в:\n{chosen}", "success",
            )

    # ------------------------------------------------------------------ #
    #  Папка документов: разрешение пути (п. 2.8, 6.2)
    # ------------------------------------------------------------------ #
    def default_documents_dir(self, interactive: bool = True) -> str:
        """Папка документов по умолчанию: «Документы/EW.DocHub» (п. 6.2).

        Раньше по умолчанию использовался скрытый %APPDATA%\\DogovorGenerator\\documents:
        пользователь не видел, где лежат его файлы. Теперь — видимая папка
        «Документы», которая не исчезает при переустановке системы и попадает
        в обычные резервные копии. Если «Документы» недоступны — возврат
        к прежнему месту в профиле пользователя.
        """
        app = self.app
        import main  # поздний импорт: константы main.py (см. докстринг)

        base = os.path.join(win_shell.documents_known_folder(), "EW.DocHub")
        try:
            os.makedirs(base, exist_ok=True)
            return base
        except OSError as exc:
            app._log.error(
                "Папка документов по умолчанию недоступна: %s", type(exc).__name__,
            )
            if interactive:
                ui.alert(
                    app.root, "Папка документов недоступна",
                    "Не удалось создать папку документов по умолчанию:\n"
                    f"{base}\n\n{exc}\n\n"
                    "Выберите другую папку в «Настройках» или проверьте права доступа.",
                    "error",
                )
        base = os.path.join(
            os.environ.get("APPDATA") or os.path.expanduser("~"),
            main.SETTINGS_DIR_NAME, "documents",
        )
        try:
            os.makedirs(base, exist_ok=True)
        except OSError:
            pass
        return base

    def resolve_unavailable_documents_dir(self, custom: str, exc: OSError) -> str:
        """Спрашивает, куда сохранять, если выбранная папка недоступна (п. 2.8)."""
        app = self.app
        app._log.warning("Папка документов недоступна: %s", type(exc).__name__)
        if app._docs_dir_fallback_active or app._docs_dir_dialog_active:
            return self.default_documents_dir()
        if app._docs_dir_retry_count >= self.MAX_DIR_RETRIES:
            app._docs_dir_fallback_active = True
            return self.default_documents_dir()

        app._docs_dir_dialog_active = True
        try:
            choice = app._choice_dialog(
                "Папка документов недоступна",
                f"Не удалось открыть папку:\n{custom}\n\n{exc}\n\n"
                "Куда сохранять новые документы?",
                [
                    ("default", "Папка по умолчанию", False),
                    ("choose", "Выбрать другую", False),
                    ("retry", "Повторить", True),
                ],
            )
        finally:
            app._docs_dir_dialog_active = False

        if choice == "retry":
            app._docs_dir_retry_count += 1
            return self.documents_dir(interactive=True)
        if choice == "choose":
            chosen = filedialog.askdirectory(
                parent=app.root, title="Выберите папку для документов",
            )
            if chosen:
                app._settings["documents_dir"] = chosen
                app._docs_dir_retry_count = 0
                app._docs_dir_fallback_active = False
                app._persist_settings()
                return chosen
        app._docs_dir_fallback_active = True
        return self.default_documents_dir()

    def documents_dir(self, interactive: bool = True) -> str:
        """Папка для готовых документов: выбранная в настройках либо %APPDATA%.

        Недоступная папка больше не подменяется молча (п. 2.8): пользователь
        выбирает «повторить / выбрать другую / по умолчанию» и видит, куда
        реально сохранится документ.
        """
        app = self.app
        custom = str(app._settings.get("documents_dir") or "").strip()
        if custom:
            try:
                os.makedirs(custom, exist_ok=True)
                app._docs_dir_retry_count = 0
                return custom
            except OSError as exc:
                if interactive:
                    return self.resolve_unavailable_documents_dir(custom, exc)
                app._log.warning(
                    "Папка документов недоступна (фоновая задача): %s",
                    type(exc).__name__,
                )
        return self.default_documents_dir(interactive=interactive)

    def documents_dir_for_import(self) -> str:
        """Папка документов для фонового импорта: без Tk-диалогов.

        Импорт выполняется в отдельном потоке, где нельзя показывать окна.
        Если выбранная пользователем папка недоступна, возвращается ошибка —
        она попадёт в отчёт импорта, а не приведёт к молчаливой подмене пути.
        """
        app = self.app
        custom = str(app._settings.get("documents_dir") or "").strip()
        if custom:
            try:
                os.makedirs(custom, exist_ok=True)
                return custom
            except OSError as exc:
                raise OSError(
                    "Папка документов недоступна:\n"
                    f"{custom}\n{exc}\n\n"
                    "Выберите другую папку в «Настройках» и повторите импорт."
                ) from exc
        return self.default_documents_dir(interactive=False)

    @staticmethod
    def is_inside_directory(base: str, path: str) -> bool:
        """Проверяет, что ``path`` не выходит за пределы ``base`` (п. 2.7)."""
        try:
            base_norm = os.path.normcase(os.path.normpath(base))
            path_norm = os.path.normcase(os.path.normpath(path))
            return os.path.commonpath([base_norm, path_norm]) == base_norm
        except (ValueError, OSError):
            return False
