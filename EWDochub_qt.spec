# -*- mode: python ; coding: utf-8 -*-
# ONEDIR-сборка Qt-версии EW.DocHub (этап M7, финализация; измерительная
# версия каркаса — M2, критерий К-3). Отличия от EWDochub.spec (Tk-сборка):
# точка входа main_qt.py, PySide6 включается (Tk-сборка его исключает), стек
# Tk/CustomTkinter исключается. Qt-плагины (platforms/styles/imageformats)
# PyInstaller 6.x подтягивает хуками PySide6 по фактическим импортам каркаса.
#
# Полный набор шаблонов .docx (M7): договор/доп. соглашение (M6) + спецификация
# — как у Tk-сборки; app_icon.ico для окна и установщика.
#
# Установщик: installer_qt.iss (те же AppMutex/AppId — обновление Tk 2.28.2 → Qt
# в ту же папку, критерий CP2). Ассет релиза: EW.DocHub-<версия>-qt-setup.exe
# (суффикс флейвора — И.2 плана).
#
# Сборка: pyinstaller EWDochub_qt.spec --noconfirm

import re
import sys
from pathlib import Path

# Каталог spec-файла: PyInstaller исполняет spec через exec без __file__
# (PyInstaller 6.x), официальный глобал — SPECPATH (та же гигиена, что в
# EWDochub.spec).
try:
    _SPEC_DIR = Path(SPECPATH)  # noqa: F821 — глобал PyInstaller
except NameError:
    try:
        _SPEC_DIR = Path(__file__).resolve().parent
    except NameError:
        _SPEC_DIR = Path.cwd()

# VersionInfo для exe — единый источник версии app_common.APP_VERSION
# (тот же контракт, что у Tk-сборки; строка чисто числовая «X.Y.Z»).
_match = re.search(
    r'APP_VERSION\s*=\s*"(\d+)\.(\d+)\.(\d+)"',
    (_SPEC_DIR / "app_common.py").read_text(encoding="utf-8"),
)
if not _match:
    raise SystemExit("EWDochub_qt.spec: не удалось прочитать APP_VERSION из app_common.py")
APP_VERSION = ".".join(_match.groups())
FILEVERS = tuple(int(group) for group in _match.groups()) + (0,)

exe_version = None
if sys.platform == "win32":
    try:
        from PyInstaller.utils.win32.versioninfo import (
            VSVersionInfo, FixedFileInfo, StringFileInfo, StringTable,
            VarFileInfo, StringStruct, VarStruct,
        )

        exe_version = VSVersionInfo(
            ffi=FixedFileInfo(
                filevers=FILEVERS,
                prodvers=FILEVERS,
                mask=0x3F,
                flags=0x0,
                OS=0x40004,
                fileType=0x1,
                subtype=0x0,
                date=(0, 0),
            ),
            kids=[
                StringFileInfo([
                    StringTable(
                        "041904B0",
                        [
                            StringStruct("CompanyName", "EW.DocHub"),
                            StringStruct("FileDescription",
                                         "EW.DocHub — контрагенты и документы"),
                            StringStruct("FileVersion", APP_VERSION),
                            StringStruct("InternalName", "EW.DocHub-qt"),
                            StringStruct("LegalCopyright", "© EW.DocHub"),
                            StringStruct("OriginalFilename", "EW.DocHub-qt.exe"),
                            StringStruct("ProductName", "EW.DocHub"),
                            StringStruct("ProductVersion", APP_VERSION),
                        ],
                    )
                ]),
                VarFileInfo([VarStruct("Translation", [1049, 1200, 1033, 1200])]),
            ],
        )
    except Exception:
        exe_version = None  # сборка без versioninfo — лучше, чем падение сборки

a = Analysis(
    ['main_qt.py'],
    pathex=[],
    binaries=[],
    # Полный набор шаблонов печатных форм (M7-финализация; как у Tk-сборки):
    # M3 — спецификация, M6 — договор/доп. соглашение; .docx читаются через
    # services.resource_path (заморозка поддерживается).
    datas=[
        ('app_icon.ico', '.'),
        ('dogovor_template.docx', '.'),
        ('dogovor_template_ip.docx', '.'),
        ('additional_agreement.docx', '.'),
        ('additional_agreement_ip.docx', '.'),
        ('spec_template.docx', '.'),
        ('spec_template_ip.docx', '.'),
    ],
    hiddenimports=[
        'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Исключаем стек Tk-сборки и тяжёлые пакеты, которых у Qt-каркаса нет:
    # замер размера/старта (К-3) должен быть честным для M2-каркаса.
    excludes=[
        'tkinter', 'customtkinter', 'darkdetect', 'Pillow',
        'numpy', 'matplotlib', 'pandas', 'scipy', 'sympy',
        'IPython', 'jupyter', 'notebook',
        'PyQt5', 'PyQt6', 'PySide2',
        'pytest', 'sphinx', 'setuptools', 'pkg_resources', 'pip', 'wheel',
        'Pygments', 'cryptography',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='EW.DocHub-qt',
    version=exe_version,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='EW.DocHub-qt',
)
