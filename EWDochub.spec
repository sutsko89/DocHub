# -*- mode: python ; coding: utf-8 -*-
# Сборка ONEDIR приложения EW.DocHub: exe лежит в папке вместе с библиотеками.
# Запуск почти мгновенный — ничего не распаковывается во временную папку.
# Итоговая папка dist/EW.DocHub упаковывается установщиком Inno Setup (installer.iss).
# Сборка: pyinstaller EWDochub.spec --noconfirm

import re
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Каталог spec-файла: PyInstaller исполняет spec через exec без __file__
# (PyInstaller 6.x), официальный глобал — SPECPATH. Fallback — каталог
# текущего файла при выполнении вне PyInstaller (проверки линтером).
try:
    _SPEC_DIR = Path(SPECPATH)  # noqa: F821 — глобал PyInstaller
except NameError:
    try:
        _SPEC_DIR = Path(__file__).resolve().parent
    except NameError:
        _SPEC_DIR = Path.cwd()

# VersionInfo для exe — версия видна в свойствах файла (п. 6.1 плана).
# Аудит 2.0.1: версия больше НЕ дублируется в spec — она читается из
# APP_VERSION в main.py (единый источник), и бинарные filevers/prodvers
# вычисляются из неё же (раньше строка была 2.0.1, а filevers=(1,5,0,0)).
_match = re.search(
    r'APP_VERSION\s*=\s*"(\d+)\.(\d+)\.(\d+)"',
    (_SPEC_DIR / "main.py").read_text(encoding="utf-8"),
)
if not _match:
    raise SystemExit("EWDochub.spec: не удалось прочитать APP_VERSION из main.py")
APP_VERSION = ".".join(_match.groups())
FILEVERS = tuple(int(group) for group in _match.groups()) + (0,)

exe_version = None
if sys.platform == "win32":
    # VersionInfo осмыслен только для Windows-сборки; вне Windows блок не
    # выполняется вовсе (аудит 2.0.1: не «глотать» исключение, а не заходить).
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
                            StringStruct("FileDescription", "EW.DocHub — контрагенты и документы"),
                            StringStruct("FileVersion", APP_VERSION),
                            StringStruct("InternalName", "EW.DocHub"),
                            StringStruct("LegalCopyright", "© EW.DocHub"),
                            StringStruct("OriginalFilename", "EW.DocHub.exe"),
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

customtkinter_datas = collect_data_files("customtkinter")
customtkinter_imports = collect_submodules("customtkinter")

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('dogovor_template.docx', '.'),
        ('dogovor_template_ip.docx', '.'),
        ('additional_agreement.docx', '.'),
        ('additional_agreement_ip.docx', '.'),
        ('spec_template.docx', '.'),
        ('spec_template_ip.docx', '.'),
        ('app_icon.ico', '.'),
    ] + customtkinter_datas,
    hiddenimports=customtkinter_imports + ['darkdetect'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Исключаем тяжёлые пакеты, которые приложению не нужны:
    # это ускоряет и сборку, и импорт при старте.
    # ВАЖНО: pkg_resources тоже исключаем. Он входит в setuptools и
    # зависит от jaraco.text / jaraco.context / jaraco.functools.
    # Если pkg_resources попадает в сборку, PyInstaller добавляет
    # runtime-хук pyi_rth_pkgres, который падает с
    # ModuleNotFoundError: No module named 'jaraco.text'.
    # Приложение pkg_resources не использует, поэтому исключаем и его —
    # тогда хук вообще не включается в exe.
    excludes=[
        'numpy', 'matplotlib', 'pandas', 'scipy', 'sympy',
        'IPython', 'jupyter', 'notebook',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
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
    name='EW.DocHub',
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
    name='EW.DocHub',
)
