# -*- mode: python ; coding: utf-8 -*-
# Сборка ONEDIR приложения EW.DocHub: exe лежит в папке вместе с библиотеками.
# Запуск почти мгновенный — ничего не распаковывается во временную папку.
# Итоговая папка dist/EW.DocHub упаковывается установщиком Inno Setup (installer.iss).
# Сборка: pyinstaller EWDochub.spec --noconfirm

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

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
