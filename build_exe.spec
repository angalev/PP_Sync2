# -*- mode: python ; coding: utf-8 -*-
"""Сборка одного exe: build.cmd (или python -m PyInstaller build_exe.spec).

Имя файла spec — латиницей намеренно: кириллица в имени exe задаётся здесь,
внутри UTF-8 файла, и не зависит от кодировки консоли, из которой идёт сборка.

Готовый файл: dist\\ПП_синхронизация.exe. Ему нужен только установленный Excel;
config.json, backup и журнал программа по-прежнему держит в C:\\ПП_синхронизация.
"""

from PyInstaller.utils.hooks import collect_data_files

# Готового хука для xlwings нет: он ищет рядом с собой xlam-надстройку и служебные
# файлы и без них падает уже при импорте.
datas = collect_data_files("xlwings")

hiddenimports = [
    "win32timezone",  # pywin32 тянет его лениво, анализатор импорт не видит
]

# Ничего из этого программа не использует, но PyInstaller кладёт по привычке.
excludes = [
    "tkinter",
    "unittest",
    "pydoc_data",
    "numpy",
    "pandas",
    "matplotlib",
    "PIL",
    "pytest",
    "IPython",
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ПП_синхронизация",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX выключен намеренно: сжатый exe корпоративные антивирусы любят
    # принимать за упакованный вредонос, а выигрыш — десяток мегабайт.
    upx=False,
    runtime_tmpdir=None,
    console=False,  # GUI: консольное окно за спиной программе не нужно
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version="version_info.txt",
)
