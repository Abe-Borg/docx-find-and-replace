# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec - single-file portable build.

One self-contained .exe for people who cannot or would rather not install
anything: copy it anywhere and double-click. It unpacks itself to a temp
directory on each launch, so it starts more slowly than the installed build
and is likelier to be flagged by antivirus heuristics.

Build:  pyinstaller packaging/app-portable.spec --noconfirm
Output: dist/DocxFindReplace-Portable.exe
"""

import os
import sys

from PyInstaller.utils.hooks import collect_data_files

ROOT = os.path.dirname(SPECPATH)
sys.path.insert(0, ROOT)
from version import APP_ID                # noqa: E402

datas = collect_data_files("docx")

a = Analysis(
    [os.path.join(ROOT, "main.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=["lxml._elementpath"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "PIL", "numpy", "matplotlib", "setuptools"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=f"{APP_ID}-Portable",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon=os.path.join(SPECPATH, "icon.ico"),
    version=os.path.join(SPECPATH, "file_version_info.txt"),
)
