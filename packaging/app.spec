# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec - one-folder build.

This is the build the Inno Setup installer wraps. One-folder rather than
one-file because it starts faster (no self-extraction to a temp directory on
every launch) and draws fewer antivirus false positives, both of which matter
on a managed corporate desktop.

Build:  pyinstaller packaging/app.spec --noconfirm
Output: dist/DocxFindReplace/DocxFindReplace.exe
"""

import os
import sys

from PyInstaller.utils.hooks import collect_data_files

ROOT = os.path.dirname(SPECPATH)          # packaging/ -> repository root
sys.path.insert(0, ROOT)
from version import APP_ID                # noqa: E402

# python-docx loads XML templates from its package directory at runtime, and
# PyInstaller cannot see them from the import graph alone.
datas = collect_data_files("docx")

a = Analysis(
    [os.path.join(ROOT, "main.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    # lxml resolves this one dynamically, so the import graph misses it.
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
    [],
    exclude_binaries=True,
    name=APP_ID,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX compression is a common AV trigger
    console=False,             # GUI app: no console window behind the window
    disable_windowed_traceback=False,
    icon=os.path.join(SPECPATH, "icon.ico"),
    version=os.path.join(SPECPATH, "file_version_info.txt"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name=APP_ID,
)
