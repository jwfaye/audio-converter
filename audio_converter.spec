# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for audio-converter GUI
# Build: pyinstaller audio_converter.spec

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ["audio_converter/gui.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        "numpy",
        "soundfile",
        "scipy.signal",
        "openpyxl",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib",
        "PIL",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "IPython",
        "jupyter",
        "notebook",
        "pytest",
        "sphinx",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="audio-converter-gui",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No console window on Windows
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/app.ico",
)
