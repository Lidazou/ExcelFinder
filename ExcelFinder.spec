# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for ExcelFinder.

Produces a single self-contained ExcelFinder.exe (windowed, no console, no
Python installation required on the target machine).
"""
import os

block_cipher = None

ROOT = os.path.abspath(os.getcwd())
SRC = os.path.join(ROOT, "src")

a = Analysis(
    [os.path.join(SRC, "excelfinder_gui.py")],
    pathex=[SRC],
    binaries=[],
    datas=[],
    hiddenimports=["excelfinder_core"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Keep the binary lean: nothing here is used by ExcelFinder.
        # Deliberately conservative -- only third-party packages and stdlib
        # modules the app provably never touches.
        "numpy", "pandas", "matplotlib", "scipy", "PIL", "lxml",
        "markitdown", "onnxruntime", "magika", "bs4", "requests",
        "pytest", "setuptools", "pip", "sqlite3", "distutils", "test",
        "pydoc", "doctest", "unittest", "lib2to3", "idlelib", "turtledemo",
        "tkinter.test", "curses", "asyncio",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ExcelFinder",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI app: no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(ROOT, "assets", "excelfinder.ico")
    if os.path.exists(os.path.join(ROOT, "assets", "excelfinder.ico"))
    else None,
    version=os.path.join(ROOT, "assets", "version_info.txt")
    if os.path.exists(os.path.join(ROOT, "assets", "version_info.txt"))
    else None,
)
