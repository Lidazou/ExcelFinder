# -*- coding: utf-8 -*-
"""Launcher for the bundled Python fallback (U盘便携版).

Run this when ExcelFinder.exe is blocked by SmartScreen / AppLocker:

    python\\pythonw.exe run_gui.py

It only puts the sibling ``src`` folder on the path and hands over to the GUI;
the engine code is identical to the one compiled into the EXE.
"""
from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "src")
if os.path.isdir(SRC) and SRC not in sys.path:
    sys.path.insert(0, SRC)

try:
    from excelfinder_gui import main
except ImportError as exc:                      # pragma: no cover - packaging guard
    sys.stderr.write(
        f"无法载入程序：{exc}\n"
        f"请确认 src 目录与 run_gui.py 在同一文件夹内：{SRC}\n"
    )
    raise SystemExit(1)

if __name__ == "__main__":
    raise SystemExit(main())
