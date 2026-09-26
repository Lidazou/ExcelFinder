# -*- coding: utf-8 -*-
"""ExcelFinder command-line tool (console build).

Two jobs:
  * ``ExcelFinder-cli.exe <dir> --query 关键词`` -- search without the GUI
  * ``ExcelFinder-cli.exe <dir> --content``       -- full engine self-check

Because this shares the exact engine modules with the GUI build, a passing run
here proves the frozen binary's libraries, threading and Excel parsers all work.
"""
from __future__ import annotations

import sys

from excelfinder_core import run_selftest, setup_console

if __name__ == "__main__":
    setup_console()          # must run before the first print
    sys.exit(run_selftest(sys.argv[1:]))
