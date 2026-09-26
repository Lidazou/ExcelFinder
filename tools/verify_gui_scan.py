# -*- coding: utf-8 -*-
"""Verify the packaged GUI's scan pipeline through its own side effects.

A successful scan run writes two files the user can also see:
  %APPDATA%\\ExcelFinder\\config.json     - remembered root directory
  %APPDATA%\\ExcelFinder\\index_v1.json   - the built index

index_v1.json is only written from inside the "scan finished" handler, so its
presence plus a correct record count proves the GUI queue pump, the scan worker
and the completion path all ran -- which is exactly what a headless harness
cannot observe inside a windowed EXE.

Usage: python tools/verify_gui_scan.py <exe> <directory> [expect_records]
"""
from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from check_window import descendants, windows_of  # noqa: E402

APPDATA_DIR = os.path.join(os.environ.get("APPDATA", ""), "ExcelFinder")


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    exe, directory = sys.argv[1], os.path.abspath(sys.argv[2])
    expect = int(sys.argv[3]) if len(sys.argv) > 3 else None

    shutil.rmtree(APPDATA_DIR, ignore_errors=True)
    print(f"cleared state: {APPDATA_DIR}")
    print(f"launching {os.path.basename(exe)} [{directory}]")

    env = dict(os.environ)
    env["TMP"] = env["TEMP"] = os.path.join(os.getcwd(), ".runtime-tmp")
    os.makedirs(env["TMP"], exist_ok=True)
    errlog = os.path.join(env["TMP"], "gui_stderr.txt")
    errfh = open(errlog, "w", encoding="utf-8", errors="replace")
    proc = subprocess.Popen([exe, directory], env=env, stderr=errfh)

    index_path = os.path.join(APPDATA_DIR, "index_v1.json")
    config_path = os.path.join(APPDATA_DIR, "config.json")
    deadline = time.time() + 120
    records = None
    last_windows: list = []
    while time.time() < deadline:
        time.sleep(1.0)
        if proc.poll() is not None:
            print(f"  !! process exited early with code {proc.returncode}")
            break
        if os.path.exists(index_path):
            try:
                with open(index_path, "r", encoding="utf-8") as fh:
                    payload = json.load(fh)
                records = payload.get("records", [])
                if records:
                    break
            except (OSError, ValueError):
                pass
        last_windows = windows_of(descendants(proc.pid))
        visible = [w for w in last_windows if w["visible"]]
        if len(visible) > 1:                     # a dialog appeared
            print(f"  dialog detected: {[(w['class'], w['title']) for w in visible]}")
            break

    errfh.close()
    try:
        with open(errlog, "r", encoding="utf-8", errors="replace") as fh:
            errtext = fh.read().strip()
    except OSError:
        errtext = ""
    if errtext:
        print("  --- EXE stderr ---")
        for line in errtext.splitlines()[:25]:
            print(f"     {line}")
    print(f"  config.json exists: {os.path.exists(config_path)}")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as fh:
                print(f"  config content   : {json.load(fh)}")
        except (OSError, ValueError):
            pass
    print(f"  APPDATA dir listing: {os.listdir(APPDATA_DIR) if os.path.isdir(APPDATA_DIR) else '(missing)'}")

    ok = False
    if records:
        root_in_index = None
        try:
            with open(index_path, "r", encoding="utf-8") as fh:
                root_in_index = json.load(fh).get("root")
        except (OSError, ValueError):
            pass
        cfg = {}
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as fh:
                    cfg = json.load(fh)
            except (OSError, ValueError):
                pass
        print(f"  index written : {len(records):,} records")
        print(f"  index root    : {root_in_index}")
        print(f"  config root   : {cfg.get('root')}")
        print(f"  config keys   : {sorted(cfg)[:8]}")
        same = os.path.normcase(cfg.get("root", "")) == os.path.normcase(directory)
        print(f"  config root matches launch dir: {same}")
        ok = same and (expect is None or len(records) == expect)
        if expect is not None:
            print(f"  expected {expect:,} records -> {'MATCH' if len(records) == expect else 'MISMATCH'}")
    else:
        print("  !! no index was written -> scan pipeline did not complete")
        wins = windows_of(descendants(proc.pid))
        print(f"  windows still open: {[(w['class'], w['visible']) for w in wins]}")

    subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True, text=True)
    print("  GUI SCAN PIPELINE VERIFIED" if ok else "  GUI SCAN PIPELINE FAILED")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
