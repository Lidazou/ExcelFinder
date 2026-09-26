# -*- coding: utf-8 -*-
"""Verify the assembled U盘便携版 package end to end.

Checks that a USB stick would actually work:
  1. no-install runtime  - EXEs launch with no Python/Office present
  2. portable storage    - config + index land in <pkg>\\data, NOT %APPDATA%
  3. drive-letter independence - the package works from another location
  4. GUI really appears  - a visible Tk window is created
  5. CLI works           - console build scans and searches
  6. fallback runtime    - bundled python can start the GUI from source

Usage: python tools/verify_portable.py [package_dir]
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import tkinter as tk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from check_window import descendants, windows_of  # noqa: E402

FAILURES: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
    if not ok:
        FAILURES.append(label)


def hr(t: str) -> None:
    print("\n" + "=" * 70 + f"\n{t}\n" + "=" * 70)


def appdata_state() -> str:
    return os.path.join(os.environ.get("APPDATA", ""), "ExcelFinder")


def run_gui_and_wait(pkg: str, args: list[str], data_dir: str, timeout: float = 120.0):
    """Launch the GUI, wait for its index cache, then kill the whole tree."""
    state_before = os.path.exists(appdata_state())
    env = dict(os.environ)
    env["TMP"] = env["TEMP"] = os.path.join(pkg, "data", ".tmp")
    os.makedirs(env["TMP"], exist_ok=True)
    proc = subprocess.Popen([os.path.join(pkg, "ExcelFinder.exe"), *args], env=env)

    index_path = os.path.join(data_dir, "index_v1.json")
    saw_window = False
    records = None
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(1.0)
        if proc.poll() is not None:
            break
        if not saw_window:
            vis = [w for w in windows_of(descendants(proc.pid)) if w["visible"] and w["size"][0] > 50]
            if vis:
                saw_window = True
                print(f"     窗口出现: {vis[0]['title']!r} {vis[0]['size']}")
        if os.path.exists(index_path):
            try:
                with open(index_path, encoding="utf-8") as fh:
                    records = json.load(fh).get("records", [])
                if records:
                    break
            except (OSError, ValueError):
                pass

    subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True, text=True)
    time.sleep(1.0)
    return {
        "saw_window": saw_window,
        "records": records,
        "config": os.path.exists(os.path.join(data_dir, "config.json")),
        "appdata_created": (not state_before) and os.path.exists(appdata_state()),
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("package", nargs="?", default=os.path.join(HERE, "..", "portable"))
    args = ap.parse_args()
    pkg = os.path.abspath(args.package)
    data_dir = os.path.join(pkg, "data")

    hr(f"待验证的便携包：{pkg}")
    for name in ("ExcelFinder.exe", "ExcelFinder-cli.exe", "portable.marker", "run_gui.py", "使用说明.md"):
        check(f"存在 {name}", os.path.exists(os.path.join(pkg, name)))

    hr("1. 干净状态启动（模拟一台从没见过本程序的电脑）")
    shutil.rmtree(data_dir, ignore_errors=True)
    shutil.rmtree(appdata_state(), ignore_errors=True)
    print("     已清空 data/ 与 %APPDATA%\\ExcelFinder")

    sample = os.path.join(HERE, "..", "testdata")
    if not os.path.isdir(sample):
        print(f"     ! 缺少测试数据 {sample}，跳过部分检查")
        sample = pkg
    sample = os.path.abspath(sample)
    n_expect = sum(
        1 for b, _d, fs in os.walk(sample) for f in fs
        if os.path.splitext(f)[1].lower() in (".xlsx", ".xlsm", ".csv", ".tsv", ".xls", ".xlsb", ".ods")
    )

    print(f"     启动 GUI 并扫描 {sample}")
    res = run_gui_and_wait(pkg, [sample], data_dir)
    check("GUI 窗口成功显示", res["saw_window"])
    check("索引缓存写入了 portable\\data", bool(res["records"]),
          f"{len(res['records']) if res['records'] else 0} 条记录")
    if res["records"]:
        check("索引记录数与磁盘文件数一致", len(res["records"]) == n_expect,
              f"索引={len(res['records'])} 磁盘={n_expect}")
    check("没有在 %APPDATA% 留下任何东西", not res["appdata_created"])

    if os.path.exists(os.path.join(data_dir, "config.json")):
        with open(os.path.join(data_dir, "config.json"), encoding="utf-8") as fh:
            cfg = json.load(fh)
        check("配置里记录了正确的目录",
              os.path.normcase(cfg.get("root", "")) == os.path.normcase(sample))

    hr("2. 命令行版（免安装、脚本可用）")
    # The CLI now emits UTF-8 on purpose; decode it as such or Chinese garbles.
    cli = os.path.join(pkg, "ExcelFinder-cli.exe")
    p = subprocess.run([cli, sample, "--query", "预算", "--limit", "3"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace", timeout=300)
    check("CLI 退出码为 0", p.returncode == 0, f"exit={p.returncode}")
    check("CLI 输出了检索结果", "命中" in (p.stdout or ""))
    if p.stdout:
        for line in p.stdout.strip().splitlines()[-4:]:
            print(f"       {line}")

    p2 = subprocess.run([cli, sample, "--query", "扫地机器人", "--content", "--limit", "2"],
                        capture_output=True, text=True, encoding="utf-8",
                        errors="replace", timeout=600)
    check("CLI 内容检索可用", p2.returncode == 0 and "命中" in (p2.stdout or ""))

    if os.path.exists(os.path.join(pkg, "ExcelFinder.exe")):
        p3 = subprocess.run(
            [cli, sample, "--query", "预算", "--content", "--cells", "1", "--limit", "1"],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=600,
        )
        out3 = p3.stdout or ""
        check("CLI 可定位到单元格（工作表 + 地址）",
              p3.returncode == 0 and "工作表" in out3,
              next((ln.strip() for ln in out3.splitlines() if "工作表" in ln), "")[:80])

    hr("3. 换盘符 / 换位置也能用（模拟插到另一台电脑）")
    moved = os.path.join(os.environ.get("TEMP", "."), "ExcelFinder_moved_test")
    shutil.rmtree(moved, ignore_errors=True)
    try:
        shutil.copytree(pkg, moved, ignore=shutil.ignore_patterns("data", ".tmp"))
        os.makedirs(os.path.join(moved, "data"), exist_ok=True)
        print(f"     整包复制到 {moved}")
        res2 = run_gui_and_wait(moved, [sample], os.path.join(moved, "data"), timeout=120)
        check("复制到新位置后 GUI 仍能启动", res2["saw_window"])
        check("新位置的 data\\ 写入了索引", bool(res2["records"]),
              f"{len(res2['records']) if res2['records'] else 0} 条")
    finally:
        shutil.rmtree(moved, ignore_errors=True)

    hr("4. 内置 Python 备用启动方式")
    py = os.path.join(pkg, "python", "python.exe")
    pyw = os.path.join(pkg, "python", "pythonw.exe")
    if os.path.exists(py):
        probe = subprocess.run(
            [py, "-c", "import tkinter, sys, os; sys.path.insert(0, os.path.join(%r,'src')); "
                       "import excelfinder_core as c; print('OK', c.APP_VERSION, c.is_portable())" % pkg],
            capture_output=True, text=True, errors="replace", timeout=120,
        )
        check("内置解释器可载入引擎（含 tkinter）",
              probe.returncode == 0 and "OK" in (probe.stdout or ""),
              (probe.stdout or probe.stderr or "").strip()[:120])

        env = dict(os.environ)
        env["TMP"] = env["TEMP"] = os.path.join(pkg, "data", ".tmp")
        os.makedirs(env["TMP"], exist_ok=True)
        p3 = subprocess.Popen([pyw, os.path.join(pkg, "run_gui.py"), sample], env=env)
        win = False
        for _ in range(40):
            time.sleep(0.5)
            if p3.poll() is not None:
                break
            vis = [w for w in windows_of(descendants(p3.pid)) if w["visible"] and w["size"][0] > 50]
            if vis:
                win = True
                print(f"     备用启动窗口: {vis[0]['title']!r}")
                break
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(p3.pid)], capture_output=True, text=True)
        check("run_gui.py 备用启动能显示窗口", win)
    else:
        check("内置 Python 存在", False, "缺少 python/python.exe")

    hr("结果")
    if FAILURES:
        print(f"便携包验证失败（{len(FAILURES)} 项）：{FAILURES}")
        return 1
    print("便携包验证全部通过 —— 可直接拷到U盘随插随用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
