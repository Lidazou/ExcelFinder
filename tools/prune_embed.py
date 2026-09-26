# -*- coding: utf-8 -*-
"""Prune the bundled embeddable Python runtime -- verified, not guessed.

python.org's embeddable zip carries a lot a tkinter app never loads (Tcl/Tk
demos, sample images, message catalogs, Tcl's own build leftovers). Those are
most of the file count, and hundreds of tiny files also cost directory space on
a USB stick.

Guessing which are safe is how you ship a broken package: ``tcl/tk8.6/ttk``
looks unused but ``tk.tcl`` sources ``ttk/ttk.tcl`` the moment ``tk.Tk()`` runs.
So every candidate is moved aside and the GUI is *actually started*; only a
candidate whose removal still yields a live Tk window is deleted for real.

Usage:
    python tools/prune_embed.py <python_dir>            # test then delete
    python tools/prune_embed.py <python_dir> --dry-run  # report sizes only
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

#: Candidate paths relative to the runtime root.
CANDIDATES = (
    "tcl/tk8.6/demos",
    "tcl/tk8.6/images",
    "tcl/tk8.6/msgs",
    "tcl/tk8.6/ttk",
    "tcl/tcl8.6/msgs",
    "tcl/tcl8.6/opt0.4",
    "tcl/tcl8.6/http1.0",
    "tcl/tcl8.6/tzdata",
    "tcl/tcl8.6/encoding",
    "lib2to3",
    "test",
    "idlelib",
    "tkinter/test",
    "__pycache__",
    "LICENSE.txt",
    "python.cat",
)

#: A real window with the widgets this app actually uses. A smoke test that only
#: imported tkinter would have passed while the GUI stayed broken.
PROBE = r"""
import sys
import tkinter as tk
from tkinter import ttk
root = tk.Tk()
root.withdraw()
ttk.Style().theme_use('vista')
nb = ttk.Notebook(root); nb.pack()
tab = ttk.Frame(nb); nb.add(tab, text='标签')
ttk.Treeview(tab, columns=('a',), show='headings').pack()
ttk.Combobox(tab, values=['x', 'y'], state='readonly').pack()
ttk.Progressbar(tab, mode='determinate').pack()
ttk.Checkbutton(tab, text='中文选项').pack()
ttk.Label(tab, text='中文标签').pack()
ttk.Entry(tab).pack()
ttk.Scrollbar(tab).pack()
root.update_idletasks()
root.update()
root.destroy()
print('PROBE_OK')
"""


def tree_size(path: str) -> tuple[int, int]:
    total = count = 0
    for base, _dirs, files in os.walk(path):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(base, f))
                count += 1
            except OSError:
                pass
    return count, total


def probe_ok(python_dir: str) -> tuple[bool, str]:
    """Start a real Tk window with the widgets the app uses."""
    exe = os.path.join(python_dir, "python.exe")
    with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False, encoding="utf-8") as fh:
        fh.write(PROBE)
        script = fh.name
    try:
        proc = subprocess.run([exe, script], capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=90)
        out = (proc.stdout or "") + (proc.stderr or "")
        return ("PROBE_OK" in out), out.strip().splitlines()[-1][:140] if out.strip() else ""
    except (OSError, subprocess.TimeoutExpired) as exc:
        return False, f"{type(exc).__name__}"
    finally:
        try:
            os.remove(script)
        except OSError:
            pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("target")
    ap.add_argument("--dry-run", action="store_true", help="report sizes, change nothing")
    ap.add_argument("--keep", action="append", default=[],
                    help="candidate to leave alone (repeatable)")
    args = ap.parse_args()
    target = os.path.abspath(args.target)
    if not os.path.isdir(target):
        print(f"目录不存在：{target}")
        return 1

    before_count, before_bytes = tree_size(target)
    print(f"处理前: {before_count:,} 个文件 / {before_bytes/1024/1024:.1f} MB")

    baseline_ok, baseline_msg = probe_ok(target)
    print(f"基线自检（未改动）: {'OK' if baseline_ok else 'BAD — ' + baseline_msg}")
    if not baseline_ok:
        print("运行时本身就不能启动 Tk，先修好再谈裁剪。")
        return 1

    if args.dry_run:
        print("\n候选（仅统计，不删除）:")
        for rel in CANDIDATES:
            path = os.path.join(target, rel)
            if not os.path.exists(path):
                continue
            if os.path.isdir(path):
                count, size = tree_size(path)
                print(f"  {rel:<28} {count:>5} 个文件  {size/1024:>8.0f} KB")
            else:
                print(f"  {rel:<28} {'':>5}          {os.path.getsize(path)/1024:>8.0f} KB")
        return 0

    print("\n逐个隔离测试：")
    removed_files = removed_bytes = 0
    for rel in CANDIDATES:
        if rel in args.keep:
            print(f"  [跳过] {rel}（--keep）")
            continue
        path = os.path.normpath(os.path.join(target, rel))
        if not path.startswith(target) or not os.path.exists(path):
            continue

        stash = path + ".prune-test"
        try:
            os.rename(path, stash)
        except OSError as exc:
            print(f"  [跳过] {rel}（无法改名：{exc}）")
            continue

        ok, msg = probe_ok(target)
        if not ok:
            os.rename(stash, path)          # restore: this one is needed
            print(f"  [保留] {rel:<28} 删除后会坏（{msg}）")
            continue

        count, size = tree_size(stash) if os.path.isdir(stash) else (1, os.path.getsize(stash))
        if os.path.isdir(stash):
            shutil.rmtree(stash, ignore_errors=True)
        else:
            os.remove(stash)
        removed_files += count
        removed_bytes += size
        print(f"  [删除] {rel:<28} {count:>5} 个文件  {size/1024:>8.0f} KB")

    after_count, after_bytes = tree_size(target)
    print(f"\n处理后: {after_count:,} 个文件 / {after_bytes/1024/1024:.1f} MB")
    print(f"减少  : {removed_files:,} 个文件 / {removed_bytes/1024/1024:.1f} MB")

    final_ok, final_msg = probe_ok(target)
    print(f"最终自检: {'OK' if final_ok else 'BAD — ' + final_msg}")

    # Also confirm the *application* itself still opens a window from the
    # pruned runtime. A plain `--help` run would hang (this is a GUI app), so
    # reuse the Win32 window checker the other tests use.
    if final_ok:
        from check_window import descendants, windows_of

        env = dict(os.environ)
        env["TMP"] = env["TEMP"] = os.path.join(ROOT, ".runtime-tmp")
        os.makedirs(env["TMP"], exist_ok=True)
        proc = subprocess.Popen(
            [os.path.join(target, "pythonw.exe"), os.path.join(ROOT, "run_gui.py")],
            env=env,
        )
        window = False
        for _ in range(40):
            time.sleep(0.5)
            if proc.poll() is not None:
                break
            visible = [w for w in windows_of(descendants(proc.pid))
                       if w["visible"] and w["size"][0] > 50]
            if visible:
                window = True
                print(f"应用启动自检: OK — 窗口 {visible[0]['title']!r}")
                break
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                       capture_output=True, text=True)
        if not window:
            print("应用启动自检: BAD — 裁剪后无法打开窗口")
            return 1
        time.sleep(1)

    return 0 if final_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
