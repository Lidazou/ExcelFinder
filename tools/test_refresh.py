# -*- coding: utf-8 -*-
"""Test the 刷新列表 (refresh) button: does it pick up real file changes?

Adds, renames and deletes files in a scratch copy of the corpus, presses
refresh, and checks the visible list follows the disk.

Usage: python tools/test_refresh.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import time
import tkinter as tk

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, HERE)

import excelfinder_gui as G  # noqa: E402

FAILURES: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
    if not ok:
        FAILURES.append(label)


def drive(root: tk.Tk, predicate, timeout: float = 120.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if predicate():
            return True
        time.sleep(0.02)
    return False


def main() -> int:
    scratch = os.path.join(ROOT, ".refresh-test")
    shutil.rmtree(scratch, ignore_errors=True)
    os.makedirs(scratch, exist_ok=True)
    # a small fixture: two files up front
    sys.path.insert(0, HERE)
    from xlsxwriter_lite import write_xlsx

    write_xlsx(os.path.join(scratch, "甲_预算表.xlsx"),
               [("S1", [["姓名", "金额"], ["张伟", 100]])])
    write_xlsx(os.path.join(scratch, "乙_合同台账.xlsx"),
               [("S1", [["姓名", "金额"], ["李娜", 200]])])

    root = tk.Tk()
    root.withdraw()
    app = G.ExcelFinderApp(root)
    # stop the restore-triggered auto scan and start clean on the scratch dir
    drive(root, lambda: not app.busy, timeout=120)
    app.dir_var.set(scratch)
    app.index = G.Index()
    app._clear_results()
    app.query_var.set("")
    app._start_scan()
    drive(root, lambda: not app.busy and bool(app.index.records), timeout=120)

    print("1. 初次索引")
    check("索引到 2 个文件", len(app.index.records) == 2, f"{len(app.index.records)}")
    check("刷新按钮默认可用", str(app.btn_refresh.cget("state")) == "normal",
          str(app.btn_refresh.cget("state")))

    print("2. 新增一个文件后刷新")
    newfile = os.path.join(scratch, "丙_客户名单.xlsx")
    write_xlsx(newfile, [("S1", [["姓名"], ["王芳"]])])
    app._refresh_list()
    check("刷新期间按钮禁用", str(app.btn_refresh.cget("state")) == "disabled",
          str(app.btn_refresh.cget("state")))
    ok = drive(root, lambda: not app.busy, timeout=120)
    check("刷新完成", ok)
    check("新文件已进入索引", len(app.index.records) == 3, f"{len(app.index.records)}")
    check("状态栏报告了“新增”", "新增" in app.status.get(), app.status.get()[-90:])
    check("状态提到新文件名", "丙_客户名单" in app.status.get(), app.status.get()[-90:])

    print("3. 删除一个文件后刷新")
    os.remove(newfile)
    app._refresh_list()
    drive(root, lambda: not app.busy, timeout=120)
    check("已删除的文件移出索引", len(app.index.records) == 2, f"{len(app.index.records)}")
    check("状态栏报告了“移除”", "移除" in app.status.get(), app.status.get()[-90:])

    print("4. 目录无变化时刷新")
    app._refresh_list()
    drive(root, lambda: not app.busy, timeout=120)
    check("报告目录无变化", "无变化" in app.status.get(), app.status.get()[-90:])

    print("5. 刷新后重新套用关键词")
    app.query_var.set("合同")
    app._do_search()
    root.update()
    hits_before = len(app.tree.get_children())
    check("关键词检索命中 1 个", hits_before == 1, f"{hits_before}")
    write_xlsx(os.path.join(scratch, "丁_合同附件.xlsx"), [("S1", [["备注"], ["x"]])])
    app._refresh_list()
    drive(root, lambda: not app.busy, timeout=120)
    hits_after = len(app.tree.get_children())
    check("刷新后新文件立即出现在结果里", hits_after == 2, f"{hits_before} -> {hits_after}")
    check("关键词仍保留", app.query_var.get() == "合同", app.query_var.get())

    print("6. 未建索引时点刷新")
    app.index = G.Index()
    app._clear_results()
    app._refresh_list()
    drive(root, lambda: not app.busy, timeout=120)
    check("会自动转为首次索引", len(app.index.records) == 3, f"{len(app.index.records)}")

    root.destroy()
    shutil.rmtree(scratch, ignore_errors=True)
    print()
    if FAILURES:
        print(f"REFRESH TEST FAILED ({len(FAILURES)}): {FAILURES}")
        return 1
    print("REFRESH TEST PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
