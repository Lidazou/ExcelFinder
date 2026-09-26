# -*- coding: utf-8 -*-
"""In-process integration test for the ExcelFinder GUI logic.

``check_window.py`` / ``verify_gui_scan.py`` prove the packaged EXE starts and
scans. This test drives the *same* ExcelFinderApp object through the paths a
user exercises by hand -- build index, run a query, enable content matching --
while asserting on the real widgets and on the on-disk cache.

Usage: python tools/gui_integration.py <directory> [--content]
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import threading
import time
import tkinter as tk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "src"))

import excelfinder_gui as G  # noqa: E402

FAILURES: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
    if not ok:
        FAILURES.append(label)


def widget_state(widget, option: str) -> str:
    """Read a ttk option as a plain string.

    ttk returns Tcl_Obj wrappers (``<index object: 'disabled'>``), which never
    compare equal to a Python str -- always go through str().
    """
    return str(widget.cget(option))


def busy_state(app) -> tuple[bool, bool]:
    """(stop_enabled, rebuild_disabled).

    ttk's ``widget['state']`` returns a Tcl index object, so ``instate()`` is the
    only reliable way to query it.
    """
    return app.btn_stop.instate(["!disabled"]), app.btn_scan.instate(["disabled"])


def drive_until(root: tk.Tk, predicate, timeout: float = 180.0) -> bool:
    """Run the tk event loop by hand until `predicate()` or the timeout."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        root.update()
        if predicate():
            return True
        time.sleep(0.02)
    return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("directory")
    ap.add_argument("--content", action="store_true", help="also exercise content matching")
    args = ap.parse_args()
    directory = os.path.abspath(args.directory)

    state_dir = G.CONFIG_DIR          # respects portable mode
    shutil.rmtree(state_dir, ignore_errors=True)
    os.makedirs(state_dir, exist_ok=True)
    index_path = os.path.join(state_dir, "index_v1.json")
    print(f"状态目录：{state_dir}（便携模式={G.PORTABLE}）")

    print("1. 构建界面")
    root = tk.Tk()
    app = G.ExcelFinderApp(root)
    root.update_idletasks()
    check("窗口与控件构建成功", len(root.winfo_children()) == 1)
    check("索引初始为空", len(app.index.records) == 0)

    print("2. 扫描目录（走真实的后台线程 + 消息泵）")
    app.dir_var.set(directory)
    root.update()
    # The app may already be scanning: ExcelFinderApp._restore() loads a cached
    # index and kicks off an incremental scan when a root directory is known.
    if not app.busy:
        app._start_scan()
    stop_on, rebuild_off = busy_state(app)
    check("扫描启动时进入忙碌状态（停止可用/重建禁用）", stop_on and rebuild_off,
          f"stop_enabled={stop_on} rebuild_disabled={rebuild_off}")
    t0 = time.perf_counter()
    ok = drive_until(root, lambda: not app.busy and bool(app.index.records), timeout=180)
    check("索引构建完成", ok)
    if not ok:
        root.destroy()
        return 1
    print(f"     索引 {len(app.index.records):,} 个文件，用时 {time.perf_counter()-t0:.2f}s")
    stop_on, rebuild_off = busy_state(app)
    check("扫描结束后恢复可用", stop_on is False and rebuild_off is False)
    check("结果表格已填充", len(app.tree.get_children()) == len(app.index.records))
    check("状态栏报告了命中数", "命中" in app.status.get(), app.status.get()[:80])
    check("导出按钮已启用", app.btn_export.instate(["!disabled"]))

    print("3. 关键词检索")
    for query, expect_min in (("预算", 1), ("ZZZ_不存在_ZZZ", 0)):
        app.query_var.set(query)
        app._do_search()
        root.update()
        rows = len(app.tree.get_children())
        check(f"查询 {query!r} 返回 {rows} 行", rows >= expect_min)
    matching = sum(1 for r in app.index.records if "预算" in r.name)
    app.query_var.set("预算")
    app._do_search()
    root.update()
    check(
        "命中数与手工统计一致",
        len(app.tree.get_children()) == matching,
        f"GUI={len(app.tree.get_children())} 期望={matching}",
    )

    print("3b. 进度反馈")
    app._do_search()
    root.update()
    check("检索结束后进度条显示完成", "完成" in app.progress_text.get(),
          app.progress_text.get())
    check("进度条为满格", int(float(app.progress["value"])) == 100,
          f"value={app.progress['value']}")
    check("进度条为确定模式（不是无休止条纹）",
          widget_state(app.progress, "mode") == "determinate",
          widget_state(app.progress, "mode"))

    print("4. 排序")
    app._sort_by("name")
    root.update()
    first_asc = app.tree.item(app.tree.get_children()[0])["values"][0]
    app._sort_by("name")
    root.update()
    first_desc = app.tree.item(app.tree.get_children()[0])["values"][0]
    check("名称排序可切换升降序", first_asc != first_desc, f"{first_asc!r} -> {first_desc!r}")
    check("排序后进度条报告已排序", "排序" in app.progress_text.get(), app.progress_text.get())

    print("4b. 打开方式（Office / WPS）选择")
    from excelfinder_core import OPEN_WITH_CHOICES, installed_apps

    detected = installed_apps(refresh=True)
    print(f"     本机检测到的表格程序: {detected or '（无）'}")
    check("能检测本机表格程序", isinstance(detected, dict))
    labels = [label for label, _cid, _apps in OPEN_WITH_CHOICES]
    combo_values = [str(v) for v in app.open_with_combo.cget("values")]
    check("下拉框含全部三个选项", combo_values == labels, str(combo_values))
    check("当前选项合法", app.open_with_var.get() in labels, app.open_with_var.get())
    for label, cid, _apps in OPEN_WITH_CHOICES:
        app.open_with_var.set(label)
        check(f"选项「{label}」→ id 正确", app._open_with_id() == cid, app._open_with_id())
    app.open_with_var.set("只用 WPS 表格打开")
    app._on_open_with()
    root.update()
    check("选择已记入内存配置", app.cfg.get("open_with") == "wps", str(app.cfg.get("open_with")))
    app._save_cfg()
    with open(os.path.join(state_dir, "config.json"), encoding="utf-8") as fh:
        check("config.json 持久化了打开方式", json.load(fh).get("open_with") == "wps")
    app.open_with_var.set(OPEN_WITH_CHOICES[0][0])
    app._on_open_with()
    root.update()

    print("5. 索引缓存落盘")
    ok = drive_until(root, lambda: os.path.exists(index_path), timeout=30)
    check("index_v1.json 已写入", ok)
    if ok:
        with open(index_path, encoding="utf-8") as fh:
            payload = json.load(fh)
        check("缓存记录数与内存一致", len(payload["records"]) == len(app.index.records))
        check("缓存 root 正确", os.path.normcase(payload["root"]) == os.path.normcase(directory))
    check("config.json 已写入", os.path.exists(os.path.join(state_dir, "config.json")))

    print("6. 索引重载")
    reloaded = G.Index()
    check("可从磁盘重新加载索引", reloaded.load(index_path))
    check("重载后记录数一致", len(reloaded.records) == len(app.index.records))

    if args.content:
        print("7. 内容匹配：勾选后搜索应自动建立内容索引")
        # Wait out the restore-triggered auto scan so the app is idle.
        drive_until(root, lambda: not app.busy, timeout=180)
        check("内容索引初始为空", app.index.stats()["with_content"] == 0)

        # Searching cell data means "content only"; leaving both boxes ticked
        # would AND the terms and correctly match nothing.
        app.req_name.set(False)
        app.req_content.set(True)
        check("范围为“仅内容”", app._build_matcher().scope() == "content")
        app.query_var.set("扫地机器人")     # appears in cell data, never in a file name
        app._do_search()                     # must kick off content parsing itself
        check("搜索自动触发了内容索引", app.busy)
        rows_during = len(app.tree.get_children())
        app._do_search()                     # second call while busy must not stack
        check("内容索引进行中不返回不完整结果", len(app.tree.get_children()) == rows_during)
        ok = drive_until(
            root,
            lambda: not app.busy
            and app.index.records
            and all(r.content_loaded for r in app.index.records),
            timeout=600,
        )
        check("全部文件内容已解析", ok)
        loaded = sum(1 for r in app.index.records if r.content_loaded)
        print(f"     已解析 {loaded:,}/{len(app.index.records):,}")
        check("内容索引统计已更新", app.index.stats()["with_content"] == loaded)

        # _on_content_done re-runs the query, so wait for the pump to refill the
        # tree instead of asserting the instant `busy` clears.
        ok = drive_until(root, lambda: len(app.tree.get_children()) > 0, timeout=60)
        check("内容索引完成后自动重新搜索", ok)
        rows = len(app.tree.get_children())
        check("内容查询返回结果", rows > 0, f"{rows} 行")
        if rows:
            values = [app.tree.item(i)["values"] for i in app.tree.get_children()[:5]]
            where = [str(v[4]) for v in values]
            names = [str(v[0]) for v in values]
            # The column now carries the cell address (e.g. 西北汇总!E6 等3处);
            # fall back to the older "内容" label when location is unavailable.
            check("命中位置列标注了内容或单元格地址",
                  all(("内容" in w) or ("!" in w) for w in where), str(where))
            check("命中片段非空", all(str(v[5]).strip() for v in values),
                  str([str(v[5])[:40] for v in values[:2]]))
            check("关键词确实不在文件名中（证明是内容命中）",
                  all("扫地机器人" not in n for n in names), str(names[:3]))

        print("7b. 内容检索阶段的进度与排序复用")
        check("内容检索结束后进度显示完成", "完成" in app.progress_text.get(),
              app.progress_text.get())
        rows_now = [str(app.tree.item(i)["values"][4]) for i in app.tree.get_children()]
        without_addr = [w for w in rows_now if "!" not in w]
        # .csv/.txt matches have no A1-style address, so "内容" is the honest
        # label for them -- assert that, rather than demanding an address everywhere.
        check("无单元格地址的行都是非 OOXML 格式（csv/txt）",
              all(w in ("内容", "文件名", "文件名+内容") for w in without_addr),
              f"{len(without_addr)}/{len(rows_now)} 行，示例={without_addr[:3]}")

        t0 = time.perf_counter()
        app._sort_by("size")
        sort_secs = time.perf_counter() - t0
        rows_after = [str(app.tree.item(i)["values"][4]) for i in app.tree.get_children()]
        # Re-reading every workbook on sort would take seconds for 500 files.
        check("排序不需要重新读取工作簿（<1s）", sort_secs < 1.0, f"{sort_secs:.3f}s")
        check("排序后带地址的行数与排序前一致",
              len([w for w in rows_after if "!" in w]) == len([w for w in rows_now if "!" in w]),
              f"{len([w for w in rows_after if '!' in w])} vs {len([w for w in rows_now if '!' in w])}")
        check("排序后地址没有被降级为“内容”",
              len([w for w in rows_after if w == "内容"]) == len(without_addr),
              f"{len([w for w in rows_after if w == '内容'])} vs {len(without_addr)}")

    root.destroy()
    print()
    if FAILURES:
        print(f"GUI INTEGRATION FAILED ({len(FAILURES)}): {FAILURES}")
        return 1
    print("GUI INTEGRATION PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
