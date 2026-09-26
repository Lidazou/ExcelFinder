# -*- coding: utf-8 -*-
"""
ExcelFinder GUI -- a tkinter front-end for the ExcelFinder core engine.

Features
  * Pick a folder, build an in-memory index (multi-threaded), then search
    by file name and/or workbook content.
  * Query language: space-separated tokens are AND-ed, e.g. "2023 销售".
  * Modes: contains / starts-with / exact name / wildcard / regex.
  * Scope checkboxes implement "name must match", "content must match", or both.
  * Click-to-sort results, open / reveal / copy, CSV export.
  * Content index is cached to disk, so repeat launches are near-instant.
"""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from excelfinder_core import (
    APP_NAME,
    APP_VERSION,
    CONTENT_EXTS,
    DEFAULT_THREADS,
    EXCEL_EXTS,
    MAX_CONTENT_FILESIZE,
    OPEN_WITH_CHOICES,
    Index,
    MatchOpts,
    Matcher,
    Record,
    ScanEngine,
    app_dir,
    export_csv,
    extract_text,
    find_cells,
    fmt_size,
    fmt_time,
    human_int,
    index_content,
    installed_apps,
    is_portable,
    open_excel_at_cell,
    open_in_explorer,
    setup_console,
    state_dir,
)

# --------------------------------------------------------------------------

#: How many cell hits to locate per file (the second pass reads the workbook).
MAX_CELL_HITS = 50

#: Portable builds keep config + index beside the .exe so a USB stick can move
#: between PCs (and drive letters) without leaving anything behind.
CONFIG_DIR = state_dir()
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")
INDEX_PATH = os.path.join(CONFIG_DIR, "index_v1.json")
PORTABLE = is_portable()

EXT_PRESETS = {
    "全部（表格+文本）": "",
    "Excel 全部格式": ",".join(sorted(e for e in EXCEL_EXTS)),
    "仅 .xlsx/.xlsm": ".xlsx,.xlsm",
    "仅 .xls（旧版）": ".xls",
    "Excel + CSV": ",".join(sorted(e for e in EXCEL_EXTS | {".csv", ".tsv"})),
    "自定义…": "custom",
}

MODE_LABELS = {
    "contains": "包含（默认）",
    "prefix": "文件名前缀",
    "exact": "完全相等",
    "wildcard": "通配符 * ?",
    "regex": "正则表达式",
}
MODE_KEYS = {v: k for k, v in MODE_LABELS.items()}

RESULT_LIMIT = 20000

#: Set EXCELFINDER_TRACE=<path> to append a line-by-line trace of the scan
#: pipeline. Used to diagnose the frozen GUI build, where a windowed EXE has no
#: stdout or stderr to inspect. Inert unless the variable is set.
TRACE_PATH = os.environ.get("EXCELFINDER_TRACE", "")


def trace(msg: str) -> None:
    if not TRACE_PATH:
        return
    try:
        with open(TRACE_PATH, "a", encoding="utf-8") as fh:
            fh.write(f"{time.strftime('%H:%M:%S')} {msg}\n")
    except OSError:
        pass


def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
            return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_config(cfg: dict) -> None:
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False, indent=2)
    except OSError:
        pass


# --------------------------------------------------------------------------

class ExcelFinderApp(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=8)
        self.master: tk.Tk = master
        self.grid(row=0, column=0, sticky="nsew")
        master.rowconfigure(0, weight=1)
        master.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        self.columnconfigure(0, weight=1)

        self.cfg = load_config()
        self.index = Index()
        self.msg_q: queue.Queue = queue.Queue()
        self.cancel_scan = threading.Event()
        self.cancel_content = threading.Event()
        self.scan_thread: threading.Thread | None = None
        self.content_thread: threading.Thread | None = None
        self.busy = False
        self.last_hits: list[tuple[Record, int, str]] = []
        self.last_total = 0
        self.sort_col = ""
        self.sort_desc = True
        self._scan_files = 0
        self._t0 = 0.0
        self._locate_cache: list[dict] | None = None
        self._locate_path: str = ""
        self._progress_label = "处理中"
        self._cancel_search = threading.Event()
        self._located_where: dict[str, str] = {}
        self._scan_reason = ""
        self._paths_before: set[str] = set()
        self._refresh_summary = ""

        self._build_ui()
        self._restore()
        self.after(80, self._pump)
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        self._build_dir_bar()
        self._build_query_bar()
        self._build_results()

        self.status = tk.StringVar(value="就绪。选择目录后点击“重建索引”。")
        bar = ttk.Frame(self)
        bar.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        bar.columnconfigure(0, weight=1)
        ttk.Label(bar, textvariable=self.status, anchor="w").grid(row=0, column=0, sticky="ew")
        mode = "U盘便携模式（配置存于程序目录）" if PORTABLE else f"配置目录：{CONFIG_DIR}"
        ttk.Label(bar, text=mode, anchor="e", foreground="#888").grid(row=0, column=1, sticky="e", padx=(8, 8))

        # Progress feedback: label + bar + percentage, always visible so every
        # long-running step is observable at a glance.
        self.progress = ttk.Progressbar(bar, mode="determinate", length=190, maximum=100)
        self.progress.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.progress_text = tk.StringVar(value="待机")
        ttk.Label(bar, textvariable=self.progress_text, anchor="e", width=22,
                  foreground="#0a5").grid(row=0, column=3, sticky="e", padx=(6, 0))

    # -- directory / index ------------------------------------------------
    def _build_dir_bar(self) -> None:
        box = ttk.LabelFrame(self, text="1  目标目录与索引", padding=8)
        box.grid(row=0, column=0, sticky="ew")
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="目录：").grid(row=0, column=0, sticky="w")
        self.dir_var = tk.StringVar(value=self.cfg.get("root", ""))
        self.dir_entry = ttk.Entry(box, textvariable=self.dir_var)
        self.dir_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(box, text="浏览…", command=self._browse).grid(row=0, column=2, padx=2)
        self.btn_scan = ttk.Button(box, text="重建索引", command=self._start_scan)
        self.btn_scan.grid(row=0, column=3, padx=2)
        self.btn_incr = ttk.Button(box, text="增量更新", command=lambda: self._start_scan(incremental=True))
        self.btn_incr.grid(row=0, column=4, padx=2)
        self.btn_stop = ttk.Button(box, text="停止", command=self._stop_all, state="disabled")
        self.btn_stop.grid(row=0, column=5, padx=2)

        self.stats_var = tk.StringVar(value="尚未建立索引。")
        ttk.Label(box, textvariable=self.stats_var, anchor="w", foreground="#0a5").grid(
            row=1, column=0, columnspan=6, sticky="ew", pady=(6, 0)
        )

        opt = ttk.Frame(box)
        opt.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(6, 0))
        self.skip_hidden = tk.BooleanVar(value=bool(self.cfg.get("skip_hidden", True)))
        self.cache_index = tk.BooleanVar(value=bool(self.cfg.get("cache_index", True)))
        self.threads = tk.IntVar(value=int(self.cfg.get("threads", DEFAULT_THREADS)))
        self.follow_links = tk.BooleanVar(value=bool(self.cfg.get("follow_links", False)))
        ttk.Checkbutton(opt, text="跳过隐藏/临时文件", variable=self.skip_hidden).pack(side="left")
        ttk.Checkbutton(opt, text="索引进度缓存到磁盘", variable=self.cache_index).pack(side="left", padx=8)
        ttk.Checkbutton(opt, text="跟随符号链接", variable=self.follow_links).pack(side="left")
        ttk.Label(opt, text="  扫描线程：").pack(side="left")
        ttk.Spinbox(opt, from_=1, to=128, width=5, textvariable=self.threads).pack(side="left")
        ttk.Button(opt, text="加载上次索引", command=self._load_cached).pack(side="right")

    # -- query ------------------------------------------------------------
    def _build_query_bar(self) -> None:
        box = ttk.LabelFrame(self, text="2  关键词检索", padding=8)
        box.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="关键词：").grid(row=0, column=0, sticky="w")
        self.query_var = tk.StringVar()
        self.query_entry = ttk.Entry(box, textvariable=self.query_var, font=("Microsoft YaHei UI", 11))
        self.query_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.query_entry.bind("<Return>", lambda _e: self._do_search())
        ttk.Button(box, text="搜索", command=self._do_search).grid(row=0, column=2, padx=2)
        ttk.Button(box, text="清空", command=self._clear_query).grid(row=0, column=3, padx=2)
        self.btn_refresh = ttk.Button(box, text="刷新列表", command=self._refresh_list)
        self.btn_refresh.grid(row=0, column=4, padx=2)
        self.btn_content = ttk.Button(box, text="建立内容索引", command=self._start_content)
        self.btn_content.grid(row=0, column=5, padx=2)
        self.btn_export = ttk.Button(box, text="导出 CSV", command=self._export, state="disabled")
        self.btn_export.grid(row=0, column=6, padx=2)
        self.bind_all("<F5>", lambda _e: self._refresh_list())

        ttk.Label(
            box,
            text="「刷新列表」(F5) 会重新扫描当前目录，把新增/删除/改名的文件同步进来，再按当前关键词重搜一遍。",
            foreground="#666",
        ).grid(row=1, column=0, columnspan=7, sticky="w", pady=(4, 0))

        ttk.Label(
            box,
            text="提示：空格分隔多个关键词 = 同时包含（例如： 2023 销售）。  文件名可用通配符 *  ?  或正则。",
            foreground="#666",
        ).grid(row=2, column=0, columnspan=7, sticky="w", pady=(4, 0))

        row = ttk.Frame(box)
        row.grid(row=3, column=0, columnspan=7, sticky="ew", pady=(6, 0))

        self.req_name = tk.BooleanVar(value=True)
        self.req_content = tk.BooleanVar(value=bool(self.cfg.get("req_content", False)))
        ttk.Checkbutton(row, text="文件名需命中", variable=self.req_name).pack(side="left")
        ttk.Checkbutton(row, text="内容需命中", variable=self.req_content).pack(side="left", padx=(4, 4))
        ttk.Label(row, text="（两个都勾 = 必须同时满足；只勾一个 = 只在文件名或只在内容里找）",
                  foreground="#666").pack(side="left", padx=(0, 12))

        ttk.Label(row, text="匹配方式：").pack(side="left")
        self.mode_var = tk.StringVar(value=MODE_LABELS.get(self.cfg.get("mode", "contains"), "包含（默认）"))
        combo = ttk.Combobox(row, textvariable=self.mode_var, values=list(MODE_LABELS), state="readonly", width=13)
        combo.pack(side="left")

        self.case_var = tk.BooleanVar(value=bool(self.cfg.get("case", False)))
        ttk.Checkbutton(row, text="区分大小写", variable=self.case_var).pack(side="left", padx=(8, 16))

        ttk.Label(row, text="格式：").pack(side="left")
        self.ext_var = tk.StringVar(value=self.cfg.get("ext_preset", "全部（表格+文本）"))
        ext_combo = ttk.Combobox(row, textvariable=self.ext_var, values=list(EXT_PRESETS), state="readonly", width=17)
        ext_combo.pack(side="left")
        ext_combo.bind("<<ComboboxSelected>>", self._on_ext_preset)

        row2 = ttk.Frame(box)
        row2.grid(row=4, column=0, columnspan=7, sticky="ew", pady=(6, 0))
        ttk.Label(row2, text="排除路径含：").pack(side="left")
        self.exclude_var = tk.StringVar(value=self.cfg.get("exclude", ""))
        ttk.Entry(row2, textvariable=self.exclude_var, width=28).pack(side="left", padx=(0, 12))
        ttk.Label(row2, text="自定义扩展名：").pack(side="left")
        self.ext_custom = tk.StringVar(value=self.cfg.get("ext_custom", ""))
        ttk.Entry(row2, textvariable=self.ext_custom, width=22).pack(side="left", padx=(0, 12))
        ttk.Label(row2, text="最小(KB)：").pack(side="left")
        self.min_kb = tk.StringVar(value=str(self.cfg.get("min_kb", "")))
        ttk.Entry(row2, textvariable=self.min_kb, width=8).pack(side="left")
        ttk.Label(row2, text="  修改时间(YYYY-MM-DD~)：").pack(side="left")
        self.date_from = tk.StringVar(value=self.cfg.get("date_from", ""))
        self.date_to = tk.StringVar(value=self.cfg.get("date_to", ""))
        ttk.Entry(row2, textvariable=self.date_from, width=11).pack(side="left")
        ttk.Label(row2, text="~").pack(side="left")
        ttk.Entry(row2, textvariable=self.date_to, width=11).pack(side="left")

    # -- results ----------------------------------------------------------
    def _build_results(self) -> None:
        outer = ttk.Frame(self)
        outer.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        # Results on top, preview drawer below (collapsed by default so the
        # list keeps the full window height).
        self.pane = ttk.PanedWindow(outer, orient="vertical")
        self.pane.grid(row=0, column=0, sticky="nsew")

        box = ttk.Frame(self.pane)
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)

        cols = ("name", "folder", "size", "mtime", "where", "preview")
        self.tree = ttk.Treeview(box, columns=cols, show="headings", selectmode="extended")
        headings = {
            "name": ("文件名", 300, "w"),
            "folder": ("所在目录", 340, "w"),
            "size": ("大小", 80, "e"),
            "mtime": ("修改时间", 130, "center"),
            "where": ("命中位置", 90, "center"),
            "preview": ("命中内容片段", 360, "w"),
        }
        for key, (text, width, anchor) in headings.items():
            self.tree.heading(key, text=text, command=lambda c=key: self._sort_by(c))
            self.tree.column(key, width=width, anchor=anchor, stretch=(key in ("folder", "preview")))

        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree.tag_configure("odd", background="#f7f9fc")
        self.tree.bind("<Double-1>", lambda _e: self._open_file())
        self.tree.bind("<Button-3>", self._popup)
        self.tree.bind("<Control-c>", lambda _e: self._copy_paths())
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.pane.add(box, weight=4)

        pbox = ttk.LabelFrame(self.pane, text="预览（选中一行即可查看文件内容）", padding=4)
        pbox.rowconfigure(1, weight=1)
        pbox.columnconfigure(0, weight=1)
        self.preview_head = tk.StringVar(value="")
        ttk.Label(pbox, textvariable=self.preview_head, anchor="w",
                  foreground="#0a5").grid(row=0, column=0, columnspan=2, sticky="ew")
        self.preview_list = tk.Listbox(pbox, height=7, font=("Consolas", 9))
        pvsb = ttk.Scrollbar(pbox, orient="vertical", command=self.preview_list.yview)
        phsb = ttk.Scrollbar(pbox, orient="horizontal", command=self.preview_list.xview)
        self.preview_list.configure(yscrollcommand=pvsb.set, xscrollcommand=phsb.set)
        self.preview_list.grid(row=1, column=0, sticky="nsew")
        pvsb.grid(row=1, column=1, sticky="ns")
        phsb.grid(row=2, column=0, sticky="ew")
        self.preview_box = pbox

        bar = ttk.Frame(outer)
        bar.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.show_preview = tk.BooleanVar(value=bool(self.cfg.get("show_preview", False)))
        ttk.Checkbutton(bar, text="显示预览面板", variable=self.show_preview,
                        command=self._toggle_preview).pack(side="left")

        ttk.Label(bar, text="   打开方式：").pack(side="left")
        self.open_with_var = tk.StringVar(value=self._initial_open_with())
        self.open_with_combo = ttk.Combobox(
            bar, textvariable=self.open_with_var, state="readonly", width=26,
            values=[label for label, _cid, _apps in OPEN_WITH_CHOICES],
        )
        self.open_with_combo.pack(side="left")
        self.open_with_combo.bind("<<ComboboxSelected>>", self._on_open_with)

        ttk.Label(bar, text="提示：双击＝打开并跳到命中单元格；右键＝更多操作；Ctrl+C＝复制路径。",
                  foreground="#888").pack(side="left", padx=(12, 0))
        if self.show_preview.get():
            self.pane.add(pbox, weight=1)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="打开文件", command=self._open_file)
        self.menu.add_command(label="定位到命中单元格（按下方“打开方式”）", command=self._open_at_cell)
        self.menu.add_command(label="在资源管理器中定位", command=lambda: self._reveal())
        self.menu.add_command(label="复制完整路径", command=self._copy_paths)
        self.menu.add_separator()
        self.menu.add_command(label="按此文件名再搜索", command=self._search_from_selection)
        self.menu.add_separator()
        self.menu.add_command(label="导出当前结果到 CSV", command=self._export)

    def _toggle_preview(self) -> None:
        if self.show_preview.get():
            self.pane.add(self.preview_box, weight=1)
        else:
            self.pane.forget(self.preview_box)

    # ------------------------------------------------- spreadsheet program
    def _initial_open_with(self) -> str:
        """Label for the saved choice, defaulting to auto.

        When only one program is installed, preselect that one so the common
        single-app case needs no thought.
        """
        saved = self.cfg.get("open_with", "auto")
        available = installed_apps()
        if saved == "auto" and len(available) == 1:
            saved = next(iter(available))
        for label, choice_id, _apps in OPEN_WITH_CHOICES:
            if choice_id == saved:
                return label
        return OPEN_WITH_CHOICES[0][0]

    def _open_with_id(self) -> str:
        label = self.open_with_var.get()
        for text, choice_id, _apps in OPEN_WITH_CHOICES:
            if text == label:
                return choice_id
        return "auto"

    def _on_open_with(self, _event=None) -> None:
        choice = self._open_with_id()
        available = installed_apps(refresh=True)
        if choice != "auto" and choice not in available:
            names = "、".join(available.values()) or "无"
            messagebox.showwarning(
                APP_NAME,
                f"这台电脑上没有检测到该表格程序，将回退到可用的：{names}",
            )
        self.cfg["open_with"] = choice
        self._save_cfg()
        self.status.set(f"打开方式已设为：{self.open_with_var.get()}")

    def _keyword(self) -> str:
        tokens = self.query_var.get().split()
        return tokens[0] if tokens else ""

    def _locate(self, rec: Record) -> list[dict]:
        """Cell-level hits inside one file, cached per selection."""
        if self._locate_cache is not None and self._locate_path == rec.path:
            return self._locate_cache
        hits: list[dict] = []
        keyword = self._keyword()
        if keyword and rec.ext in CONTENT_EXTS:
            try:
                hits = find_cells(
                    rec.path, keyword, case=self.case_var.get(), limit=MAX_CELL_HITS
                )
            except Exception:
                hits = []
        self._locate_cache = hits
        self._locate_path = rec.path
        return hits

    def _on_select(self, _event=None) -> None:
        recs = self._selected()
        self._locate_cache = None
        self._locate_path = ""
        if not self.show_preview.get() or not recs:
            return

        rec = recs[0]
        self.preview_list.delete(0, "end")
        self.preview_head.set(
            f"{rec.name}   ·   {fmt_size(rec.size)}   ·   {fmt_time(rec.mtime)}"
        )

        located = self._locate(rec)
        if located:
            self.preview_list.insert("end", f"在内容中找到 {len(located)} 处命中：")
            for i, c in enumerate(located, 1):
                self.preview_list.insert(
                    "end", f"  {i}. 工作表「{c['sheet']}」 {c['ref']} = {c['value'][:60]}"
                )
                if c.get("row_text"):
                    self.preview_list.insert("end", f"       该行: {c['row_text']}")
            self.preview_list.insert("end", "─" * 60)
            self.preview_list.insert("end", "提示：双击可直接打开并跳到第一处命中；右键可定位文件。")

        text = rec.content
        if text is None:
            # Content index not built yet: parse this one file on demand so the
            # preview works even in name-only mode.
            self.preview_list.insert("end", "（尚未建立内容索引，正在读取该文件…）")
            self.preview_list.update_idletasks()
            try:
                if rec.size > MAX_CONTENT_FILESIZE:
                    self.preview_list.delete(0, "end")
                    self.preview_head.set(self.preview_head.get() + "   ·   文件过大，未预览")
                    return
                text = extract_text(rec.path, rec.size)
            except Exception as exc:
                self.preview_list.delete(0, "end")
                self.preview_list.insert("end", f"无法读取内容：{exc}")
                return
            self.preview_list.delete(0, "end")
            if located:
                self.preview_list.insert("end", f"在内容中找到 {len(located)} 处命中：")
                for i, c in enumerate(located, 1):
                    self.preview_list.insert(
                        "end", f"  {i}. 工作表「{c['sheet']}」 {c['ref']} = {c['value'][:60]}"
                    )
                    if c.get("row_text"):
                        self.preview_list.insert("end", f"       该行: {c['row_text']}")
                self.preview_list.insert("end", "─" * 60)

        needle = self._keyword()
        self.preview_list.insert("end", rec.path)
        self.preview_list.insert("end", "─" * 60)
        body = text or "（该文件没有可提取的文本）"
        if needle:
            pos = body.casefold().find(needle.casefold())
            if pos > 0:
                start = max(0, pos - 120)
                self.preview_list.insert("end", f"…{body[start:pos + 200]}…")
                self.preview_list.insert("end", "─" * 60)
        for line in body.split("\n"):
            for chunk_start in range(0, max(len(line), 1), 400):
                self.preview_list.insert("end", line[chunk_start:chunk_start + 400])

    # ------------------------------------------------------------- helpers
    def _restore(self) -> None:
        root = self.dir_var.get()
        trace(f"restore: root={root!r} index_exists={os.path.exists(INDEX_PATH)}")
        if root and os.path.isdir(root) and os.path.exists(INDEX_PATH):
            self._load_cached(silent=True)
        if root and os.path.isdir(root):
            self._start_scan(incremental=True)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self.btn_scan, self.btn_incr, self.btn_content, self.btn_refresh):
            btn.configure(state=state)
        self.btn_stop.configure(state="normal" if busy else "disabled")

    # ------------------------------------------------------------ progress
    def _progress_start(self, label: str, total: int | None = None) -> None:
        """Begin a progress phase.

        ``total`` given -> determinate bar with a live percentage.
        ``total`` None  -> indeterminate stripe for work whose size is unknown
        (a directory walk cannot know how many files it will find up front).
        """
        self._progress_label = label
        if total and total > 0:
            self.progress.configure(mode="determinate", maximum=total, value=0)
            self.progress_text.set(f"{label} 0%")
        else:
            self.progress.configure(mode="indeterminate")
            self.progress.start(60)
            self.progress_text.set(f"{label} …")

    def _progress_update(self, done: int, total: int, label: str | None = None) -> None:
        if not total or total <= 0:
            return
        pct = min(100, int(done * 100 / total))
        self.progress.configure(mode="determinate", maximum=total, value=done)
        self.progress_text.set(f"{label or self._progress_label} {pct}%")

    def _progress_done(self, text: str = "完成 100%") -> None:
        try:
            self.progress.stop()
        except tk.TclError:
            pass
        self.progress.configure(mode="determinate", maximum=100, value=100)
        self.progress_text.set(text)

    def _progress_idle(self, text: str = "待机") -> None:
        try:
            self.progress.stop()
        except tk.TclError:
            pass
        self.progress.configure(mode="determinate", maximum=100, value=0)
        self.progress_text.set(text)

    def _browse(self) -> None:
        initial = self.dir_var.get() or os.path.expanduser("~")
        chosen = filedialog.askdirectory(title="选择要搜索的目录", initialdir=initial)
        if chosen:
            self.dir_var.set(os.path.normpath(chosen))
            self._start_scan()

    def _clear_query(self) -> None:
        self.query_var.set("")
        self.query_entry.focus_set()

    def _clear_results(self) -> None:
        """Drop the previous result set (used when the index is rebuilt)."""
        self.tree.delete(*self.tree.get_children())
        self.last_hits = []
        self.last_total = 0
        self.sort_col = ""
        self.sort_desc = True
        self._located_where.clear()
        if hasattr(self, "btn_export"):
            self.btn_export.configure(state="disabled")

    def _refresh_list(self) -> None:
        """Re-scan the current folder and re-run the query.

        Unlike 重建索引 this keeps the incremental cache, so only new, deleted
        or modified files cost anything; the query is then re-applied so the
        visible list reflects what is on disk right now.
        """
        if self.busy:
            self.status.set("正在忙，请等当前操作结束或点“停止”。")
            return
        root = self.dir_var.get().strip().strip('"')
        if not root or not os.path.isdir(root):
            messagebox.showwarning(APP_NAME, "请先选择一个存在的目录。")
            return

        if not self.index.records:
            # Nothing indexed yet, so a "refresh" is really the first build.
            self.status.set("尚未建立索引，正在首次扫描…")
            self._start_scan()
            return

        # Count changes so the user learns whether the refresh did anything.
        self.status.set("正在刷新列表：重新扫描目录…")
        self._start_scan(incremental=True, reason="refresh")

    def _on_ext_preset(self, _event=None) -> None:
        if self.ext_var.get() == "自定义…":
            self.ext_custom.focus_set()

    def _current_exts(self) -> str:
        preset = self.ext_var.get()
        if preset == "自定义…":
            return self.ext_custom.get().strip()
        return EXT_PRESETS.get(preset, "")

    # ---------------------------------------------------------------- scan
    def _start_scan(self, incremental: bool = False, reason: str = "") -> None:
        """Start a scan. `reason` is carried to the completion handler so a
        refresh can report what actually changed."""
        if self.busy:
            trace("start_scan: ignored, busy")
            return
        root = self.dir_var.get().strip().strip('"')
        if not root or not os.path.isdir(root):
            trace(f"start_scan: bad root {root!r}")
            messagebox.showwarning(APP_NAME, "请先选择一个存在的目录。")
            return
        trace(f"start_scan: root={root!r} incremental={incremental} reuse={len(self.index.records)}")

        # Snapshot tkinter values BEFORE any worker starts. Reading a tkinter
        # variable off the main thread raises "main thread is not in main loop".
        threads = max(1, int(self.threads.get()))
        self._use_cache = bool(self.cache_index.get())
        self._scan_reason = reason
        # Remember the pre-scan file set so a refresh can report the delta.
        self._paths_before = {r.path for r in self.index.records} if incremental else set()

        reuse = self.index.records if (incremental and self.index.records) else None
        if reuse is None:
            self.index = Index()
            self._clear_results()

        self.cancel_scan.clear()
        self._scan_files = 0
        self._t0 = time.perf_counter()
        self.stats_var.set(f"正在扫描 {root} …")
        self._progress_start("扫描目录")          # unknown size -> stripe
        self._set_busy(True)

        def report(rec: Record, n: int) -> None:
            if n % 200 == 0 or n < 40:
                self.msg_q.put(("progress", n))

        engine = ScanEngine(
            root,
            threads=threads,
            on_file=report,
            skip_hidden=bool(self.skip_hidden.get()),
            follow_links=bool(self.follow_links.get()),
            reuse=reuse,
        )
        self._engine = engine

        def runner() -> None:
            try:
                records, seconds, reused = engine.run()
                trace(f"scan runner: {len(records)} records in {seconds:.3f}s, reused={reused}")
                self.msg_q.put(("scan_done", records, seconds, reused, root))
            except Exception as exc:                      # pragma: no cover
                trace(f"scan runner FAILED: {exc!r}")
                self.msg_q.put(("error", f"扫描失败：{exc}"))

        self.scan_thread = threading.Thread(target=runner, name="scan-runner", daemon=True)
        self.scan_thread.start()

    def _on_scan_done(self, records, seconds, reused, root) -> None:
        trace(f"on_scan_done: {len(records)} records, root={root!r}")
        reason = getattr(self, "_scan_reason", "")
        paths_before = getattr(self, "_paths_before", set())
        paths_after = {r.path for r in records}
        added = paths_after - paths_before if paths_before else set()
        removed = paths_before - paths_after if paths_before else set()

        if self.index.records and self.index.root == root:
            # Merge incremental result into the existing index.
            old = {r.path: r for r in self.index.records}
            merged: list[Record] = []
            for rec in records:
                prev = old.get(rec.path)
                merged.append(prev if prev is not None else rec)
            self.index.replace_all(merged, root, seconds)
        else:
            self.index.replace_all(records, root, seconds)

        self._update_stats()
        self._set_busy(False)
        self._progress_done(f"索引完成 {human_int(len(self.index.records))} 个")
        self.cfg["root"] = root
        self._save_cfg()

        if getattr(self, "_use_cache", True) and self._cache_worth_writing():
            def saver() -> None:
                try:
                    self.index.save(INDEX_PATH)
                except OSError:
                    pass
            threading.Thread(target=saver, daemon=True).start()

        if reason == "refresh":
            # Build the summary now, but attach it after _do_search() rewrites
            # the status line, so the user sees the change report and the hits.
            if added or removed:
                parts = []
                if added:
                    names = "、".join(sorted(os.path.basename(p) for p in added)[:3])
                    more = f" 等 {len(added)} 个" if len(added) > 3 else ""
                    parts.append(f"新增 {names}{more}")
                if removed:
                    parts.append(f"移除 {len(removed)} 个")
                summary = "刷新完成：" + "；".join(parts)
            else:
                summary = "刷新完成：目录无变化"
            self._refresh_summary = summary
        else:
            self._refresh_summary = ""

        self._do_search(auto=True)

        if self._refresh_summary:
            self.status.set(self.status.get() + f"  ·  {self._refresh_summary}")
            self._refresh_summary = ""

    def _update_stats(self) -> None:
        st = self.index.stats()
        exts = "  ".join(f"{k}:{v}" for k, v in list(st["by_ext"].items())[:6])
        stamp = fmt_time(st["built_at"]) if st["built_at"] else "-"
        self.stats_var.set(
            f"已索引 {human_int(st['count'])} 个表格文件 · {fmt_size(st['bytes'])} · "
            f"扫描耗时 {st['scan_seconds']:.2f}s · 内容索引 {human_int(st['with_content'])} 个 · 建立于 {stamp}\n{exts}"
        )

    def _load_cached(self, silent: bool = False) -> None:
        if not os.path.exists(INDEX_PATH):
            if not silent:
                messagebox.showinfo(APP_NAME, "没有找到磁盘缓存索引，请先建立索引。")
            return
        if self.index.load(INDEX_PATH):
            self.dir_var.set(self.index.root or self.dir_var.get())
            self._update_stats()
            if not silent:
                self.status.set("已从磁盘缓存加载索引（可直接搜索，无需重新扫描）。")
        elif not silent:
            messagebox.showwarning(APP_NAME, "缓存索引版本不匹配，已忽略。")

    # ------------------------------------------------------------- content
    def _start_content(self) -> None:
        if self.busy:
            return
        if not self.index.records:
            messagebox.showinfo(APP_NAME, "请先建立索引。")
            return
        todo = self.index.paths_needing_content()
        if not todo:
            self.status.set("所有可解析文件的内容索引均已建立。")
            return

        self.cancel_content.clear()
        self._set_busy(True)
        self._progress_start("解析内容", len(todo))
        self.status.set(f"正在解析 {human_int(len(todo))} 个文件的内容…（可继续输入关键词，索引完成后搜索）")

        # Snapshot on the main thread -- see the note in _start_scan.
        threads = max(1, int(self.threads.get()))
        use_cache = bool(self.cache_index.get())
        cancel_event = self.cancel_content
        msg_q = self.msg_q

        def progress(done: int, total: int, rec: Record) -> None:
            # Runs on a worker thread: touching tkinter here would raise
            # "main thread is not in main loop". Queue only; the pump updates UI.
            if done % 5 == 0 or done == total:
                msg_q.put(("content_progress", done, total))

        def runner() -> None:
            try:
                trace(f"content runner: start, todo={len(todo)}")
                done, failed = index_content(
                    todo,
                    threads=threads,
                    on_progress=progress,
                    cancel=cancel_event,
                )
                trace(f"content runner: done={done} failed={failed}")
                msg_q.put(("content_done", done, failed))
            except Exception as exc:                      # pragma: no cover
                trace(f"content runner FAILED: {exc!r}")
                msg_q.put(("error", f"内容索引失败：{exc}"))

        self.content_thread = threading.Thread(target=runner, name="content-runner", daemon=True)
        self.content_thread.start()

    def _on_content_done(self, done: int, failed: int) -> None:
        self._set_busy(False)
        self._progress_done(f"内容索引完成 {human_int(done)} 个")
        self._update_stats()
        note = ""
        if self.cache_index.get():
            # The text cache can be enormous (the extracted text of every file),
            # so only persist it when the write is worth the disk space.
            if self._cache_worth_writing():
                try:
                    self.index.save(INDEX_PATH)
                    note = "　索引已缓存到磁盘。"
                except OSError:
                    pass
            else:
                note = "　内容较大，未写磁盘缓存（重启后需重新解析，但不占空间）。"
        self.status.set(
            f"内容索引完成：成功 {human_int(done)}，失败/跳过 {human_int(failed)}。{note}"
        )
        if self.req_content.get() and self.query_var.get().strip():
            self._do_search()

    #: Above this the on-disk text cache costs more than the rebuild it saves.
    CACHE_SIZE_LIMIT = 256 * 1024 * 1024

    def _cache_worth_writing(self) -> bool:
        """Would persisting the index be reasonable?

        The cached index embeds the extracted text of every file, so a content
        index of a large tree can run to hundreds of megabytes -- far more than
        the program itself. Past the limit the index simply stays in memory.
        """
        chars = sum(len(r.content or "") for r in self.index.records)
        estimated = chars * 1.2 + len(self.index.records) * 200
        return estimated < self.CACHE_SIZE_LIMIT

    # -------------------------------------------------------------- search
    def _build_matcher(self) -> Matcher:
        opts = MatchOpts(
            query=self.query_var.get(),
            mode=MODE_KEYS.get(self.mode_var.get(), "contains"),
            case=self.case_var.get(),
            ext_filter=self._current_exts(),
            req_name=self.req_name.get(),
            req_content=self.req_content.get(),
            exclude=self.exclude_var.get(),
        )
        try:
            kb = float(self.min_kb.get() or 0)
        except ValueError:
            kb = 0.0
        opts.min_size = int(kb * 1024)
        opts.mtime_after = self._parse_date(self.date_from.get(), end=False)
        opts.mtime_before = self._parse_date(self.date_to.get(), end=True)
        return Matcher(opts)

    @staticmethod
    def _parse_date(text: str, end: bool) -> float:
        text = (text or "").strip()
        if not text:
            return 0.0
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y-%m", "%Y"):
            try:
                st = time.strptime(text, fmt)
                ts = time.mktime(st)
                if fmt == "%Y-%m":
                    ts = time.mktime((st.tm_year, st.tm_mon + 1, 1, 0, 0, 0, 0, 0, -1)) - 1
                elif fmt == "%Y":
                    ts = time.mktime((st.tm_year + 1, 1, 1, 0, 0, 0, 0, 0, -1)) - 1
                elif end:
                    ts += 86399
                return ts
            except (ValueError, OverflowError):
                continue
        return 0.0

    def _do_search(self, auto: bool = False) -> None:
        if not self.index.records:
            if not auto:
                messagebox.showinfo(APP_NAME, "请先建立索引。")
            return

        matcher = self._build_matcher()
        if matcher.scope() in ("content", "both"):
            missing = len(self.index.paths_needing_content())
            running = self.content_thread is not None and self.content_thread.is_alive()
            if missing:
                if not running:
                    # Don't dead-end the user on an empty search: kick off the
                    # content pass and re-run this query when it finishes.
                    self.status.set(
                        f"需要内容匹配，正在解析 {human_int(missing)} 个文件的内容，完成后将自动重新搜索…"
                    )
                    self._start_content()
                return

        t0 = time.perf_counter()
        self._cancel_search.clear()
        self._located_where.clear()          # new query -> new addresses
        self._progress_start("检索索引")
        hits, total = self.index.search(matcher, limit=RESULT_LIMIT)
        elapsed_ms = (time.perf_counter() - t0) * 1000

        self.last_hits = hits
        self.last_total = total
        # Cell location re-reads each workbook, so it gets its own visible phase.
        locate_total = 0
        if hits and matcher.tokens and matcher.scope() in ("content", "both"):
            locate_total = min(len(hits), 500)
            self._progress_start("定位单元格", locate_total)
            self.status.set(f"正在定位命中单元格（0/{human_int(locate_total)}）…")
            self.update_idletasks()
        self._fill_results(hits, matcher, on_locate=lambda done, tot: self._progress_update(
            done, tot, "定位单元格"))
        self.btn_export.configure(state="normal" if hits else "disabled")

        extra = "" if total <= len(hits) else f"（仅显示前 {human_int(RESULT_LIMIT)} 条，请用“导出 CSV”获取全部）"
        located_note = ""
        if hits and locate_total:
            located_note = f" · 扫描 {human_int(locate_total)} 个文件定位到单元格（双击可跳转）"
        self._progress_done(f"完成 {human_int(total)} 命中")
        self.status.set(
            f"命中 {human_int(total)} 个文件{extra} · 查询耗时 {elapsed_ms:.1f} ms · "
            f"索引 {human_int(len(self.index.records))} 个文件{located_note}"
        )

    def _fill_results(self, hits, matcher: Matcher, on_locate=None,
                      locate_cells: bool = True) -> None:
        tree = self.tree
        tree.delete(*tree.get_children())
        first_needle = matcher.tokens[0] if matcher.tokens else ""
        # Locating cells means re-reading the workbook, so only do it for the
        # rows actually shown. Beyond this the preview panel locates on demand.
        locate_budget = 500
        keyword = first_needle
        can_locate = (
            locate_cells and bool(keyword) and matcher.scope() in ("content", "both")
        )
        locate_done = 0

        for i, (rec, _score, preview) in enumerate(hits):
            where = ""
            if matcher.match_name(rec.name):
                where = "文件名"
            content_hit = bool(
                rec.content_loaded and matcher.tokens and matcher.match_content_folded(rec.folded(), rec.content)
            )
            if content_hit:
                where = f"{where}+内容" if where else "内容"
            if not where:
                where = "筛选"

            show_preview = preview or (local_preview(rec.content or "", first_needle) if content_hit else "")

            if content_hit and can_locate and i < locate_budget:
                cached = self._located_where.get(rec.path)
                if cached:
                    where = cached
                else:
                    try:
                        located = find_cells(rec.path, keyword, case=self.case_var.get(), limit=3)
                    except Exception:
                        located = []
                    if located:
                        more = f" 等{len(located)}处" if len(located) > 1 else ""
                        where = f"{located[0]['sheet']}!{located[0]['ref']}{more}"
                    # Remember it so re-sorting redraws instantly instead of
                    # re-reading every workbook for the same answer.
                    self._located_where[rec.path] = where
                locate_done += 1
            elif content_hit:
                # Re-sort path: reuse whatever the search already resolved.
                where = self._located_where.get(rec.path, where)

            tree.insert(
                "",
                "end",
                iid=str(i),
                values=(rec.name, rec.folder, fmt_size(rec.size), fmt_time(rec.mtime), where, show_preview),
                tags=("odd",) if i % 2 else (),
            )

            # Keep the UI honest during long renders: refresh the progress bar in
            # small batches instead of freezing until the whole list is built.
            if (i + 1) % 100 == 0:
                if on_locate is not None and locate_done:
                    on_locate(locate_done, locate_budget)
                tree.update_idletasks()
                if self._cancel_search.is_set():
                    break

        if on_locate is not None and locate_done:
            on_locate(locate_done, locate_budget)
        if hits:
            tree.see(str(0))

    def _sort_by(self, col: str) -> None:
        if not self.last_hits:
            return
        self.sort_desc = not self.sort_desc if self.sort_col == col else False
        self.sort_col = col
        desc = self.sort_desc
        keyers = {
            "name": lambda t: t[0].name.casefold(),
            "folder": lambda t: t[0].folder.casefold(),
            "size": lambda t: t[0].size,
            "mtime": lambda t: t[0].mtime,
            "where": lambda t: t[1],
            "preview": lambda t: t[1],
        }
        key = keyers.get(col, lambda t: t[1])
        rows = sorted(self.last_hits, key=key, reverse=desc)
        matcher = self._build_matcher()
        self.last_hits = rows
        # Re-sorting redraws the list; locating cells again would re-read every
        # workbook, so the cached addresses from the search are reused instead.
        self._progress_start("重新排序")
        self._fill_results(rows, matcher, locate_cells=False)
        self._progress_done(f"已按 {col} 排序")

    # ------------------------------------------------------------- actions
    def _selected(self) -> list[Record]:
        out = []
        for iid in self.tree.selection():
            try:
                out.append(self.last_hits[int(iid)][0])
            except (ValueError, IndexError):
                continue
        return out

    def _popup(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self.menu.tk_popup(event.x_root, event.y_root)

    def _open_file(self) -> None:
        """Open the first selected file -- at the matching cell when we know it."""
        recs = self._selected()
        if not recs:
            return
        rec = recs[0]
        located = self._locate(rec)
        choice = self._open_with_id()
        if located:
            first = located[0]
            self.status.set(f"正在打开并定位到「{first['sheet']}」{first['ref']} …")
            self.update_idletasks()
            ok, detail = open_excel_at_cell(
                rec.path, first["sheet"], first["ref"], prefer=choice
            )
            if ok:
                self.status.set(
                    f"{detail}：{rec.name} → 工作表「{first['sheet']}」{first['ref']}"
                    + (f"（该文件共 {len(located)} 处命中）" if len(located) > 1 else "")
                )
                return
            self.status.set(f"{detail}；已改用系统默认程序打开文件。")
        elif self._keyword() and rec.ext in (".xlsx", ".xlsm"):
            self.status.set("该文件是文件名命中，内容中没有找到关键词；已直接打开文件。")

        try:
            open_in_explorer(rec.path, select=False)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"无法打开：\n{rec.path}\n\n{exc}")

    def _open_at_cell(self) -> None:
        """Explicitly jump to the first matching cell using the chosen program."""
        recs = self._selected()
        if not recs:
            return
        rec = recs[0]
        located = self._locate(rec)
        if not located:
            self.status.set("该文件没有单元格级命中，无法定位到具体单元格。")
            return
        first = located[0]
        self.status.set(f"正在定位到「{first['sheet']}」{first['ref']} …")
        self.update_idletasks()
        ok, detail = open_excel_at_cell(
            rec.path, first["sheet"], first["ref"], prefer=self._open_with_id()
        )
        if ok:
            self.status.set(f"{detail}，定位到工作表「{first['sheet']}」{first['ref']}")
        else:
            messagebox.showwarning(
                APP_NAME,
                f"{detail}\n\n"
                "可以在结果列表下方切换「打开方式」，或右键用「在资源管理器中定位」找到文件。\n"
                "如果只用 WPS：请确认 WPS 已安装且能正常启动表格。",
            )

    def _reveal(self) -> None:
        for rec in self._selected():
            try:
                open_in_explorer(rec.path, select=True)
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"无法定位：\n{rec.path}\n\n{exc}")
            break

    def _copy_paths(self) -> None:
        recs = self._selected()
        if not recs:
            return
        text = "\r\n".join(r.path for r in recs)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.set(f"已复制 {len(recs)} 条完整路径到剪贴板。")

    def _search_from_selection(self) -> None:
        recs = self._selected()
        if not recs:
            return
        self.query_var.set(os.path.splitext(recs[0].name)[0])
        self.req_name.set(True)
        self._do_search()

    def _export(self) -> None:
        if not self.last_hits:
            return
        path = filedialog.asksaveasfilename(
            title="导出搜索结果",
            defaultextension=".csv",
            initialfile=f"搜索结果_{time.strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not path:
            return
        matcher = self._build_matcher()
        # Resolving sheet/cell columns re-reads every workbook, so ask first.
        with_cells = False
        if matcher.tokens and matcher.scope() in ("content", "both"):
            with_cells = messagebox.askyesno(
                APP_NAME,
                f"是否在导出结果里包含「工作表 + 单元格地址」？\n\n"
                f"需要重新读取这 {len(self.last_hits)} 个表格文件，"
                f"大约多花 {max(1, len(self.last_hits) * 4 // 1000)} 秒。",
            )
        try:
            export_csv(path, self.last_hits, matcher,
                       with_cells=with_cells, case=self.case_var.get())
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"导出失败：{exc}")
            return
        note = "（含工作表/单元格列）" if with_cells else ""
        self.status.set(f"已导出 {len(self.last_hits)} 条结果{note}到 {path}")
        if messagebox.askyesno(APP_NAME, "导出完成，是否立即打开所在文件夹？"):
            open_in_explorer(path, select=True)

    # -------------------------------------------------------------- plumbing
    def _save_cfg(self) -> None:
        self.cfg.update(
            {
                "root": self.dir_var.get(),
                "mode": MODE_KEYS.get(self.mode_var.get(), "contains"),
                "case": self.case_var.get(),
                "req_content": self.req_content.get(),
                "exclude": self.exclude_var.get(),
                "ext_preset": self.ext_var.get(),
                "ext_custom": self.ext_custom.get(),
                "min_kb": self.min_kb.get(),
                "date_from": self.date_from.get(),
                "date_to": self.date_to.get(),
                "threads": self.threads.get(),
                "skip_hidden": self.skip_hidden.get(),
                "cache_index": self.cache_index.get(),
                "follow_links": self.follow_links.get(),
                "show_preview": self.show_preview.get(),
                "open_with": self._open_with_id(),
            }
        )
        save_config(self.cfg)

    def _stop_all(self) -> None:
        self.cancel_scan.set()
        self.cancel_content.set()
        self._cancel_search.set()
        engine = getattr(self, "_engine", None)
        if engine is not None:
            engine.cancel()
        self.progress_text.set("正在停止…")
        self.status.set("已请求停止…")

    def _pump(self) -> None:
        kind = None
        try:
            while True:
                msg = self.msg_q.get_nowait()
                kind = msg[0]
                trace(f"pump: {kind}")
                if kind == "progress":
                    self.stats_var.set(f"正在扫描… 已处理 {human_int(msg[1])} 个条目")
                elif kind == "scan_done":
                    self._on_scan_done(*msg[1:])
                elif kind == "content_progress":
                    _, done, total = msg
                    self.progress.configure(value=done, maximum=total)
                    self.status.set(f"内容索引 {human_int(done)}/{human_int(total)} …")
                elif kind == "content_done":
                    self._on_content_done(*msg[1:])
                elif kind == "error":
                    self._set_busy(False)
                    self.progress.stop()
                    self.status.set(msg[1])
                    messagebox.showerror(APP_NAME, msg[1])
        except queue.Empty:
            pass
        except Exception as exc:                          # pragma: no cover
            # Never let a UI error kill the pump loop: if `after` is not
            # rescheduled the window freezes with no way back.
            trace(f"pump FAILED handling {kind!r}: {exc!r}")
            self._set_busy(False)
            try:
                self.progress.stop()
            except tk.TclError:
                pass
            self.status.set(f"处理 {kind} 时出错：{exc}")
        finally:
            self.after(80, self._pump)

    def _on_close(self) -> None:
        self._stop_all()
        self._save_cfg()
        self.master.destroy()


def local_preview(text: str, needle: str, width: int = 120) -> str:
    if not text or not needle:
        return ""
    pos = text.casefold().find(needle.casefold())
    if pos < 0:
        return text[:width]
    start = max(0, pos - 40)
    end = min(len(text), pos + len(needle) + width)
    return ("..." if start else "") + text[start:end] + ("..." if end < len(text) else "")


def main() -> int:
    trace(f"main: start argv={sys.argv[1:]!r} frozen={getattr(sys, 'frozen', False)} portable={PORTABLE}")
    setup_console()
    root = tk.Tk()
    suffix = "  [U盘便携版]" if PORTABLE else ""
    root.title(f"{APP_NAME} v{APP_VERSION} — 本地 Excel 极速检索{suffix}")
    root.geometry("1280x780")
    root.minsize(980, 620)
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    try:
        root.iconbitmap(default="")
    except tk.TclError:
        pass
    app = ExcelFinderApp(root)
    trace("main: app constructed")

    # Allow `ExcelFinder.exe "D:\data" 关键词` from the command line / shortcuts.
    argv = [a for a in sys.argv[1:] if a.strip()]
    if argv:
        if os.path.isdir(argv[0]):
            app.dir_var.set(os.path.normpath(argv[0]))
            if len(argv) > 1:
                app.query_var.set(" ".join(argv[1:]))
        else:
            app.query_var.set(" ".join(argv))
        trace(f"main: after argv -> dir={app.dir_var.get()!r} query={app.query_var.get()!r}")
        if os.path.isdir(app.dir_var.get()):
            app._start_scan()

    trace("main: entering mainloop")
    root.mainloop()
    trace("main: mainloop ended")
    return 0


if __name__ == "__main__":
    sys.exit(main())
