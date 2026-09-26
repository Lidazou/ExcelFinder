# -*- coding: utf-8 -*-
"""Compare ExcelFinder against the search facilities Windows ships with.

Baselines measured:
  1. ``where /r``          - the classic recursive name search built into cmd.exe
  2. ``Get-ChildItem -Recurse`` - PowerShell's recursive filter
  3. Windows Search index (OLE DB ``Search.CollatorDSO``) - what the Explorer
     search box actually uses. Only counted when the indexer has covered the
     target folder; otherwise reported as unavailable rather than faked.

Usage: python compare_windows.py <root> [keyword]
"""
from __future__ import annotations

import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from excelfinder_core import Index, MatchOpts, Matcher, ScanEngine  # noqa: E402


def timed(label: str, fn):
    t0 = time.perf_counter()
    try:
        result = fn()
        err = None
    except Exception as exc:                                  # noqa: BLE001
        result, err = None, exc
    ms = (time.perf_counter() - t0) * 1000
    return label, ms, result, err


# --------------------------------------------------------------------------
def where_exe(root: str, pattern: str):
    """`where /r root *pattern*` -- cmd.exe's built-in recursive search."""
    p = subprocess.run(
        ["where", "/r", root, f"*{pattern}*"],
        capture_output=True, text=True, errors="replace",
    )
    lines = [ln for ln in (p.stdout or "").splitlines() if ln.strip()]
    return len(lines), lines


def powershell_gci(root: str, pattern: str):
    """PowerShell recursive name filter -- the common 'fast' workaround."""
    cmd = (
        f"Get-ChildItem -LiteralPath '{root}' -Recurse -File -Force "
        f"-ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.Name -like '*{pattern}*' }} | "
        f"Select-Object -ExpandProperty FullName"
    )
    p = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", cmd],
        capture_output=True, text=True, errors="replace",
    )
    lines = [ln for ln in (p.stdout or "").splitlines() if ln.strip()]
    return len(lines), lines


def windows_search_index(root: str, pattern: str):
    """Query the real Windows Search index through OLE DB via .NET interop."""
    script = f"""
$ErrorActionPreference = 'Stop'
$conn = New-Object -ComObject ADODB.Connection
$conn.Open('Provider=Search.CollatorDSO;Extended Properties="Application=Windows";')
$sql = "SELECT System.ItemPathDisplay FROM SYSTEMINDEX WHERE SCOPE='file:{root}' AND System.FileName LIKE '%{pattern}%'"
$rs = $conn.Execute($sql)
$n = 0
while (-not $rs.EOF) {{ $n++; $rs.MoveNext() }}
Write-Output "COUNT=$n"
$rs.Close(); $conn.Close()
"""
    p = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, errors="replace", timeout=300,
    )
    out = (p.stdout or "").strip()
    if "COUNT=" not in out:
        raise RuntimeError((p.stderr or out or "no output").strip().splitlines()[0][:160])
    return int(out.split("COUNT=")[1].split()[0]), []


# --------------------------------------------------------------------------
def main() -> int:
    root = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.abspath(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "testdata")
    )
    keyword = sys.argv[2] if len(sys.argv) > 2 else "预算"

    n_files = sum(len(f) for _b, _d, f in os.walk(root))
    print(f"corpus : {root}")
    print(f"entries: {n_files} files on disk   keyword: {keyword!r}")
    print()

    rows = []

    # --- Windows baselines ------------------------------------------------
    for label, fn in (
        ("where.exe /r (CMD 递归搜索)", lambda: where_exe(root, keyword)),
        ("PowerShell Get-ChildItem -Recurse", lambda: powershell_gci(root, keyword)),
        ("Windows Search 索引 (Explorer 搜索)", lambda: windows_search_index(root, keyword)),
    ):
        name, ms, result, err = timed(label, fn)
        if err is not None:
            rows.append((name, ms, None, f"不可用: {err}"))
        else:
            count, _ = result
            rows.append((name, ms, count, ""))

    # --- ExcelFinder ------------------------------------------------------
    eng = ScanEngine(root, threads=16)
    t0 = time.perf_counter()
    records, _s, _r = eng.run()
    scan_ms = (time.perf_counter() - t0) * 1000

    idx = Index()
    idx.replace_all(records, root, _s)

    matcher = Matcher(MatchOpts(query=keyword, req_name=True))
    t0 = time.perf_counter()
    hits, total = idx.search(matcher)
    query_ms = (time.perf_counter() - t0) * 1000

    rows.append(("ExcelFinder（首次：扫描目录）", scan_ms + query_ms, total, ""))
    rows.append(("ExcelFinder（已建索引：纯检索）", query_ms, total, ""))

    # warm re-scan through the incremental cache
    eng2 = ScanEngine(root, threads=16, reuse=records)
    t0 = time.perf_counter()
    eng2.run()
    rescan_ms = (time.perf_counter() - t0) * 1000
    rows.append(("ExcelFinder（增量复扫）", rescan_ms + query_ms, total, ""))

    # --- report -----------------------------------------------------------
    base_first = next((r[1] for r in rows if r[0].startswith("where.exe") and r[2] is not None), None)
    base_ps = next((r[1] for r in rows if r[0].startswith("PowerShell") and r[2] is not None), None)

    print(f"{'方案':<38} {'耗时(ms)':>10}  {'命中':>8}   {'相对加速':>10}  备注")
    print("-" * 96)
    for name, ms, count, note in rows:
        speed = ""
        if count is not None and ms > 0:
            if name.startswith("ExcelFinder"):
                ref = base_first
            else:
                ref = base_first
            if ref:
                speed = f"{ref / ms:,.1f}x"
        shown = "-" if count is None else f"{count:,}"
        print(f"{name:<38} {ms:>10.1f}  {shown:>8}   {speed:>10}  {note}")
    print()
    print(f"基线: where.exe /r = {base_first:,.1f} ms  ·  PowerShell = {base_ps:,.1f} ms"
          if base_first and base_ps else "")
    return 0


if __name__ == "__main__":
    sys.exit(main())
