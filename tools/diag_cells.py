# -*- coding: utf-8 -*-
"""Check cell-level location quality and cost on the test corpus."""
from __future__ import annotations

import json
import os
import sys
import time
import zipfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from excelfinder_core import ScanEngine, _sheet_member_map, find_cells, iter_cells  # noqa: E402


def main() -> int:
    root = sys.argv[1] if len(sys.argv) > 1 else "testdata"
    keyword = sys.argv[2] if len(sys.argv) > 2 else "张伟"

    eng = ScanEngine(root, threads=4)
    recs, _sec, _r = eng.run()
    xlsx = [r for r in recs if r.ext == ".xlsx"]
    print(f"corpus: {len(recs)} files ({len(xlsx)} xlsx)")

    target = xlsx[0]
    print(f"\nsample file: {target.name}")
    with zipfile.ZipFile(target.path) as zf:
        print("sheet member map:", _sheet_member_map(zf))
    cells = list(iter_cells(target.path))
    print(f"cells parsed: {len(cells)}")
    print("first 6 cells:", cells[:6])

    hits = find_cells(target.path, keyword, limit=3)
    print(f"\nfind_cells('{keyword}') -> {len(hits)} hit(s)")
    print(json.dumps(hits, ensure_ascii=False, indent=1)[:900])

    # cost of the second pass
    sample = xlsx[:30]
    t0 = time.perf_counter()
    n = 0
    for r in sample:
        n += len(find_cells(r.path, keyword, limit=3))
    el = time.perf_counter() - t0
    print(f"\nlocate {len(sample)} files: {el:.2f}s ({el/len(sample)*1000:.1f} ms/file), {n} hits total")

    # how many keystrokes before a workbook with no keyword is scanned
    t0 = time.perf_counter()
    for r in sample:
        find_cells(r.path, "ZZZ_不存在_ZZZ", limit=1)
    el = time.perf_counter() - t0
    print(f"locate {len(sample)} files (no match): {el:.2f}s ({el/len(sample)*1000:.1f} ms/file)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
