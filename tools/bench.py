# -*- coding: utf-8 -*-
"""Headless correctness + performance harness for the ExcelFinder engine."""
from __future__ import annotations

import collections
import os
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))

from excelfinder_core import (  # noqa: E402
    MatchOpts,
    Matcher,
    ScanEngine,
    Index,
    extract_text,
    fmt_size,
    index_content,
)


def hr(title: str) -> None:
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def main() -> int:
    root = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "..", "testdata"
    )
    threads = int(sys.argv[2]) if len(sys.argv) > 2 else 16
    failures = 0

    hr(f"1. 目录扫描 (threads={threads})")
    eng = ScanEngine(root, threads=threads)
    t0 = time.perf_counter()
    recs, sec, reused = eng.run()
    total_bytes = sum(r.size for r in recs)
    print(f"files      : {len(recs)}")
    print(f"bytes      : {fmt_size(total_bytes)}")
    print(f"wall time  : {sec:.3f}s  ({len(recs)/max(sec,1e-9):,.0f} files/s)")
    print(f"reused     : {reused}")
    if eng.errors:
        print(f"walk errors: {len(eng.errors)} (first: {eng.errors[0][:100]})")
    print("by ext     :", dict(collections.Counter(r.ext for r in recs).most_common()))

    hr("2. 名称匹配正确性")
    checks = [
        ("2023 销售部", {"req_name": True}, None),
        ("客户", {"req_name": True}, None),
        ("ZZZ_不存在_ZZZ", {"req_name": True}, 0),
        ("2023", {"req_name": True, "mode": "prefix"}, None),
        ("*预算*.xlsx", {"req_name": True, "mode": "wildcard"}, None),
        (r"^2024.*_\d{3}\.xlsx$", {"req_name": True, "mode": "regex"}, None),
        ("xlsx", {"req_name": True, "ext_filter": ".xlsx"}, None),
    ]
    idx = Index()
    idx.replace_all(recs, root, sec)
    for query, kwargs, expect in checks:
        m = Matcher(MatchOpts(query=query, **kwargs))
        t0 = time.perf_counter()
        hits, total = idx.search(m)
        ms = (time.perf_counter() - t0) * 1000
        ok = "OK " if (expect is None or total == expect) else "BAD"
        if ok == "BAD":
            failures += 1
        print(f"  [{ok}] {query!r:<28} -> {total:>6} hits in {ms:6.2f} ms   (top: {hits[0][0].name if hits else '-'})")

    hr("3. 内容解析 (xlsx/csv)")
    sample = next((r for r in recs if r.ext == ".xlsx"), None)
    if sample:
        t0 = time.perf_counter()
        text = extract_text(sample.path, sample.size)
        ms = (time.perf_counter() - t0) * 1000
        print(f"sample    : {sample.name}")
        print(f"extract   : {ms:.2f} ms, {len(text)} chars")
        print(f"head      : {text[:180]}")
        if not text:
            print("  BAD: empty extraction")
            failures += 1
    else:
        print("  BAD: no .xlsx in corpus")
        failures += 1

    hr("4. 全量内容索引")
    t0 = time.perf_counter()
    done, failed = index_content(recs, threads=threads)
    elapsed = time.perf_counter() - t0
    print(f"extracted  : {done}, failed {failed} in {elapsed:.2f}s ({done/max(elapsed,1e-9):,.0f} files/s)")
    if failed:
        bad = [r for r in recs if r.error][:3]
        for r in bad:
            print(f"   ! {r.name}: {r.error}")
        failures += 1
    chars = sum(len(r.content or "") for r in recs)
    print(f"text kept  : {chars:,} chars ({chars/1024/1024:.2f} MiB)")

    hr("5. 内容检索速度")
    for query in ["扫地机器人", "财务部 已驳回", "ZZZ_不存在_ZZZ"]:
        m = Matcher(MatchOpts(query=query, req_name=False, req_content=True))
        t0 = time.perf_counter()
        hits, total = idx.search(m)
        ms = (time.perf_counter() - t0) * 1000
        print(f"  {query!r:<22} -> {total:>6} hits in {ms:7.1f} ms   top={hits[0][0].name if hits else '-'}")

    hr("6. 增量扫描 (reuse cache)")
    eng2 = ScanEngine(root, threads=threads, reuse=recs)
    t0 = time.perf_counter()
    recs2, sec2, reused2 = eng2.run()
    print(f"wall time  : {sec2:.3f}s (was {sec:.3f}s), reused {reused2}/{len(recs2)}")
    if reused2 != len(recs2):
        print("  note: not every record reused (files may have changed)")
    if len(recs2) != len(recs):
        print(f"  BAD: record count drifted {len(recs)} -> {len(recs2)}")
        failures += 1

    hr("7. 索引持久化")
    tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_index_test.json")
    idx.save(tmp)
    size = os.path.getsize(tmp)
    idx2 = Index()
    ok = idx2.load(tmp)
    print(f"save/load  : {'OK' if ok else 'BAD'}, {fmt_size(size)}, {len(idx2.records)} records")
    if not ok or len(idx2.records) != len(recs):
        failures += 1
    t0 = time.perf_counter()
    m = Matcher(MatchOpts(query="扫地机器人", req_name=False, req_content=True))
    hits3, total3 = idx2.search(m)
    print(f"query after reload: {total3} hits in {(time.perf_counter()-t0)*1000:.1f} ms")
    # Derive the expectation from the index itself rather than hard-coding a
    # number for one corpus -- any test-data change would otherwise look like a
    # regression.
    expected = sum(1 for r in idx.records if "扫地机器人" in (r.content or ""))
    if total3 != expected:
        print(f"  BAD: expected {expected} content hits, got {total3}")
        failures += 1
    else:
        print(f"  consistent with in-memory index ({expected} content hits)")
    os.remove(tmp)

    hr("RESULT")
    print(f"{'ALL CHECKS PASSED' if failures == 0 else f'{failures} CHECK(S) FAILED'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
