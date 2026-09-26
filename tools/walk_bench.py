# -*- coding: utf-8 -*-
"""Micro-benchmark of the directory-walk inner loop variants."""
from __future__ import annotations

import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor

ROOT = sys.argv[1] if len(sys.argv) > 1 else "testdata_big"
EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xls", ".xlt", ".xla",
        ".xlsb", ".csv", ".tsv", ".ods", ".txt", ".md", ".log"}


def variant_splitext(top):
    out = []
    stack = [top]
    while stack:
        cur = stack.pop()
        try:
            with os.scandir(cur) as it:
                for e in it:
                    if e.is_dir(follow_symlinks=False):
                        stack.append(e.path)
                        continue
                    if not e.is_file(follow_symlinks=False):
                        continue
                    if os.path.splitext(e.name)[1].lower() not in EXTS:
                        continue
                    st = e.stat(follow_symlinks=False)
                    out.append((e.path, e.name, st.st_size, st.st_mtime))
        except OSError:
            pass
    return out


def variant_rpartition(top):
    """rpartition + endswith fast-path: avoids building a list per entry."""
    out = []
    stack = [top]
    push = stack.append
    pop = stack.pop
    exts = EXTS
    while stack:
        cur = pop()
        try:
            with os.scandir(cur) as it:
                for e in it:
                    name = e.name
                    if e.is_dir(follow_symlinks=False):
                        push(e.path)
                        continue
                    dot = name.rfind(".")
                    if dot < 0:
                        continue
                    if name[dot:].lower() not in exts:
                        continue
                    st = e.stat(follow_symlinks=False)
                    out.append((e.path, name, st.st_size, st.st_mtime))
        except OSError:
            pass
    return out


def variant_d_type(top):
    """Use the cached dirent type instead of two is_*() syscalls per entry."""
    out = []
    stack = [top]
    push = stack.append
    pop = stack.pop
    exts = EXTS
    DT_DIR = 1 << 0
    while stack:
        cur = pop()
        try:
            with os.scandir(cur) as it:
                for e in it:
                    name = e.name
                    try:
                        is_dir = e.is_dir(follow_symlinks=False)
                    except OSError:
                        continue
                    if is_dir:
                        push(e.path)
                        continue
                    dot = name.rfind(".")
                    if dot < 0 or name[dot:].lower() not in exts:
                        continue
                    try:
                        st = e.stat(follow_symlinks=False)
                    except OSError:
                        continue
                    out.append((e.path, name, st.st_size, st.st_mtime))
        except OSError:
            pass
    return out


def jobs_for(root):
    jobs = []
    with os.scandir(root) as it:
        for e in it:
            if e.is_dir(follow_symlinks=False):
                jobs.append(e.path)
            elif e.is_file(follow_symlinks=False):
                jobs.append(e.path)
    return jobs


def run(fn, jobs, threads, rounds=3):
    best = 9e9
    for _ in range(rounds):
        t0 = time.perf_counter()
        out = []
        if threads == 1:
            for j in jobs:
                out.extend(fn(j))
        else:
            with ThreadPoolExecutor(max_workers=threads) as pool:
                for chunk in pool.map(fn, jobs):
                    out.extend(chunk)
        best = min(best, time.perf_counter() - t0)
    return best, len(out)


def main():
    jobs = [j for j in jobs_for(ROOT) if os.path.isdir(j)]
    print(f"root={ROOT}  top-level dirs={len(jobs)}")
    for name, fn in (("splitext", variant_splitext), ("rpartition", variant_rpartition), ("d_type", variant_d_type)):
        for threads in (1, 2, 4, 8):
            ms, n = run(fn, jobs, threads)
            print(f"  {name:<12} threads={threads:<3} {ms*1000:7.1f} ms  ({n} matched)")
        print()


if __name__ == "__main__":
    main()
