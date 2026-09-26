# -*- coding: utf-8 -*-
"""Explain why a folder takes far more space than its byte size.

Windows reports "Size" and "Size on disk"; the gap is cluster slack. A USB stick
formatted with huge allocation units (exFAT in the megabytes) turns a 50 MB
folder into over a gigabyte of consumed space.

Run with no arguments to analyse every ready drive, or point it at a folder:

    python tools/check_usage.py
    python tools/check_usage.py "E:\\ExcelFinder-U盘版"
"""
from __future__ import annotations

import ctypes
import os
import sys


class _FreeSpace(ctypes.Structure):
    pass


def cluster_size(root: str) -> int:
    """Allocation unit size of the volume holding `root`."""
    sectors = ctypes.c_ulong()
    per_sector = ctypes.c_ulong()
    free = ctypes.c_ulong()
    total = ctypes.c_ulong()
    drive = os.path.splitdrive(os.path.abspath(root))[0] + "\\"
    ok = ctypes.windll.kernel32.GetDiskFreeSpaceW(
        ctypes.c_wchar_p(drive),
        ctypes.byref(sectors), ctypes.byref(per_sector),
        ctypes.byref(free), ctypes.byref(total),
    )
    if not ok:
        return 4096
    return max(512, (sectors.value or 8) * (per_sector.value or 512))


def analyse(path: str) -> None:
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        print(f"  {path} 不是目录")
        return
    cluster = cluster_size(path)
    sizes: list[int] = []
    for base, _dirs, files in os.walk(path):
        for f in files:
            try:
                sizes.append(os.path.getsize(os.path.join(base, f)))
            except OSError:
                pass
    if not sizes:
        print(f"  {path}: 空目录")
        return
    total = sum(sizes)
    on_disk = sum(((s + cluster - 1) // cluster) * cluster for s in sizes)
    waste = on_disk - total
    avg = total / len(sizes)
    drive = os.path.splitdrive(path)[0] or "?"

    print(f"\n目录  : {path}")
    print(f"所在卷: {drive}\\   簇大小 = {cluster:,} 字节 ({cluster/1024:,.1f} KB)")
    print(f"文件数: {len(sizes):,}")
    print(f"字节数: {total:,} ({total/1024/1024:,.1f} MB)")
    print(f"占用  : {on_disk:,} ({on_disk/1024/1024:,.1f} MB)")
    print(f"浪费  : {waste/1024/1024:,.1f} MB  ({waste/max(on_disk,1)*100:.0f}% 被簇尾空隙吃掉)")
    print(f"平均文件: {avg/1024:,.0f} KB")

    if waste > total:
        # Each file wastes on average half a cluster.
        need = int((total / max(len(sizes), 1)) + cluster) * len(sizes)
        print("\n诊断: 簇过大 —— 每个文件都占满整数个簇，小文件浪费严重。")
        print(f"      本目录平均文件 {avg/1024:,.0f} KB，而一个簇就有 {cluster/1024:,.0f} KB。")
        print("\n建议: 把U盘重新格式化为 exFAT，分配单元大小选 32 KB 或 64 KB。")
        print("      Windows 默认的 exFAT 大簇（MB 级）不适合放大量小文件。")
        print(f"      按 32 KB 簇估算，同样这些文件只需约 "
              f"{sum(((s + 32768 - 1)//32768)*32768 for s in sizes)/1024/1024:,.1f} MB。")
    else:
        print("\n诊断: 簇大小正常，占用与字节数接近。")

    print("\n最大的 8 个文件:")
    for size, name in sorted(
        ((os.path.getsize(os.path.join(b, f)), os.path.join(b, f))
         for b, _d, fs in os.walk(path) for f in fs), reverse=True
    )[:8]:
        print(f"  {size/1024/1024:8.2f} MB  {os.path.relpath(name, path)}")

    print("\n文件数最多的 5 个子目录:")
    counts: dict[str, int] = {}
    for base, _dirs, files in os.walk(path):
        rel = os.path.relpath(base, path).split(os.sep)[0]
        counts[rel] = counts.get(rel, 0) + len(files)
    for name, count in sorted(counts.items(), key=lambda kv: -kv[1])[:5]:
        print(f"  {count:6,} 个文件  {name}")


def main() -> int:
    if len(sys.argv) > 1:
        analyse(sys.argv[1])
        return 0

    print("=" * 72)
    print("驱动器概览")
    print("=" * 72)
    for letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
        root = f"{letter}:\\"
        if not os.path.exists(root):
            continue
        try:
            total = ctypes.c_ulonglong(0)
            free = ctypes.c_ulonglong(0)
            ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                ctypes.c_wchar_p(root), None, ctypes.byref(total), ctypes.byref(free))
            cluster = cluster_size(root)
            kind = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(root))
            names = {2: "可移动", 3: "固定", 4: "网络", 5: "光驱", 6: "内存盘"}
            print(f"  {root}  {names.get(kind, kind):<6} 簇={cluster/1024:>8,.1f} KB"
                  f"  容量={total.value/1024**3:6.1f} GB  可用={free.value/1024**3:6.1f} GB")
        except Exception as exc:
            print(f"  {root}  查询失败: {exc}")

    print("\n提示: 把U盘插上后运行 `python tools/check_usage.py E:\\` 查看它的簇大小。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
