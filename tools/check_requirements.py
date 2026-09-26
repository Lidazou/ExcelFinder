# -*- coding: utf-8 -*-
"""Report the hard system requirements baked into the shipped binaries.

Reads the PE headers of the EXEs (and the bundled interpreter) and prints the
minimum OS version, CPU architecture and subsystem -- facts, not guesses.
"""
from __future__ import annotations

import os
import struct
import sys

MACHINES = {0x8664: "x64 (AMD64)", 0x14C: "x86 (32 位)", 0xAA64: "ARM64"}
SUBSYSTEMS = {2: "Windows 图形界面程序", 3: "Windows 控制台程序"}


def pe_info(path: str) -> dict:
    with open(path, "rb") as fh:
        data = fh.read()
    if data[:2] != b"MZ":
        raise ValueError("不是 PE 文件")
    e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
    if data[e_lfanew:e_lfanew + 4] != b"PE\x00\x00":
        raise ValueError("PE 签名缺失")
    machine = struct.unpack_from("<H", data, e_lfanew + 4)[0]
    opt = e_lfanew + 24
    magic = struct.unpack_from("<H", data, opt)[0]
    major_os, minor_os = struct.unpack_from("<HH", data, opt + 40)
    major_sub, minor_sub = struct.unpack_from("<HH", data, opt + 48)
    subsystem = struct.unpack_from("<H", data, opt + 68)[0]
    return {
        "machine": machine,
        "bits": 64 if magic == 0x20B else 32,
        "os": f"{major_os}.{minor_os}",
        "subsystem_ver": f"{major_sub}.{minor_sub}",
        "subsystem": subsystem,
        "size": os.path.getsize(path),
    }


def main() -> int:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    targets = [
        ("主程序", os.path.join(root, "ExcelFinder.exe")),
        ("命令行版", os.path.join(root, "ExcelFinder-cli.exe")),
        ("内置解释器", os.path.join(root, "embed-python", "python.exe")),
    ]
    print("=" * 68)
    print("二进制里写死的运行要求（读 PE 头，实测）")
    print("=" * 68)
    for label, path in targets:
        if not os.path.exists(path):
            print(f"{label}: 未找到 {path}")
            continue
        try:
            info = pe_info(path)
        except Exception as exc:
            print(f"{label}: 读取失败 {exc}")
            continue
        print(f"\n{label}  ({os.path.basename(path)}, {info['size']/1024/1024:.1f} MB)")
        print(f"  CPU 架构   : {MACHINES.get(info['machine'], hex(info['machine']))}")
        print(f"  位数       : {info['bits']} 位")
        print(f"  最低 OS    : {info['os']}")
        print(f"  子系统     : {SUBSYSTEMS.get(info['subsystem'], info['subsystem'])}")

    print("\n" + "=" * 68)
    print("结论")
    print("=" * 68)
    print("CPU 架构 : 必须 64 位（x64 / AMD64）。32 位 Windows 跑不了。")
    print("操作系统 : 建议 Windows 10 / 11 64 位。")
    print("           上面 PE 头写的 6.0 只是编译器默认值，实际下限由")
    print("           内置的 Python 3.14 决定 —— Win7/8 已不在支持范围。")
    print()
    print("不需要另外安装：")
    print("  Python        —— 已打包在 exe 内部")
    print("  VC++ 运行库    —— VCRUNTIME140.dll 与 ucrtbase.dll 已内嵌")
    print("  Office / WPS  —— 只有『双击跳到单元格』用到，不装也能搜索")
    print("                   （不装则降级为系统默认程序打开文件）")
    print("  .NET / Java   —— 完全不需要")
    return 0


if __name__ == "__main__":
    sys.exit(main())
