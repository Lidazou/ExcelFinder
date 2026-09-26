# -*- coding: utf-8 -*-
"""Assemble and verify the U盘便携版 (portable USB) package.

Builds the layout in ``portable/``:
    ExcelFinder.exe           GUI (self-contained)
    ExcelFinder-cli.exe       console diagnostics / scriptable search
    启动 ExcelFinder.bat       launcher
    自检.bat                   one-click self check
    使用说明.md                user guide
    portable.marker           switches the app into portable mode
    data/                     config + index cache (created on first run)
    python/                   embeddable CPython fallback (tkinter grafted)
    src/                      engine sources

Then it verifies the package and optionally zips it.

Usage: python tools/make_portable.py [--zip] [--skip-python]
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")
DIST = os.path.join(ROOT, "dist")
PKG = os.path.join(ROOT, "portable")

FILES = ["ExcelFinder.exe", "ExcelFinder-cli.exe"]
SCRIPTS = ["启动 ExcelFinder.bat", "自检.bat", "使用说明.md"]
SOURCES = ["excelfinder_core.py", "excelfinder_gui.py", "excelfinder_cli.py"]

MARKER = """ExcelFinder portable mode marker.
Delete this file to switch back to storing settings in %APPDATA%\\ExcelFinder.
"""


def hr(title: str) -> None:
    print("\n" + "=" * 70 + f"\n{title}\n" + "=" * 70)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip", action="store_true", help="also produce a distributable zip")
    ap.add_argument("--skip-python", action="store_true", help="omit the embeddable runtime")
    args = ap.parse_args()
    failures: list[str] = []

    hr("1. 检查构建产物")
    for name in FILES:
        path = os.path.join(DIST, name)
        ok = os.path.exists(path)
        print(f"  [{'OK ' if ok else 'BAD'}] dist/{name}"
              f"{f'  ({os.path.getsize(path)/1024/1024:.2f} MB)' if ok else '  <缺失，请先运行 build.py>'}")
        if not ok:
            failures.append(f"missing {name}")
    if failures:
        return 1

    hr("2. 组装 portable/ 目录")
    os.makedirs(PKG, exist_ok=True)
    os.makedirs(os.path.join(PKG, "src"), exist_ok=True)

    for name in FILES:
        shutil.copy2(os.path.join(DIST, name), os.path.join(PKG, name))
        print(f"  + {name}")

    # Launchers and docs are authored at the repo root (next to the EXEs), so
    # copy them in rather than expecting them to already sit in the package.
    for name in SCRIPTS:
        src = os.path.join(ROOT, name)
        dst = os.path.join(PKG, name)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"  + {name}")
        elif os.path.exists(dst):
            print(f"  = {name} (已有)")
        else:
            print(f"  ! {name} 缺失")
            failures.append(f"missing script {name}")

    for name in SOURCES:
        shutil.copy2(os.path.join(SRC, name), os.path.join(PKG, "src", name))
    print(f"  + src/ ({len(SOURCES)} 个源文件)")

    # Fallback GUI entry point for the bundled runtime.
    fallback = os.path.join(ROOT, "run_gui.py")
    if os.path.exists(fallback):
        shutil.copy2(fallback, os.path.join(PKG, "run_gui.py"))
        print("  + run_gui.py （备用启动入口）")

    marker = os.path.join(PKG, "portable.marker")
    with open(marker, "w", encoding="utf-8") as fh:
        fh.write(MARKER)
    print("  + portable.marker （启用便携模式）")

    if not args.skip_python:
        # The embeddable runtime is staged under embed-python/ in the repo and
        # copied in as python/ so the fallback launcher's path stays stable.
        src_py = os.path.join(ROOT, "embed-python")
        dst_py = os.path.join(PKG, "python")
        if os.path.isdir(src_py):
            if os.path.isdir(dst_py):
                shutil.rmtree(dst_py, ignore_errors=True)
            shutil.copytree(src_py, dst_py)
            print("  + python/ (内置解释器，从 embed-python/ 复制)")
        else:
            print("  ! python/ 缺失，可先运行 tools/fetch_python_embed.py embed-python")

    hr("3. 验证便携模式与界面")
    data_dir = os.path.join(PKG, "data")
    if os.path.isdir(data_dir):
        print(f"  data/ 已存在（{len(os.listdir(data_dir))} 项），保留既有配置")
    else:
        os.makedirs(data_dir, exist_ok=True)
        print("  data/ 已创建（首次运行会在此写入配置和索引）")

    # Portable detection works from source too: app_dir() derives from the
    # module location when not frozen, so loading src/ from the portable package
    # exercises exactly the same code path the shipped EXE uses.
    probe_code = (
        "import sys, os;"
        "sys.path.insert(0, os.path.join(%r, 'src'));"
        "from excelfinder_core import state_dir, is_portable, app_dir;"
        "print('app_dir   =', app_dir());"
        "print('state_dir =', state_dir());"
        "print('portable  =', is_portable());"
        "print('DATA_OK   =', state_dir() == os.path.join(%r, 'data'))"
    ) % (PKG, PKG)
    probe = subprocess.run([sys.executable, "-c", probe_code], capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
    out = (probe.stdout or "").strip()
    print(out or (probe.stderr or "")[:400])
    if "DATA_OK   = True" not in out:
        failures.append("portable mode detection failed")
        print("  [BAD] 便携模式检测失败：配置不会写到 U 盘的 data/ 目录")

    hr("4. 打包体积统计")
    total = 0
    breakdown: list[tuple[str, int]] = []
    for entry in sorted(os.listdir(PKG)):
        path = os.path.join(PKG, entry)
        if os.path.isdir(path):
            size = sum(
                os.path.getsize(os.path.join(b, f))
                for b, _d, fs in os.walk(path) for f in fs
            )
        else:
            size = os.path.getsize(path)
        total += size
        breakdown.append((entry, size))
    for entry, size in breakdown:
        print(f"  {entry:<28} {size/1024/1024:8.2f} MB")
    print(f"  {'合计':<26} {total/1024/1024:8.2f} MB")

    if args.zip:
        hr("5. 生成可分发的 zip")
        out = os.path.join(DIST, "ExcelFinder-portable.zip")
        # Everything goes under one top-level folder so extracting never litters
        # the destination (Explorer's "Extract All" would otherwise dump the
        # exe, python/ and src/ straight into the chosen directory).
        root_name = "ExcelFinder-portable"
        # Ship a pristine package: the local index cache would otherwise leak
        # absolute paths from the build machine into the distributable.
        skip = os.path.normcase(os.path.join(PKG, "data"))
        count = 0
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
            # Keep an (empty) data dir so portable mode is obvious on first run.
            zf.writestr(f"{root_name}/data/", "")
            for base, _dirs, files in os.walk(PKG):
                if os.path.normcase(base).startswith(skip):
                    continue
                for f in files:
                    full = os.path.join(base, f)
                    arc = os.path.join(root_name, os.path.relpath(full, PKG))
                    zf.write(full, arc)
                    count += 1
        print(f"  {out}  ({os.path.getsize(out)/1024/1024:.2f} MB, {count} 个文件)")
        print(f"  解压后得到 {root_name}\\ 文件夹，双击其中的 ExcelFinder.exe 即可")
        print("  （已排除 data/ 本地索引缓存，首次运行会自动重建）")

    hr("结果")
    if failures:
        print(f"打包失败：{failures}")
        return 1
    print("便携包组装完成。")
    print(f"位置：{PKG}")
    print("拷贝到U盘后双击 ExcelFinder.exe 即可使用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
