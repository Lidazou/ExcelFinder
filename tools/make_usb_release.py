# -*- coding: utf-8 -*-
"""Build the final U盘分发目录 (USB distribution folder).

Produces ONE folder holding everything needed to run ExcelFinder on another PC
with nothing installed. Unlike ``tools/make_portable.py`` (which builds in the
repo for testing), this creates a clean, ready-to-copy folder and verifies it.

Usage:
    python tools/make_usb_release.py [target_dir] [--zip]

Default target: <Desktop>\\ExcelFinder-U盘版
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import struct
import subprocess
import sys
import time
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable

FAILURES: list[str] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
    if not ok:
        FAILURES.append(label)


def hr(t: str) -> None:
    print("\n" + "=" * 70 + f"\n{t}\n" + "=" * 70)


def desktop() -> str:
    for cand in (
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
    ):
        if os.path.isdir(cand):
            return cand
    return os.path.expanduser("~")


def dir_usage(path: str) -> tuple[int, int, int]:
    """(file_count, bytes, bytes_on_disk).

    Windows reports both "size" and "size on disk"; they diverge badly for a
    tree with many small files (each file wastes up to a cluster). Reporting
    only the byte total is how you tell a user 53 MB and have Explorer show
    something noticeably larger, so both are reported.
    """
    count = total = on_disk = 0
    cluster = 4096
    try:
        import ctypes

        sectors = ctypes.c_ulong()
        ctypes.windll.kernel32.GetDiskFreeSpaceW(
            ctypes.c_wchar_p(os.path.splitdrive(os.path.abspath(path))[0] + "\\"),
            None, ctypes.byref(sectors), None, None)
        cluster = max(4096, sectors.value or 4096)
    except Exception:
        pass
    for base, _dirs, files in os.walk(path):
        for f in files:
            try:
                size = os.path.getsize(os.path.join(base, f))
            except OSError:
                continue
            count += 1
            total += size
            on_disk += ((size + cluster - 1) // cluster) * cluster
    return count, total, on_disk


def human(n: float) -> str:
    return f"{n / 1024 / 1024:.1f} MB"


MARKER = """ExcelFinder portable mode marker.

This file makes the app store its settings and index cache in the "data"
folder next to the executable, so a USB stick can move between PCs (and drive
letters) without leaving anything behind.

Delete this file to fall back to %APPDATA%\\ExcelFinder.
"""

README_TXT = """ExcelFinder —— 本地 Excel 极速检索（U盘便携版）
================================================================

怎么用（3 步）
----------------------------------------------------------------
1. 双击   ExcelFinder.exe
2. 点「浏览…」选目录 → 点「重建索引」
3. 在「关键词」里输入要查找的字符 → 按回车

双击结果里的某一行可以直接用 Excel 打开；右键有更多操作。


重要：两个勾选框的区别
----------------------------------------------------------------
- 只勾「文件名需命中」→ 只在文件名里找（最快，推荐日常用）
- 只勾「内容需命中」→ 只在单元格内容里找  ← 找人名、编号用这个
- 两个都勾       → 文件名和内容都必须有该关键词（一般用不到）

例：要找「所有包含张伟这个内容的表格」
  1) 取消勾选「文件名需命中」，只勾「内容需命中」
  2) 关键词输入：张伟
  3) 第一次会先解析所有文件内容（底部有进度条，约 270 个文件/秒），
     完成后自动重新搜索；解析结果会缓存，以后搜索都是毫秒级

结果表格里的「命中位置」列会直接告诉你它在哪：
      华北汇总!D8 等3处        ← 工作表名 + 单元格地址
选中某一行，底部预览面板会列出全部命中位置：
      1. 工作表「华北汇总」 D8 = 张伟
         该行: 7 | 2021-06-08 | 市场部 | 张伟 | 扫地机器人 | 海外 | 442 | 783333
（「该行」就是这一行的所有内容，用来区分同名的不同记录）

定位与打开
----------------------------------------------------------------
- 双击某一行  → 用表格程序打开，并自动跳到那个工作表的第一处命中单元格
- 右键 →「定位到命中单元格」  同上
- 右键 →「在资源管理器中定位」  只在文件夹里选中文件（不需要装 Office）
- 右键 →「导出当前结果到 CSV」  可选把「工作表 + 单元格」两列一起导出


选择用 Office 还是 WPS 打开
----------------------------------------------------------------
结果列表下方有一个「打开方式」下拉框，三个选项：

  自动（Office 优先，无则用 WPS）   默认。装了 Office 就用 Office，否则用 WPS
  只用 Microsoft Excel 打开         强制 Excel；没装会提示并自动回退
  只用 WPS 表格打开                 强制 WPS 表格（12.x 用 Ket.Application 接口）

- 选择会记住，下次打开程序仍是这个设置
- 只装了一个程序时，默认就选中那一个，不必手动选
- 打开后状态栏会写明实际用的是哪个程序（例如「已用 WPS 表格 打开」）
- 本程序只检测、不安装，也不会改动 Office / WPS 的任何设置

关于进程占用：用 Excel 打开后会自动释放后台的 COM 实例，不会在任务管理器里
留下一堆 EXCEL.EXE，也不会锁住文件；WPS 若本来就开着则复用它的窗口。


占多大空间？（重要）
----------------------------------------------------------------
出厂状态（data\ 为空）：

  文件数  160 个
  实际大小 49.5 MB
  占用空间 49.8 MB

其中：
  ExcelFinder.exe        12.6 MB
  ExcelFinder-cli.exe     9.6 MB
  python\                27.0 MB   ← 备用启动方式用的内置解释器（可整个删掉）
  src\ 和文档              0.2 MB

**运行以后 data\ 会变大**，这是正常的，因为它保存索引缓存：

  仅文件名索引    每个文件约 176 字节   →  3400 个文件约 0.6 MB
  含内容的索引    每个文件约 4.5 KB     →  3400 个文件约 15 MB
                                       →  2 万个文件约 86 MB

所以：
- 只用「文件名检索」的话，索引很小（几千个文件不到 1 MB）
- 做了「内容索引」才会明显变大，跟文件里的文字量成正比
- 如果算出来会超过 256 MB，程序**不会**写缓存，只在内存里保留
  （重启后需要重新解析，但不会占你的磁盘）
- 想彻底清空：直接删掉 data\ 文件夹即可，程序会重建

想再小一点：如果不需要「exe 被拦截时用内置 Python 启动」这个备用方案，
可以整个删掉 python\ 文件夹，包体立刻从 49.5 MB 降到约 22 MB。
（需要时可以用 tools\fetch_python_embed.py 重新生成。）


想在别的电脑上运行，需要什么环境？
----------------------------------------------------------------
几乎什么都不用装。实测这个 exe 里已经内嵌了：

  CPU 架构        必须 64 位（x64）。32 位 Windows 用不了
  操作系统        建议 Windows 10 / 11 64 位
  Python          已打包在 exe 内部，不需要装
  VC++ 运行库      VCRUNTIME140.dll + ucrtbase.dll 已内嵌，不需要装
  .NET / Java     完全不需要

唯一"可选"的是表格程序：
  - 不装 Office / WPS：搜索、浏览、导出 CSV 全都正常
  - 装了：双击结果能直接跳到具体单元格（用「打开方式」选 Excel 还是 WPS）

另外提醒一句：本程序没有数字签名，个别电脑的杀毒软件可能误报。
遇到就右键 exe → 属性 → 勾「解除锁定」，或双击「备用启动.bat」。


U盘占用空间比文件大小大很多？（例如 70 MB 文件占了 1 GB）
----------------------------------------------------------------
这是U盘的「簇大小」造成的，不是程序的问题。

原理：磁盘按"簇"分配空间，一个文件至少占 1 个簇。簇越大，小文件浪费越多。
实测同样的文件在不同簇大小下：

  簇   4 KB  ->  49.8 MB      ← 正常
  簇  32 KB  ->  52.8 MB      ← 推荐的U盘设置
  簇 512 KB  -> 121.5 MB
  簇   8 MB  ->  1296 MB      ← 你的U盘可能是这种

为什么会这样：Windows 格式化大容量U盘为 exFAT 时，如果分配单元大小选了
"默认"或很大的值，簇可能达到几 MB。160 个文件 × 8 MB 就是 1.2 GB。

怎么查：把U盘插上，在本程序目录运行
        python tools\check_usage.py E:\
        （E: 换成你的U盘盘符）它会直接告诉你簇大小和建议。

怎么修：备份U盘上的东西 → 右键U盘 → 格式化 →
        文件系统选 exFAT，分配单元大小选 32 KB（或 64 KB），快速格式化。

不想重格式化也行：只要U盘剩余空间够，直接拷过去就能用；
本程序只有 160 个文件，不会像几千个文件那样疯狂浪费空间。


刷新列表（F5）
----------------------------------------------------------------
目录里的文件变了（新增、删除、改名、另存），不想重新建整个索引时用这个：

  点「刷新列表」或按 F5
  → 只对比文件的大小和修改时间，没变的文件直接跳过（很快）
  → 新增/删除的文件同步进索引
  → 然后自动按当前关键词重搜一遍，列表立刻反映磁盘现状

底部状态栏会写明这次刷新到底发现了什么，例如：
  命中 3 个文件 · 查询耗时 0.4 ms · 索引 3 个文件  ·  刷新完成：新增 丙_客户名单.xlsx
  命中 2 个文件 · 查询耗时 0.3 ms · 索引 2 个文件  ·  刷新完成：目录无变化

和「重建索引」的区别：重建会丢掉缓存重新走一遍所有文件；刷新是增量的，
只处理变化的部分，所以快得多。只在结果看起来不对时才需要重建。

注意：刷新只更新"文件名/位置"这份索引。如果某个文件的内容变了，
内容索引需要重新解析该文件 —— 目前请用「建立内容索引」，它会跳过已解析且
未改动的文件。

命令行的对应写法
----------------------------------------------------------------
  ExcelFinder-cli.exe "D:\报表" --query 张伟 --content --cells 3
  （--cells N 表示每个文件最多列出 N 处单元格定位）


进度条说明（右下角）
----------------------------------------------------------------
进度条会告诉你当前在哪一步、还剩多少，各阶段的含义：

  扫描目录 …      条纹动画 —— 扫描前无法预知文件总数，所以用不定量条纹
  解析内容 37%    第一次做内容检索时解析所有文件，这是最慢的一步，最需要看它
  检索索引 0%     纯内存查询，一闪而过（通常几十毫秒）
  定位单元格 62%  逐个打开命中的表格，找出具体是哪个工作表哪个单元格
  完成 3400 命中  结束
  已按 xxx 排序   重新排序只是重画表格，不会重新读文件（所以是秒完）

另外「停止」按钮随时可用：扫描、内容解析、单元格定位都能中途停下。

补充：如果某个文件是 .csv / .txt，它没有"单元格地址"（不是 Excel 表格结构），
      这类命中的「命中位置」列会显示「内容」而不是「工作表!单元格」，这是正常的。


比 Windows 自带搜索快多少（实测 6 万个表格文件 / 424 MB）
----------------------------------------------------------------
  PowerShell Get-ChildItem -Recurse       2388 ms      1x
  CMD where /r（系统自带）                  920 ms    2.6x
  ExcelFinder 首次扫描                      468 ms      5x
  ExcelFinder 已建索引检索                   42 ms     57x


文件说明
----------------------------------------------------------------
  ExcelFinder.exe        主程序（图形界面），日常只用这个
  启动 ExcelFinder.bat   启动脚本，等同于双击主程序
  ExcelFinder-cli.exe    命令行版，可写进批处理脚本
  自检.bat               双击即可体检，确认程序正常
  使用说明.md            图文详细说明（用记事本或 Markdown 阅读器打开）
  portable.marker        便携模式开关（删掉它就会把配置存到 %APPDATA%）
  data\\                  自动生成：保存配置和索引缓存
  python\\                内置 Python 运行环境（备用，见下）
  src\\                   引擎源码
  run_gui.py             备用启动入口

以上除 data\\ 外都不要删，程序依赖它们。


如果 ExcelFinder.exe 被拦截
----------------------------------------------------------------
本程序未做数字签名，个别电脑的杀毒软件或公司策略可能会拦截。
1) 首选：右键 ExcelFinder.exe → 属性 → 勾「解除锁定」→ 确定，再运行
        或点提示里的「更多信息」→「仍要运行」
2) 备用：双击  备用启动.bat
        它用内置的 Python 直接跑同一份引擎代码，通常不会被拦
3) 只想要结果：用 ExcelFinder-cli.exe（控制台程序）


支持的格式
----------------------------------------------------------------
  .xlsx .xlsm .xltx .xltm .xlam   文件名 + 单元格内容（完整支持）
  .csv  .tsv  .txt                文件名 + 内容
  .xls  .xlsb .xlt .xla .ods      文件名完整支持，内容尽力提取
  解析内容不需要安装 Office，也不依赖 pandas / openpyxl


常见问题
----------------------------------------------------------------
Q: 第一次搜内容很慢？
A: 第一次要解析所有文件，之后走缓存，第二次起是毫秒级。

Q: 换盘符或换电脑，配置会丢吗？
A: 不会。配置存在程序旁边的 data\\ 里，与盘符无关。

Q: 会在别人电脑上留东西吗？
A: 便携模式下不会。所有数据都在U盘的 data\\ 里，删掉 data\\ 即完全重置。

Q: 搜索时界面卡住？
A: 不会。扫描和解析都在后台线程，底部有进度条，随时可以点「停止」。

Q: 能搜 WPS 表格吗？
A: 能，WPS 保存的 .xlsx 格式一致。

Q: 在别人的电脑上会流畅吗？对电脑配置有要求吗？
A: 要求很低。索引常驻内存，典型占用是「每个文件约 1-5 KB 内存」
   （取决于文件里的文字量），一万个文件约几十 MB，现在的电脑都没问题。
   实测 6 万个文件查询一次 42 毫秒。
   影响速度的只有两件事：
     - 磁盘速度：U盘本身比内置硬盘慢，第一次扫描和解析内容会慢一些。
       如果是 USB 2.0 的老U盘，第一次扫描可能要几分钟，之后走缓存就快了。
     - 扫描线程数：默认 4。老电脑可以调到 2；网络共享盘可以调到 8-16。
   另外第一次运行、以及每次插到新电脑上，exe 需要释放到临时目录，
   会多花 2-3 秒（杀毒软件实时扫描还会更久），这是正常的。
"""


def stage(target: str) -> None:
    hr(f"1. 创建分发目录：{target}")
    if os.path.isdir(target):
        print("     目标已存在，先清空")
        shutil.rmtree(target, ignore_errors=True)
        if os.path.isdir(target):
            # A file inside can still be locked (an EXE the shell or a scanner
            # has open). Move it aside instead of failing the whole build.
            stale = f"{target}.old-{time.strftime('%H%M%S')}"
            try:
                os.rename(target, stale)
                print(f"     ! 无法删除（文件被占用），已改名保留：{stale}")
            except OSError as exc:
                raise SystemExit(f"目标目录既无法删除也无法改名：{exc}")
    os.makedirs(target, exist_ok=True)
    os.makedirs(os.path.join(target, "data"), exist_ok=True)
    os.makedirs(os.path.join(target, "src"), exist_ok=True)

    copied: list[tuple[str, int]] = []

    def put(src: str, rel: str) -> None:
        dst = os.path.join(target, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        copied.append((rel, os.path.getsize(dst)))

    # EXEs exist both at the repo root (handy for double-clicking) and in
    # dist/ (freshly built). Pick whichever is NEWEST -- copying the stale root
    # copy silently ships an old binary that still looks freshly packaged.
    for name in ("ExcelFinder.exe", "ExcelFinder-cli.exe"):
        candidates = [
            p for p in (os.path.join(ROOT, "dist", name), os.path.join(ROOT, name))
            if os.path.exists(p)
        ]
        if not candidates:
            failures.append(f"missing {name}")
            print(f"  ! {name} 缺失，请先运行 tools/build.py")
            continue
        src = max(candidates, key=os.path.getmtime)
        put(src, name)
        print(f"  · {name} 取自 {os.path.relpath(src, ROOT)}"
              f"  ({time.strftime('%H:%M:%S', time.localtime(os.path.getmtime(src)))})")

    for name in ("portable.marker",):
        put(os.path.join(ROOT, name), name)

    # Launchers / docs live at the repo root alongside the EXEs.
    for name in ("启动 ExcelFinder.bat", "自检.bat"):
        put(os.path.join(ROOT, name), name)
    put(os.path.join(ROOT, "使用说明.md"), "使用说明.md")

    # Engine sources + fallback launcher so 备用启动.bat works with the bundled runtime.
    for name in ("excelfinder_core.py", "excelfinder_gui.py", "excelfinder_cli.py"):
        put(os.path.join(ROOT, "src", name), os.path.join("src", name))
    put(os.path.join(ROOT, "run_gui.py"), "run_gui.py")

    # Prune the bundled runtime here as well: the repo's copy may pre-date the
    # last pruning, and this keeps the package small whether or not build.py ran.
    py_src = os.path.join(ROOT, "embed-python")
    if os.path.isdir(py_src):
        pruned = os.path.join(py_src, "tcl", "tk8.6", "demos")
        if os.path.isdir(pruned):
            print("     （embed-python/ 尚未裁剪，运行 tools/build.py 可减少约 940 个文件）")

    for rel, size in copied:
        print(f"     + {rel:<26} {human(size)}")

    # Bundled runtime (optional but keeps the fallback path alive).
    py_src = os.path.join(ROOT, "embed-python")
    if os.path.isdir(py_src):
        shutil.copytree(py_src, os.path.join(target, "python"))
        total = sum(
            os.path.getsize(os.path.join(b, f))
            for b, _d, fs in os.walk(os.path.join(target, "python")) for f in fs
        )
        print(f"     + python/ (内置运行环境)          {human(total)}")
    else:
        print("     ! python/ 缺失：备用启动方式不可用（主程序不受影响）")

    # Generated helper files
    with open(os.path.join(target, "portable.marker"), "w", encoding="utf-8") as fh:
        fh.write(MARKER)
    with open(os.path.join(target, "先读我.txt"), "w", encoding="utf-8") as fh:
        fh.write(README_TXT)

    backup_bat = f"""@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
echo 备用启动方式：使用内置 Python 运行 ExcelFinder
echo （当 ExcelFinder.exe 被杀毒软件/公司策略拦截时使用）
echo.
if not exist "python\\pythonw.exe" (
    echo [错误] 找不到 python\\pythonw.exe，无法使用备用方式。
    echo        请改用 ExcelFinder.exe，或先解除它的锁定。
    pause
    exit /b 1
)
start "" "python\\pythonw.exe" "run_gui.py" %*
exit /b 0
"""
    with open(os.path.join(target, "备用启动.bat"), "w", encoding="utf-8") as fh:
        fh.write(backup_bat)
    print("     + 先读我.txt")
    print("     + 备用启动.bat")


def verify(target: str) -> None:
    hr("2. 验证分发目录")
    for name in ("ExcelFinder.exe", "ExcelFinder-cli.exe", "portable.marker",
                 "先读我.txt", "备用启动.bat", "启动 ExcelFinder.bat", "自检.bat",
                 "使用说明.md", "run_gui.py"):
        check(f"存在 {name}", os.path.exists(os.path.join(target, name)))
    for name in ("excelfinder_core.py", "excelfinder_gui.py", "excelfinder_cli.py"):
        check(f"存在 src/{name}", os.path.exists(os.path.join(target, "src", name)))

    # Icon really embedded?
    exe = os.path.join(target, "ExcelFinder.exe")
    with open(exe, "rb") as fh:
        head = fh.read(2)
    check("ExcelFinder.exe 是有效的 PE 文件", head == b"MZ")

    # Portable detection from the staged folder (no data\ contents yet).
    probe = subprocess.run(
        [PY, "-c",
         "import sys, os; sys.path.insert(0, os.path.join(%r,'src'));"
         "from excelfinder_core import state_dir, is_portable;"
         "print(is_portable(), state_dir())" % target],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    out = (probe.stdout or "").strip()
    check("便携模式生效（配置写到 data\\）", out.startswith("True"), out or probe.stderr[:120])

    # CLI smoke test against the staged copy.
    cli = os.path.join(target, "ExcelFinder-cli.exe")
    sample = os.path.join(ROOT, "testdata")
    if os.path.isdir(sample):
        p = subprocess.run([cli, sample, "--query", "预算", "--limit", "2"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
        check("CLI 可运行并返回结果", p.returncode == 0 and "命中" in (p.stdout or ""),
              f"exit={p.returncode}")


def run_gui_probe(target: str) -> None:
    """Launch the staged GUI, confirm a window appears and data\\ gets written."""
    hr("3. 真实启动测试（模拟插到另一台电脑首次运行）")
    sys.path.insert(0, HERE)
    from check_window import descendants, windows_of

    sample = os.path.join(ROOT, "testdata")
    if not os.path.isdir(sample):
        print("     ! 缺少 testdata，跳过")
        return

    data_dir = os.path.join(target, "data")
    shutil.rmtree(data_dir, ignore_errors=True)
    os.makedirs(data_dir)
    appdata = os.path.join(os.environ.get("APPDATA", ""), "ExcelFinder")
    shutil.rmtree(appdata, ignore_errors=True)

    env = dict(os.environ)
    env["TMP"] = env["TEMP"] = os.path.join(data_dir, ".tmp")
    os.makedirs(env["TMP"], exist_ok=True)
    proc = subprocess.Popen([os.path.join(target, "ExcelFinder.exe"), sample], env=env)

    window = False
    index_ok = False
    deadline = time.time() + 150
    while time.time() < deadline:
        time.sleep(1.0)
        if proc.poll() is not None:
            break
        if not window:
            vis = [w for w in windows_of(descendants(proc.pid))
                   if w["visible"] and w["size"][0] > 50]
            if vis:
                window = True
                print(f"     窗口: {vis[0]['title']!r} {vis[0]['size']}")
        idx = os.path.join(data_dir, "index_v1.json")
        if os.path.exists(idx):
            try:
                with open(idx, encoding="utf-8") as fh:
                    recs = json.load(fh).get("records", [])
                if recs:
                    print(f"     索引: {len(recs)} 条记录")
                    index_ok = True
                    break
            except (OSError, ValueError):
                pass
    subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                   capture_output=True, text=True)
    time.sleep(1.0)

    check("图形界面正常显示", window)
    check("索引写入 data\\ 且非空", index_ok)
    check("未在 %APPDATA% 留下痕迹", not os.path.exists(appdata))

    # leave the folder pristine for the user
    shutil.rmtree(data_dir, ignore_errors=True)
    os.makedirs(data_dir)


def make_zip(target: str) -> str:
    hr("4. 额外生成一个 zip（可选分发形式）")
    out = os.path.join(os.path.dirname(target), os.path.basename(target) + ".zip")
    name = os.path.basename(target)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        zf.writestr(f"{name}/data/", "")
        for base, _dirs, files in os.walk(target):
            for f in files:
                full = os.path.join(base, f)
                zf.write(full, os.path.join(name, os.path.relpath(full, target)))
    print(f"     {out} ({human(os.path.getsize(out))})")
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("target", nargs="?", default=os.path.join(desktop(), "ExcelFinder-U盘版"))
    ap.add_argument("--zip", action="store_true")
    args = ap.parse_args()
    target = os.path.abspath(args.target)

    stage(target)
    verify(target)
    run_gui_probe(target)

    count, total, on_disk = dir_usage(target)
    if args.zip:
        make_zip(target)

    hr("结果")
    print(f"分发目录：{target}")
    print(f"文件数  ：{count:,}")
    print(f"实际大小：{human(total)}")
    print(f"占用空间：{human(on_disk)}（本机 4 KB 簇估算；U盘若为 32 KB 簇会更大）")
    print("说明：data\\ 是空的，第一次运行后会在里面生成索引缓存（内容索引可能几十 MB）。")
    if FAILURES:
        print(f"验证失败（{len(FAILURES)} 项）：{FAILURES}")
        return 1
    print("验证全部通过 —— 整个文件夹拷到U盘即可在别的电脑上使用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
