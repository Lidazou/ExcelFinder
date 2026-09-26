# -*- coding: utf-8 -*-
"""Generate the README's SVG diagrams.

Everything is drawn from code so the images stay in sync with the project and
can be regenerated after a change. SVGs are used instead of PNGs because GitHub
renders them crisply at any zoom and they stay tiny in the repo.

Usage: python tools/make_diagrams.py [output_dir]
"""
from __future__ import annotations

import os
import sys

# Palette (kept in one place so every diagram matches)
INK = "#1f2933"
MUTED = "#5c6b7a"
FAINT = "#8899a6"
GREEN = "#1f6f3c"
GREEN_L = "#e8f3ec"
GREEN_M = "#2e8b57"
BLUE = "#2a7fd4"
BLUE_L = "#eaf3fd"
AMBER = "#e08a1a"
AMBER_L = "#fdf3e3"
RED = "#c0392b"
LINE = "#d5dde5"
CARD = "#ffffff"
BG = "#f7f9fc"

FONT = ("-apple-system,BlinkMacSystemFont,'Segoe UI','Microsoft YaHei UI',"
        "'Microsoft YaHei',Roboto,'Helvetica Neue',Arial,sans-serif")
MONO = "'JetBrains Mono','Cascadia Mono',Consolas,'Courier New',monospace"


def esc(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def svg(width: int, height: int, body: str, title: str) -> str:
    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}"
     viewBox="0 0 {width} {height}" role="img" aria-label="{esc(title)}">
  <title>{esc(title)}</title>
  <defs>
    <linearGradient id="ghero" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#1f6f3c"/>
      <stop offset="100%" stop-color="#2e8b57"/>
    </linearGradient>
    <linearGradient id="gbar" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="#2a7fd4"/>
      <stop offset="100%" stop-color="#1f6f3c"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="2" stdDeviation="3" flood-color="#1f2933" flood-opacity="0.10"/>
    </filter>
  </defs>
{body}
</svg>
"""


def text(x, y, s, size=14, fill=INK, weight="400", anchor="start", family=None,
         spacing=None, opacity=None):
    attrs = [f'x="{x}"', f'y="{y}"', f'font-size="{size}"', f'fill="{fill}"']
    if weight != "400":
        attrs.append(f'font-weight="{weight}"')
    if anchor != "start":
        attrs.append(f'text-anchor="{anchor}"')
    attrs.append(f'font-family="{family or FONT}"')
    if spacing:
        attrs.append(f'letter-spacing="{spacing}"')
    if opacity:
        attrs.append(f'opacity="{opacity}"')
    return f'  <text {" ".join(attrs)}>{esc(s)}</text>'


def rect(x, y, w, h, rx=10, fill=CARD, stroke=LINE, sw=1, shadow=False, extra=""):
    sh = ' filter="url(#shadow)"' if shadow else ""
    st = f' stroke="{stroke}" stroke-width="{sw}"' if stroke else ""
    return (f'  <rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" '
            f'fill="{fill}"{st}{sh}{extra}/>')


# --------------------------------------------------------------------------
# 1. Hero banner
# --------------------------------------------------------------------------

def hero(path: str) -> None:
    W, H = 1000, 260
    b = [rect(0, 0, W, H, rx=0, fill="url(#ghero)", stroke=None)]
    # decorative grid of "spreadsheet" cells
    for i in range(9):
        for j in range(5):
            b.append(f'  <rect x="{700 + i*34}" y="{34 + j*34}" width="30" height="30" rx="4" '
                     f'fill="#ffffff" opacity="0.06"/>')
    # magnifier
    b.append('  <circle cx="800" cy="140" r="54" fill="#ffffff" opacity="0.10"/>')
    b.append('  <circle cx="800" cy="140" r="54" fill="none" stroke="#ffffff" '
             'stroke-width="9" opacity="0.85"/>')
    b.append('  <circle cx="800" cy="140" r="42" fill="#ffffff" opacity="0.16"/>')
    b.append('  <line x1="840" y1="180" x2="884" y2="224" stroke="#e08a1a" '
             'stroke-width="15" stroke-linecap="round"/>')

    b.append(text(56, 84, "ExcelFinder", 46, "#ffffff", "700", spacing="0.5"))
    b.append(text(56, 124, "本地 Excel 极速检索 · 精确到单元格", 21, "#e8f3ec"))
    b.append(rect(56, 150, 470, 40, rx=20, fill="#ffffff", stroke=None, extra=' opacity="0.14"'))
    b.append(text(76, 177, "☑ 免安装   ☑ 零依赖   ☑ U盘便携   ☑ 支持 WPS", 15, "#ffffff"))
    b.append(text(56, 224, "索引一次 · 查询 40 毫秒 · 找出「哪个工作表的哪个单元格」",
                  14, "#cfe8da"))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(svg(W, H, "\n".join(b), "ExcelFinder 项目横幅"))


# --------------------------------------------------------------------------
# 2. Architecture / data flow
# --------------------------------------------------------------------------

def architecture(path: str) -> None:
    W, H = 1000, 470
    b = [rect(0, 0, W, H, rx=0, fill=BG, stroke=None)]

    def stage(x, y, w, h, title, items, color, light):
        b.append(rect(x, y, w, h, rx=12, fill=CARD, stroke=color, sw=2, shadow=True))
        b.append(rect(x, y, w, 34, rx=12, fill=color, stroke=None))
        b.append(rect(x, y + 22, w, 12, rx=0, fill=color, stroke=None))
        b.append(text(x + w / 2, y + 23, title, 15, "#ffffff", "600", anchor="middle"))
        for i, (name, note) in enumerate(items):
            ty = y + 58 + i * 40
            b.append(text(x + 16, ty, name, 14, INK, "600"))
            if note:
                b.append(text(x + 16, ty + 17, note, 12, MUTED))
        b.append(f'  <circle cx="{x + w - 22}" cy="{y + 17}" r="11" fill="{light}"/>')

    stage(24, 26, 268, 190, "① 建立索引（一次性）", [
        ("多线程扫描目录", "2–4 线程最优，元数据密集"),
        ("筛选表格文件", "xlsx/xlsm/xls/xlsb/csv…"),
        ("缓存 (路径,大小,时间)", "未改动文件直接复用"),
    ], GREEN, GREEN_L)

    stage(366, 26, 268, 190, "② 解析内容（可选）", [
        ("直接解 OOXML 内部 XML", "不依赖 Office / pandas"),
        ("按列去重 + 截断", "重复字符串只记一次"),
        ("约 270 文件/秒", "首次约 12 秒 / 3400 文件"),
    ], BLUE, BLUE_L)

    stage(708, 26, 268, 190, "③ 检索（毫秒级）", [
        ("纯内存运算", "不碰磁盘"),
        ("空格 = AND", "2023 销售部"),
        ("通配符 / 正则 / 前缀", "字符前传等场景"),
    ], AMBER, AMBER_L)

    # arrows between stages
    for x0, x1 in ((292, 366), (634, 708)):
        b.append(f'  <path d="M{x0} 121 L{x1 - 12} 121" stroke="{FAINT}" stroke-width="2.5" '
                 f'stroke-dasharray="6 5"/>')
        b.append(f'  <path d="M{x1 - 14} 114 L{x1 - 2} 121 L{x1 - 14} 128 Z" fill="{FAINT}"/>')

    # locate step under stage 3
    b.append(rect(708, 246, 268, 96, rx=12, fill=GREEN_L, stroke=GREEN, sw=2))
    b.append(text(724, 274, "④ 定位到单元格", 15, GREEN, "700"))
    b.append(text(724, 298, "对命中的文件做第二遍扫描，", 12.5, MUTED))
    b.append(text(724, 317, "得到「工作表!单元格」+ 整行内容", 12.5, MUTED))
    b.append(f'  <path d="M842 216 L842 246" stroke="{GREEN}" stroke-width="2.5"/>')
    b.append(f'  <path d="M835 240 L842 252 L849 240 Z" fill="{GREEN}"/>')

    # index cache box
    b.append(rect(24, 246, 610, 96, rx=12, fill=CARD, stroke=LINE, sw=1.5))
    b.append(text(44, 274, "磁盘缓存 data\\index_v1.json", 14.5, INK, "600"))
    b.append(text(44, 298, "仅文件名约 176 字节/文件 · 含内容约 4.5 KB/文件", 12.5, MUTED))
    b.append(text(44, 318, "超过 256 MB 不落盘，只在内存中保留", 12.5, AMBER))

    b.append(text(24, 384, "关键取舍：索引阶段只回答「哪个文件命中」，坐标按需二次解析 ——",
                  13.5, INK, "600"))
    b.append(text(24, 408, "若在索引里预存全部单元格坐标，内存要翻好几倍；实测第二遍扫描仅 3.3 ms/文件。",
                  13, MUTED))
    b.append(text(24, 442, "全程只依赖 Python 标准库：单文件 exe 12.6 MB，目标机器无需任何运行库。",
                  13, GREEN, "600"))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(svg(W, H, "\n".join(b), "ExcelFinder 工作原理"))


# --------------------------------------------------------------------------
# 3. Performance chart
# --------------------------------------------------------------------------

def performance(path: str) -> None:
    W, H = 1000, 430
    bars = [
        ("PowerShell Get-ChildItem -Recurse", 2388, MUTED, "1×"),
        ("CMD where /r（系统自带）", 919, MUTED, "2.6×"),
        ("ExcelFinder 首次扫描目录", 468, BLUE, "5.1×"),
        ("ExcelFinder 已建索引检索", 42, GREEN, "57×"),
    ]
    b = [rect(0, 0, W, H, rx=0, fill=BG, stroke=None)]
    b.append(text(28, 44, "实测性能对比", 22, INK, "700"))
    b.append(text(28, 70, "6 万个表格文件 / 424 MB / 同一台电脑 / 关键词「预算」", 13.5, MUTED))

    x0, xmax = 330, 880
    top, rowh = 104, 68
    scale = (xmax - x0) / 2388.0
    for i, (label, ms, color, mult) in enumerate(bars):
        y = top + i * rowh
        b.append(rect(28, y, 950, 54, rx=9, fill=CARD, stroke=LINE, sw=1))
        b.append(text(46, y + 24, label, 14, INK, "600"))
        b.append(text(46, y + 42, f"{ms:,} ms", 12.5, MUTED, family=MONO))
        w = max(6, ms * scale)
        b.append(rect(x0, y + 12, w, 30, rx=6, fill=color, stroke=None))
        tx = x0 + w + 12
        if tx > 940:
            tx = x0 + w - 12
            b.append(text(tx, y + 32, mult, 15, "#ffffff", "700", anchor="end"))
        else:
            b.append(text(tx, y + 32, mult, 15, color, "700"))

    b.append(text(28, 396, "内容检索差距更大：Windows 需装 Office 并逐个调用 IFilter，"
                           "本项目直接解析 XML，约 270 文件/秒且结果缓存。",
                  12.5, MUTED))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(svg(W, H, "\n".join(b), "性能对比"))


# --------------------------------------------------------------------------
# 4. UI feature map
# --------------------------------------------------------------------------

def features(path: str) -> None:
    W, H = 1000, 560
    b = [rect(0, 0, W, H, rx=0, fill=BG, stroke=None)]
    b.append(text(28, 44, "界面与功能一览", 22, INK, "700"))
    b.append(text(28, 70, "仅示意图，说明各区域的作用", 13.5, MUTED))

    # window frame
    fx, fy, fw, fh = 28, 92, 944, 430
    b.append(rect(fx, fy, fw, fh, rx=10, fill=CARD, stroke=LINE, sw=1.5, shadow=True))
    b.append(rect(fx, fy, fw, 30, rx=10, fill="#e9eef3", stroke=None))
    b.append(rect(fx, fy + 20, fw, 10, rx=0, fill="#e9eef3", stroke=None))
    b.append(text(fx + 14, fy + 20, "ExcelFinder v1.0.0 — 本地 Excel 极速检索  [U盘便携版]",
                  12, MUTED, family=MONO))
    b.append(f'  <circle cx="{fx + fw - 20}" cy="{fy + 15}" r="5" fill="#c8d2dc"/>')

    # section 1
    b.append(rect(fx + 14, fy + 42, fw - 28, 66, rx=8, fill=GREEN_L, stroke=GREEN, sw=1.2))
    b.append(text(fx + 26, fy + 62, "① 目标目录与索引", 13, GREEN, "700"))
    b.append(text(fx + 26, fy + 84, "选择目录 · 重建索引 · 增量更新 · 扫描线程数 · 跳过隐藏文件",
                  12, MUTED))
    b.append(text(fx + 26, fy + 101, "索引统计：文件数 / 数据量 / 扫描耗时 / 内容索引数", 11.5, FAINT))

    # section 2
    b.append(rect(fx + 14, fy + 116, fw - 28, 78, rx=8, fill=BLUE_L, stroke=BLUE, sw=1.2))
    b.append(text(fx + 26, fy + 136, "② 关键词检索", 13, BLUE, "700"))
    b.append(text(fx + 26, fy + 158, "关键词 · 搜索 · 清空 · 刷新列表(F5) · 建立内容索引 · 导出 CSV",
                  12, MUTED))
    b.append(text(fx + 26, fy + 178,
                  "文件名需命中 / 内容需命中 · 匹配方式（包含·前缀·全等·通配符·正则）· 格式筛选",
                  12, MUTED))

    # results table
    b.append(rect(fx + 14, fy + 202, fw - 28, 118, rx=8, fill="#ffffff", stroke=LINE, sw=1.2))
    cols = [("文件名", 0.26), ("所在目录", 0.18), ("大小", 0.09),
            ("修改时间", 0.13), ("命中位置", 0.16), ("命中内容片段", 0.18)]
    cx = fx + 26
    for name, _frac in cols:
        b.append(text(cx, fy + 222, name, 11, MUTED, "600"))
        cx += (fw - 52) * _frac
    b.append(f'  <line x1="{fx + 20}" y1="{fy + 230}" x2="{fx + fw - 20}" y2="{fy + 230}" '
             f'stroke="{LINE}" stroke-width="1"/>')
    sample = [("2021产品路线图_000055.xlsx", "…\\批次000", "12 KB", "2026-09-21 20:31", "西北汇总!D16 等3处"),
              ("2023年华东销售部年度预算_740.xlsx", "…\\财务部", "18 KB", "2026-09-20 09:12", "文件名"),
              ("2022年东北客户名单_1208.csv", "…\\市场部", "9 KB", "2026-09-19 15:44", "内容")]
    for r, row in enumerate(sample):
        ry = fy + 252 + r * 22
        cx = fx + 26
        for (name, _f), val in zip(cols, row):
            col = GREEN if val.startswith(("西北", "内容")) else INK
            b.append(text(cx, ry, val, 11, col))
            cx += (fw - 52) * _f

    # preview
    b.append(rect(fx + 14, fy + 328, fw - 28, 40, rx=8, fill=AMBER_L, stroke=AMBER, sw=1.2))
    b.append(text(fx + 26, fy + 345, "预览面板（可折叠）", 12, AMBER, "700"))
    b.append(text(fx + 26, fy + 361,
                  "1. 工作表「西北汇总」 D16 = 张伟    该行: 15 | 2023-09-10 | 客服部 | 张伟 | …",
                  11.5, MUTED))

    # bottom toolbar
    fy2 = fy + 378
    b.append(rect(fx + 14, fy2, fw - 28, 24, rx=6, fill="#eef3f7", stroke=LINE, sw=1))
    b.append(text(fx + 26, fy2 + 16, "☑ 显示预览面板", 10.5, INK))
    b.append(text(fx + 130, fy2 + 16, "打开方式:", 10.5, MUTED))
    b.append(rect(fx + 182, fy2 + 3, 158, 18, rx=4, fill="#ffffff", stroke=BLUE, sw=1))
    b.append(text(fx + 190, fy2 + 16, "自动（Office 优先）", 10.5, BLUE))
    b.append(text(fx + 352, fy2 + 16, "双击＝打开并跳到命中单元格", 10.5, FAINT))
    b.append(rect(fx + 620, fy2 + 4, 130, 16, rx=8, fill="#dde5ec", stroke=None))
    b.append(rect(fx + 620, fy2 + 4, 96, 16, rx=8, fill="url(#gbar)", stroke=None))
    b.append(text(fx + 758, fy2 + 16, "解析内容 74%", 10.5, GREEN, "600"))

    # status line
    b.append(text(fx + 26, fy2 + 44,
                  "命中 223 个文件 · 查询耗时 2.1 ms · 索引 3,404 个文件  ·  刷新完成：新增 1 个",
                  10.5, MUTED))
    b.append(text(fx + fw - 26, fy2 + 44, "U盘便携模式（配置存于程序目录）", 10.5, FAINT,
                  anchor="end"))

    b.append(text(28, 544, "「命中位置」列直接给出 工作表!单元格；双击即跳转，"
                           "不需在一千行里自己找。", 12.5, MUTED))
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(svg(W, H, "\n".join(b), "界面与功能一览"))


def main() -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    out = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(root, "docs", "images")
    os.makedirs(out, exist_ok=True)
    made = []
    for name, fn in (("hero.svg", hero), ("architecture.svg", architecture),
                     ("performance.svg", performance), ("features.svg", features)):
        path = os.path.join(out, name)
        fn(path)
        made.append((name, os.path.getsize(path)))
    for name, size in made:
        print(f"  + docs/images/{name}  ({size:,} bytes)")
    print(f"\n共 {len(made)} 张示意图 -> {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
