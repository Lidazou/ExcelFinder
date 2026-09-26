# -*- coding: utf-8 -*-
"""Generate a synthetic Excel corpus for benchmarking ExcelFinder.

Usage:  python make_testdata.py <target_dir> [--files N] [--hardlinks N] [--seed 1]

Creates a nested tree of .xlsx/.csv files with names and cell contents drawn
from realistic Chinese/English business vocabulary, plus a configurable number
of hard links (same bytes, new names) to measure raw path-enumeration speed
without inflating disk usage.
"""
from __future__ import annotations

import argparse
import os
import random
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsxwriter_lite import write_xlsx  # noqa: E402

DEPT = ["销售部", "财务部", "人力资源部", "研发部", "市场部", "客服部", "供应链部", "法务部"]
YEAR = ["2021", "2022", "2023", "2024", "2025"]
TOPIC = [
    "季度报表", "年度预算", "客户名单", "合同台账", "考勤统计", "费用报销",
    "工资明细", "库存盘点", "采购订单", "项目排期", "会议纪要", "培训计划",
    "销售漏斗", "回款计划", "发票明细", "差旅申请", "供应商评估", "产品路线图",
]
REGION = ["华东", "华北", "华南", "西南", "西北", "东北", "海外"]
PRODUCT = ["智能音箱", "无线耳机", "扫地机器人", "智能门锁", "空气净化器", "电动牙刷"]
STATUS = ["待审核", "已通过", "已驳回", "进行中", "已完成", "已归档"]
PERSON = ["张伟", "李娜", "王芳", "刘洋", "陈静", "杨帆", "赵磊", "黄敏", "周涛", "吴倩"]


def make_book(path: str, rng: random.Random, rows: int) -> None:
    headers = ["序号", "日期", "部门", "负责人", "产品", "区域", "数量", "金额", "状态", "备注"]
    body: list[list] = [headers]
    for i in range(1, rows + 1):
        body.append(
            [
                i,
                f"{rng.choice(YEAR)}-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d}",
                rng.choice(DEPT),
                rng.choice(PERSON),
                rng.choice(PRODUCT),
                rng.choice(REGION),
                rng.randint(1, 500),
                rng.randint(100, 900000),
                rng.choice(STATUS),
                rng.choice(TOPIC) + "见附件",
            ]
        )
    sheets = [(f"{rng.choice(REGION)}{rng.choice(['明细', '汇总', '附表'])}", body)]
    if rows > 40:
        sheets.append(("说明", [["文档密级", "内部公开"], ["关键词", rng.choice(TOPIC)]]))
    write_xlsx(path, sheets)


def make_csv(path: str, rng: random.Random, rows: int) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write("序号,日期,部门,负责人,产品,区域,金额,状态\n")
        for i in range(1, rows + 1):
            fh.write(
                f"{i},{rng.choice(YEAR)}-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d},"
                f"{rng.choice(DEPT)},{rng.choice(PERSON)},{rng.choice(PRODUCT)},"
                f"{rng.choice(REGION)},{rng.randint(100, 900000)},{rng.choice(STATUS)}\n"
            )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--files", type=int, default=400)
    ap.add_argument("--hardlinks", type=int, default=0)
    ap.add_argument("--rows", type=int, default=30)
    ap.add_argument("--seed", type=int, default=20240101)
    ap.add_argument("--clean", action="store_true")
    ap.add_argument(
        "--fast",
        action="store_true",
        help="build a small template pool and copy it (much faster for 10k+ files)",
    )
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    if args.clean and os.path.isdir(root):
        shutil.rmtree(root, ignore_errors=True)
    os.makedirs(root, exist_ok=True)

    rng = random.Random(args.seed)
    dirs: list[str] = [root]
    for year in YEAR:
        for dept in DEPT:
            d = os.path.join(root, f"{year}年", dept)
            os.makedirs(d, exist_ok=True)
            dirs.append(d)

    written = 0
    templates: list[str] = []
    if args.fast:
        pool_dir = os.path.join(root, "_templates")
        os.makedirs(pool_dir, exist_ok=True)
        for j in range(60):
            p = os.path.join(pool_dir, f"_t{j:02d}.xlsx")
            make_book(p, random.Random(args.seed + j), args.rows + (j % 7) * 5)
            templates.append(p)

    for i in range(args.files):
        rng2 = random.Random(args.seed + i)
        folder = rng2.choice(dirs[1:])
        name = (
            f"{rng2.choice(YEAR)}年{rng2.choice(REGION)}{rng2.choice(DEPT)}"
            f"{rng2.choice(TOPIC)}_{rng2.randint(1, 9999):04d}"
        )
        if rng2.random() < 0.25:
            name += "_" + rng2.choice(["终版", "修订", "草稿", "V2", "final"])
        rows = args.rows if rng2.random() < 0.9 else args.rows * 20
        ext = ".csv" if rng2.random() < 0.1 else ".xlsx"
        path = os.path.join(folder, name + ext)
        try:
            if templates and ext == ".xlsx":
                shutil.copyfile(templates[rng2.randrange(len(templates))], path)
            elif ext == ".csv":
                make_csv(path, rng2, rows)
            else:
                make_book(path, rng2, rows)
            written += 1
        except Exception as exc:
            print(f"  ! failed {path}: {exc}", file=sys.stderr)

    linked = 0
    if args.hardlinks:
        template_dir = os.path.join(root, "硬链接压力测试")
        os.makedirs(template_dir, exist_ok=True)
        templates = []
        for j in range(4):
            p = os.path.join(template_dir, f"_template_{j}.xlsx")
            make_book(p, random.Random(j), 200)
            templates.append(p)
        for k in range(args.hardlinks):
            src = templates[k % len(templates)]
            sub = os.path.join(template_dir, f"批次{k // 500:03d}")
            os.makedirs(sub, exist_ok=True)
            dst = os.path.join(sub, f"{random.Random(k).choice(YEAR)}{random.Random(k + 1).choice(TOPIC)}_{k:06d}.xlsx")
            try:
                os.link(src, dst)
                linked += 1
            except OSError:
                shutil.copyfile(src, dst)
                linked += 1

    total_bytes = 0
    count = 0
    for base, _dirs, files in os.walk(root):
        for f in files:
            if os.path.splitext(f)[1].lower() in (".xlsx", ".csv"):
                count += 1
                try:
                    total_bytes += os.path.getsize(os.path.join(base, f))
                except OSError:
                    pass
    print(f"created {written} files + {linked} hard links")
    print(f"tree now holds {count} spreadsheets, {total_bytes/1024/1024:.1f} MiB (hard links share bytes)")
    print(f"root: {root}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
