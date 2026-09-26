# -*- coding: utf-8 -*-
"""Verify a published release: metadata, asset, and a real download.

Downloads the asset and inspects the zip, because an asset that lists in the API
but fails to download (or arrives truncated) is the failure that actually hurts.

Usage:  set GITHUB_TOKEN=...  (optional, needed only for private repos)
        python tools/verify_release.py [--tag v1.0.0]
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
import urllib.error
import urllib.request
import zipfile

OWNER = os.environ.get("GITHUB_OWNER", "Lidazou")
REPO = os.environ.get("GITHUB_REPO", "ExcelFinder")

EXPECTED_IN_ZIP = [
    "ExcelFinder.exe", "ExcelFinder-cli.exe",
    "先读我.txt", "使用说明.md", "portable.marker",
    "python/python.exe", "src/excelfinder_core.py", "src/excelfinder_gui.py",
]
MARKERS = {
    "src/excelfinder_core.py": ("Ket.Application", "OPEN_WITH_CHOICES"),
    "src/excelfinder_gui.py": ("_refresh_list", "open_with_combo", "_progress_start"),
}


def get(url: str, token: str | None = None, raw: bool = False, binary: bool = False):
    headers = {"User-Agent": "excelfinder-verify-release"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=900) as resp:
        body = resp.read()
        if binary:
            return resp.status, body
        text = body.decode("utf-8", "replace")
        return resp.status, (text if raw else json.loads(text))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="v1.0.0")
    args = ap.parse_args()
    token = os.environ.get("GITHUB_TOKEN") or None
    failures: list[str] = []

    def check(label: str, ok: bool, detail: str = ""):
        print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
        if not ok:
            failures.append(label)

    print(f"检查 Release {OWNER}/{REPO} @ {args.tag}\n")

    print("1. Release 元数据")
    try:
        status, rel = get(f"https://api.github.com/repos/{OWNER}/{REPO}/releases/tags/{args.tag}", token)
    except urllib.error.HTTPError as exc:
        print(f"  无法读取 Release：{exc.code} {exc.reason}")
        return 1
    check("Release 存在", True, rel["tag_name"])
    check("不是草稿", not rel["draft"])
    check("不是预发布", not rel["prerelease"])
    check("有标题", bool(rel.get("name")), rel.get("name", ""))
    check("有说明正文", len(rel.get("body") or "") > 200, f"{len(rel.get('body') or '')} 字符")
    print(f"       页面: {rel['html_url']}")
    print(f"       发布: {rel.get('published_at', '')[:19]}")

    print("\n2. 附件")
    assets = rel.get("assets", [])
    check("有附件", bool(assets), f"{len(assets)} 个")
    if not assets:
        return 1
    asset = assets[0]
    check("附件名是 ASCII（非 ASCII 会被上传接口丢字符）",
          asset["name"].isascii(), asset["name"])
    check("附件大小合理（>20 MB）", asset["size"] > 20 * 1024 * 1024,
          f"{asset['size']/1024/1024:.1f} MB")
    check("附件状态为已上传", asset["state"] == "uploaded", asset["state"])
    print(f"       名称: {asset['name']}")
    print(f"       下载: {asset['browser_download_url']}")

    print("\n3. 实际下载并校验内容")
    status, blob = get(asset["browser_download_url"], token, binary=True)
    check("能下载", status == 200, f"{len(blob)/1024/1024:.1f} MB")
    check("下载大小与元数据一致", len(blob) == asset["size"],
          f"{len(blob)} vs {asset['size']}")
    try:
        zf = zipfile.ZipFile(io.BytesIO(blob))
        bad = zf.testzip()
        check("zip 完整可解压", bad is None, bad or "CRC 全部通过")
        names = zf.namelist()
        print(f"       条目数: {len(names)}")
        for want in EXPECTED_IN_ZIP:
            check(f"包含 {want}", any(n.endswith(want) for n in names))
        # feature markers: proves the asset carries the current code, not a stale build
        for member, needles in MARKERS.items():
            path = next((n for n in names if n.endswith(member)), None)
            if not path:
                check(f"{member} 存在", False)
                continue
            text = zf.read(path).decode("utf-8", "replace")
            for needle in needles:
                check(f"{os.path.basename(member)} 含 {needle}", needle in text)
        data_files = [n for n in names
                      if "/data/" in n and not n.endswith("/")]
        check("包内 data/ 为空（出厂状态）", not data_files, f"{len(data_files)} 个文件")
    except zipfile.BadZipFile:
        check("zip 完整可解压", False, "不是有效 zip")

    print()
    if failures:
        print(f"验证失败（{len(failures)} 项）：{failures}")
        return 1
    print("全部通过 —— Release 可正常下载使用。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
