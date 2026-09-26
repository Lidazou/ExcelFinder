# -*- coding: utf-8 -*-
"""Publish a release and upload the portable zip as its asset.

Release assets are used instead of committing the archive: a 35 MB blob in git
history is permanent and every clone pays for it, while a release asset can be
replaced or deleted at will.

The token is read from GITHUB_TOKEN, never written to disk, and scrubbed from
any output.

Usage:
    set GITHUB_TOKEN=ghp_...
    python tools/publish_release.py --zip "dist\\ExcelFinder-portable.zip" --tag v1.0.0
"""
from __future__ import annotations

import argparse
import json
import mimetypes
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

OWNER = os.environ.get("GITHUB_OWNER", "Lidazou")
REPO = os.environ.get("GITHUB_REPO", "ExcelFinder")

NOTES = """## 下载

| 文件 | 说明 |
| --- | --- |
| `ExcelFinder-U盘版.zip` | **完整便携包**，解压后双击 `ExcelFinder.exe` 即可使用 |

解压后目录里已经包含使用说明（`先读我.txt`），无需安装 Python、Office 或任何运行库。

### 快速上手

1. 解压 zip，双击 `ExcelFinder.exe`
2. 「浏览…」选目录 → 点「重建索引」
3. 输入关键词 → 回车

**要找单元格里的内容**（比如某个人名）：取消勾选「文件名需命中」，只勾「内容需命中」。
首次会解析所有文件（底部有进度条，约 270 文件/秒），完成后自动重新搜索，之后走缓存是毫秒级。

### 本版本包含

- **单元格级定位** —— 结果直接给出「工作表!单元格」，双击用 Excel / WPS 打开并跳转选中
- **打开方式可选** —— 自动 / 只用 Office / 只用 WPS（注册表探测，不靠启动进程试探）
- **刷新列表 (F5)** —— 增量重扫目录并重跑当前查询，状态栏报告增删变化
- **分阶段进度条** —— 扫描 / 解析内容 / 检索 / 定位单元格各有进度
- **U盘便携模式** —— 配置与索引存于程序目录，不在宿主电脑留任何东西
- **索引缓存保护** —— 超过 256 MB 不再写磁盘，避免悄悄吃掉你的空间

### 性能（实测 6 万文件 / 424 MB）

| 方案 | 耗时 | 相对加速 |
| --- | ---: | ---: |
| PowerShell `Get-ChildItem -Recurse` | 2388 ms | 1× |
| CMD `where /r`（系统自带） | 919 ms | 2.6× |
| ExcelFinder 首次扫描 | 468 ms | 5.1× |
| **ExcelFinder 已建索引检索** | **42 ms** | **57×** |

### 运行环境

- **必须 64 位 Windows**（建议 Windows 10 / 11）
- 不需要安装 Python、VC++ 运行库、.NET 或 Java（全部已内嵌）
- Office / WPS 是可选的：不装也能搜索、浏览、导出 CSV，装了才能"双击跳到单元格"

### 注意

本程序**没有数字签名**，个别电脑的杀毒软件可能误报。遇到时：
右键 exe → 属性 → 勾选「解除锁定」，或双击包内的 `备用启动.bat`。

完整文档见 [README](https://github.com/Lidazou/ExcelFinder#readme)。
"""


def api(url: str, token: str, method: str = "GET", payload: dict | None = None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "excelfinder-release",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = resp.read().decode("utf-8")
            return resp.status, (json.loads(body) if body.strip() else {})
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        try:
            detail = json.loads(detail).get("message", detail)
        except ValueError:
            pass
        return exc.code, {"message": detail}


def upload_asset(upload_url: str, token: str, path: str, name: str) -> tuple[int, dict]:
    """Upload one file as a release asset (streamed, no third-party client)."""
    size = os.path.getsize(path)
    ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
    url = upload_url.split("{")[0] + f"?name={urllib.parse.quote(name)}"
    req = urllib.request.Request(url, data=open(path, "rb"), method="POST", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "Content-Type": ctype,
        "Content-Length": str(size),
        "User-Agent": "excelfinder-release",
    })
    try:
        with urllib.request.urlopen(req, timeout=1800) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        try:
            detail = json.loads(detail).get("message", detail)
        except ValueError:
            pass
        return exc.code, {"message": detail}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip", required=True, help="要上传的 zip 路径")
    ap.add_argument("--tag", default="v1.0.0")
    ap.add_argument("--name", default="ExcelFinder v1.0.0 — U盘便携版")
    args = ap.parse_args()

    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if not token:
        print("请先设置 GITHUB_TOKEN")
        return 2
    path = os.path.abspath(args.zip)
    if not os.path.isfile(path):
        print(f"找不到文件：{path}")
        return 1
    size_mb = os.path.getsize(path) / 1024 / 1024
    print(f"待上传: {os.path.basename(path)}  ({size_mb:.1f} MB)")

    base = f"https://api.github.com/repos/{OWNER}/{REPO}"

    # create the release, or reuse it if this tag already has one
    status, rel = api(f"{base}/releases", token, "POST", {
        "tag_name": args.tag,
        "target_commitish": "main",
        "name": args.name,
        "body": NOTES,
        "draft": False,
        "prerelease": False,
    })
    if status == 201:
        print(f"Release 已创建: {rel['html_url']}")
    elif status == 422:
        status2, rel = api(f"{base}/releases/tags/{args.tag}", token)
        if status2 != 200:
            print(f"创建失败（{status}）且无法读取已有 Release：{rel.get('message')}")
            return 1
        print(f"Release 已存在，复用: {rel['html_url']}")
    else:
        print(f"创建失败（{status}）：{rel.get('message')}")
        return 1

    # skip re-uploading an asset that is already there
    existing = {a["name"]: a for a in rel.get("assets", [])}
    name = os.path.basename(path)
    if name in existing:
        print(f"资产已存在，先删除旧的上传: {name}")
        api(f"{base}/releases/assets/{existing[name]['id']}", token, "DELETE")

    print("上传中（35 MB 左右，请稍候）…")
    status, asset = upload_asset(rel["upload_url"], token, path, name)
    if status not in (200, 201):
        print(f"上传失败（{status}）：{asset.get('message')}")
        return 1
    print(f"上传成功: {asset['name']}  {asset['size']/1024/1024:.1f} MB")
    print(f"下载地址: {asset['browser_download_url']}")

    # report the release as GitHub now sees it
    status, rel2 = api(f"{base}/releases/tags/{args.tag}", token)
    if status == 200:
        print(f"\nRelease 页面: {rel2['html_url']}")
        print(f"标签        : {rel2['tag_name']}")
        print(f"附件        : {len(rel2['assets'])} 个")
        for a in rel2["assets"]:
            print(f"   - {a['name']}  ({a['size']/1024/1024:.1f} MB, 下载 {a['download_count']} 次)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
