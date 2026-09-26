# -*- coding: utf-8 -*-
"""Verify what actually landed on GitHub after a publish.

Checks the repository metadata, that every expected file is present, and that
the README's image references resolve -- a README pointing at missing images is
the most common way a repo ends up looking broken on the web.

Usage:  set GITHUB_TOKEN=...  then  python tools/verify_github.py
"""
from __future__ import annotations

import json
import os
import re
import sys
import urllib.error
import urllib.request

OWNER = os.environ.get("GITHUB_OWNER", "Lidazou")
REPO = os.environ.get("GITHUB_REPO", "ExcelFinder")

EXPECTED = [
    "README.md", "LICENSE", ".gitignore", ".gitattributes",
    "src/excelfinder_core.py", "src/excelfinder_gui.py", "src/excelfinder_cli.py",
    "tools/build.py", "tools/bench.py", "tools/make_diagrams.py",
    "docs/images/hero.svg", "docs/images/architecture.svg",
    "docs/images/performance.svg", "docs/images/features.svg",
]

MUST_NOT = ["ExcelFinder.exe", "ExcelFinder-cli.exe", "data/index_v1.json",
            "buildtools", "embed-python", "testdata"]


def fetch(url: str, token: str | None = None, raw: bool = False):
    headers = {"User-Agent": "excelfinder-verify", "Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read()
            return resp.status, (body.decode("utf-8", "replace") if raw else json.loads(body))
    except urllib.error.HTTPError as exc:
        return exc.code, None
    except Exception as exc:                                   # noqa: BLE001
        return 0, str(exc)


def main() -> int:
    token = os.environ.get("GITHUB_TOKEN") or None
    base = f"https://api.github.com/repos/{OWNER}/{REPO}"
    failures: list[str] = []

    def check(label: str, ok: bool, detail: str = ""):
        print(f"  [{'OK ' if ok else 'BAD'}] {label}{(' -- ' + detail) if detail else ''}")
        if not ok:
            failures.append(label)

    print(f"检查仓库 {OWNER}/{REPO}\n")

    print("1. 仓库信息")
    status, repo = fetch(base, token)
    if status != 200 or not isinstance(repo, dict):
        print(f"  无法读取仓库：{status} {repo}")
        return 1
    check("仓库可访问", True, repo["full_name"])
    check("是公开仓库", not repo["private"])
    check("默认分支为 main", repo["default_branch"] == "main", repo["default_branch"])
    check("已填写描述", bool(repo.get("description")))
    check("已设置话题", len(repo.get("topics") or []) >= 5,
          f"{len(repo.get('topics') or [])} 个")
    print(f"       地址: {repo['html_url']}")
    print(f"       体积: {repo['size']} KB")

    print("\n2. 文件清单")
    status, tree = fetch(f"{base}/git/trees/{repo['default_branch']}?recursive=1", token)
    paths = {t["path"] for t in tree.get("tree", []) if t["type"] == "blob"} if tree else set()
    check("能列出文件树", bool(paths), f"{len(paths)} 个文件")
    for want in EXPECTED:
        check(f"存在 {want}", want in paths)
    for bad in MUST_NOT:
        hit = [p for p in paths if p == bad or p.startswith(bad + "/")]
        check(f"未误传 {bad}", not hit, f"{len(hit)} 个" if hit else "")

    print("\n3. README 图片引用")
    status, md = fetch(f"https://raw.githubusercontent.com/{OWNER}/{REPO}/"
                       f"{repo['default_branch']}/README.md", token, raw=True)
    if status != 200 or not isinstance(md, str):
        check("能读取 README", False, str(status))
        return 1
    check("能读取 README", True, f"{len(md):,} 字符")
    refs = re.findall(r'(?:src|href)="([^"]+)"', md)
    local = [r for r in refs if not r.startswith(("http://", "https://", "#", "mailto:"))]
    check("README 含图片引用", any(r.endswith(".svg") for r in local), f"{len(local)} 个本地引用")
    for ref in sorted(set(local)):
        target = ref.split("#")[0]
        ok = target in paths
        check(f"引用可解析: {ref}", ok)
    check("含徽章", "img.shields.io" in md)
    check("含英文摘要", "English summary" in md)

    print("\n4. 提交历史")
    status, commits = fetch(f"{base}/commits?per_page=5", token)
    if isinstance(commits, list) and commits:
        check("至少一个提交", True, f"{len(commits)} 个（显示前 5）")
        for c in commits[:5]:
            print(f"       {c['sha'][:8]}  {c['commit']['message'].splitlines()[0][:60]}")
    else:
        check("至少一个提交", False)

    print()
    if failures:
        print(f"验证失败（{len(failures)} 项）：{failures}")
        return 1
    print("全部通过 —— 仓库已就绪。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
