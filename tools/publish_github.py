# -*- coding: utf-8 -*-
"""Create the GitHub repository and push.

Kept as a script so the token is never echoed into logs, and so the exact API
payload (description / topics / homepage) is reviewable instead of living only
in a shell history.

Usage:  set GITHUB_TOKEN=ghp_...  then  python tools/publish_github.py
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import urllib.error
import urllib.request

REPO = "ExcelFinder"
DESCRIPTION = ("免安装的 Windows Excel 极速检索工具：为目录建内存索引，毫秒级搜文件名与单元格内容，"
               "并定位到「工作表 + 单元格」。单文件 exe，零依赖，支持 WPS 与 U盘便携模式。")
TOPICS = ["excel", "search", "windows", "python", "tkinter", "xlsx", "wps",
          "full-text-search", "portable", "no-dependencies", "openpyxl-alternative",
          "file-search"]
HOMEPAGE = ""


def api(url: str, token: str, method: str = "GET", payload: dict | None = None):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "excelfinder-publish",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
            return resp.status, (json.loads(body) if body.strip() else {})
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "replace")
        try:
            detail = json.loads(detail).get("message", detail)
        except ValueError:
            pass
        return exc.code, {"message": detail}


def run(cmd: list[str], token: str | None = None) -> tuple[int, str]:
    env = dict(os.environ)
    if token:
        # Feed git the PAT non-interactively; never write it to .git/config.
        env["GIT_ASKPASS"] = "echo"
        env["GIT_TERMINAL_PROMPT"] = "0"
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", env=env)
    out = (proc.stdout or "") + (proc.stderr or "")
    return proc.returncode, out.strip()


def main() -> int:
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if not token:
        print("请先设置环境变量 GITHUB_TOKEN")
        return 2

    status, me = api("https://api.github.com/user", token)
    if status != 200:
        print(f"令牌无效（{status}）：{me.get('message')}")
        return 1
    owner = me["login"]
    print(f"登录身份: {owner}")

    # 1. create (or reuse) the repository
    status, data = api("https://api.github.com/user/repos", token, "POST", {
        "name": REPO,
        "description": DESCRIPTION,
        "homepage": HOMEPAGE,
        "private": False,
        "has_issues": True,
        "has_wiki": False,
        "has_projects": False,
        "auto_init": False,
    })
    if status == 201:
        print(f"仓库已创建: {data['html_url']}")
        clone_url = data["clone_url"]
    elif status == 422:
        status2, data2 = api(f"https://api.github.com/repos/{owner}/{REPO}", token)
        if status2 != 200:
            print(f"仓库已存在但无法读取（{status2}）：{data2.get('message')}")
            return 1
        print(f"仓库已存在，复用: {data2['html_url']}")
        clone_url = data2["clone_url"]
    else:
        print(f"创建失败（{status}）：{data.get('message')}")
        return 1

    # 2. topics
    status, data = api(f"https://api.github.com/repos/{owner}/{REPO}/topics", token,
                       "PUT", {"names": TOPICS})
    topic_note = "OK" if status == 200 else f"{status} {data.get('message')}"
    print(f"话题设置: {topic_note}")

    # 3. branch -> main, then push
    run(["git", "branch", "-M", "main"])
    remote = f"https://{owner}:{token}@github.com/{owner}/{REPO}.git"
    run(["git", "remote", "remove", "origin"])
    code, out = run(["git", "remote", "add", "origin", remote])
    if code != 0:
        print(f"添加远程失败: {out}")
        return 1

    code, out = run(["git", "push", "-u", "origin", "main"])
    # Scrub anything token-shaped before printing.
    safe = out.replace(token, "***")
    print(f"推送{'成功' if code == 0 else '失败'}:\n{safe}")

    # 4. rewrite the remote without the token so it is not left on disk
    run(["git", "remote", "set-url", "origin",
         f"https://github.com/{owner}/{REPO}.git"])

    if code != 0:
        return 1

    status, data = api(f"https://api.github.com/repos/{owner}/{REPO}", token)
    if status == 200:
        print(f"\n仓库地址 : {data['html_url']}")
        print(f"默认分支 : {data['default_branch']}")
        print(f"体积     : {data['size']} KB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
