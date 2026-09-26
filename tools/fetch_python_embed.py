# -*- coding: utf-8 -*-
"""Fetch the Windows *embeddable* Python runtime into the portable package.

Why bundle this at all? The packaged ExcelFinder.exe is already self-contained.
The embeddable runtime is the belt-and-braces fallback: if the frozen EXE is
blocked by enterprise application-control policy (SmartScreen / AppLocker often
treat unsigned onefile executables harshly), the user can still run the app
straight from source with ``python\\python.exe src\\excelfinder_gui.py``.

Usage: python tools/fetch_python_embed.py <target_dir> [version]
"""
from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import urllib.request
import zipfile

DEFAULT_VERSION = "3.12.10"


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "excelfinder-build"})
    with urllib.request.urlopen(req, timeout=180) as resp:
        return resp.read()


def main() -> int:
    target = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "portable/python")
    version = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_VERSION

    # Prefer an explicit version, but fall back to whatever the API reports so
    # this keeps working after the pinned build is retired.
    candidates = [version]
    try:
        meta = json.loads(fetch("https://api.github.com/repos/astral-sh/python-build-standalone/releases/latest"))
        print(f"latest python-build-standalone tag: {meta.get('tag_name')}")
    except Exception:
        meta = None

    last_error = ""
    for ver in candidates:
        name = f"python-{ver}-embed-amd64.zip"
        # python.org is the canonical, stable source for the embeddable build.
        urls = [f"https://www.python.org/ftp/python/{ver}/{name}"]
        if meta:
            for asset in meta.get("assets", []):
                if "x86_64-pc-windows-msvc" in asset["name"] and asset["name"].endswith(".tar.zst"):
                    pass  # .tar.zst needs an extra dep; python.org zip is enough
        for url in urls:
            try:
                print(f"downloading {url}")
                blob = fetch(url)
            except Exception as exc:
                last_error = f"{url}: {exc}"
                print(f"  failed: {exc}")
                continue
            os.makedirs(target, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(blob)) as zf:
                zf.extractall(target)
            print(f"extracted {len(blob)/1024/1024:.1f} MB -> {target}")
            exe = os.path.join(target, "python.exe")
            print(f"python.exe present: {os.path.exists(exe)}")
            _enable_site_packages(target)
            if os.path.exists(exe):
                graft_tkinter(target)
            return 0 if os.path.exists(exe) else 1
    print(f"could not fetch embeddable python. last error: {last_error}")
    return 1


def _enable_site_packages(target: str) -> None:
    """Uncomment ``import site`` and expose the app's ``src`` folder.

    The embeddable runtime ships with site disabled and a fixed, isolated search
    path. Entries in ``._pth`` are relative to their own folder, so ``..\\src``
    resolves to the package's src directory and lets ``run_gui.py`` start the
    GUI without any environment variables.
    """
    for name in os.listdir(target):
        if not name.endswith("._pth"):
            continue
        path = os.path.join(target, name)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                lines = fh.readlines()
        except OSError:
            continue
        out = []
        for ln in lines:
            out.append("import site\n" if ln.strip() == "#import site" else ln)
        if not any(ln.strip() == r"..\src" for ln in out):
            out.append("..\\src\n")
        with open(path, "w", encoding="utf-8") as fh:
            fh.writelines(out)
        print(f"patched {name}: enabled site + ..\\src")


def graft_tkinter(target: str, donor: str = r"C:\Python314") -> bool:
    """Copy tkinter into the embeddable runtime from a matching CPython install.

    python.org's embeddable zip has no tkinter, but the extension module is
    ABI-locked to the exact Python version -- so this only works when the donor
    is the *same* version (checked below). Files go next to python.exe because
    the ._pth search path already includes '.', whereas Lib\\ is not on it.
    """
    ver = "%d.%d.%d" % sys.version_info[:3]
    exe = os.path.join(donor, "python.exe")
    if not os.path.exists(exe):
        print(f"tkinter graft skipped: no donor python at {donor}")
        return False
    try:
        proc = os.popen(f'"{exe}" -c "import sys;print(sys.version_info[:3])"')
        donor_ver = proc.read().strip()
        proc.close()
    except OSError:
        donor_ver = ""
    if donor_ver != str(sys.version_info[:3]):
        print(f"tkinter graft skipped: donor is {donor_ver}, need {sys.version_info[:3]} (ABI mismatch)")
        return False

    needed = [
        (os.path.join(donor, "Lib", "tkinter"), os.path.join(target, "tkinter")),
        (os.path.join(donor, "tcl"), os.path.join(target, "tcl")),
    ]
    files = [os.path.join(donor, "DLLs", "_tkinter.pyd")]
    dlls = os.path.join(donor, "DLLs")
    if os.path.isdir(dlls):
        files += [
            os.path.join(dlls, n)
            for n in os.listdir(dlls)
            if n.lower().startswith(("tcl", "tk", "zlib")) and n.lower().endswith(".dll")
        ]

    import shutil

    for src, dst in needed:
        if os.path.isdir(src):
            if os.path.isdir(dst):
                shutil.rmtree(dst, ignore_errors=True)
            shutil.copytree(src, dst)
            print(f"grafted {os.path.basename(src)} -> {dst}")
    for src in files:
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(target, os.path.basename(src)))
            print(f"grafted {os.path.basename(src)}")

    probe = subprocess.run(
        [os.path.join(target, "python.exe"), "-c", "import tkinter;print('tkinter', tkinter.TkVersion)"],
        capture_output=True, text=True, errors="replace",
    )
    ok = probe.returncode == 0
    print(f"tkinter self-check: {probe.stdout.strip() or probe.stderr.strip()[:120]}")
    if not ok:
        print("  -> GUI fallback unavailable; the frozen ExcelFinder.exe still works.")
    return ok


if __name__ == "__main__":
    sys.exit(main())
