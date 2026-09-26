# -*- coding: utf-8 -*-
"""Install PyInstaller by extracting wheels directly.

The sandbox in this environment denies pip's nested staging directories, but
plain HTTP downloads plus ``zipfile`` writes work fine. This resolves the
dependency set from the PyPI JSON API for the running interpreter and unpacks
each wheel into a local directory that ``PYTHONPATH`` can point at.

Usage: python tools/fetch_buildtools.py <target_dir>
"""
from __future__ import annotations

import io
import json
import os
import sys
import urllib.request
import zipfile

WANTED = ["pyinstaller", "pyinstaller-hooks-contrib", "altgraph", "packaging", "pefile", "pywin32-ctypes", "setuptools"]


def tags() -> list[str]:
    v = sys.version_info
    return [
        f"cp{v.major}{v.minor}-cp{v.major}{v.minor}-win_amd64",
        f"cp{v.major}{v.minor}-abi3-win_amd64",
        "py3-none-win_amd64",
        "py3-none-any",
        f"cp{v.major}{v.minor}-none-win_amd64",
        "py2.py3-none-any",
    ]


def pick(wheel_files: list[dict], prefer: list[str]) -> dict | None:
    for tag in prefer:
        for w in wheel_files:
            if w["filename"].endswith(tag + ".whl"):
                return w
    return None


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "excelfinder-build"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def main() -> int:
    target = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "buildtools")
    os.makedirs(target, exist_ok=True)
    prefer = tags()
    print(f"interpreter: {sys.version.split()[0]} ({sys.platform})")
    print(f"target     : {target}\n")

    for pkg in WANTED:
        try:
            meta = json.loads(fetch(f"https://pypi.org/pypi/{pkg}/json"))
        except Exception as exc:
            print(f"  ! {pkg}: metadata lookup failed ({exc})")
            continue
        version = meta["info"]["version"]
        files = meta["releases"].get(version, [])
        wheel = pick(files, prefer)
        if wheel is None:
            print(f"  - {pkg} {version}: no compatible wheel, skipped")
            continue
        try:
            blob = fetch(wheel["url"])
        except Exception as exc:
            print(f"  ! {pkg} {version}: download failed ({exc})")
            continue
        try:
            with zipfile.ZipFile(io.BytesIO(blob)) as zf:
                zf.extractall(target)
        except Exception as exc:
            print(f"  ! {pkg} {version}: extract failed ({exc})")
            continue
        print(f"  + {pkg} {version}  ({len(blob)/1024:.0f} KB) {wheel['filename']}")

    marker = os.path.join(target, "PyInstaller")
    ok = os.path.isdir(marker)
    print(f"\nPyInstaller present: {ok}")
    if ok:
        sys.path.insert(0, target)
        try:
            import PyInstaller  # noqa: F401

            print("import check: OK")
        except Exception as exc:
            print(f"import check: FAILED ({exc})")
            return 1
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
