# -*- coding: utf-8 -*-
"""One-shot build + verification pipeline for ExcelFinder.

    python tools/build.py              # build EXEs, assemble package, verify
    python tools/build.py --quick      # skip the slow content-index checks
    python tools/build.py --zip        # also emit dist/ExcelFinder-portable.zip

Steps
  1. assets      - icon + Windows version resource
  2. testdata    - small synthetic corpus (only if missing)
  3. exe         - PyInstaller GUI + console builds
  4. package     - assemble portable/
  5. verify      - engine bench, GUI integration, packaged EXE, portable layout
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable
BUILDTOOLS = os.path.join(ROOT, "buildtools")
RUNTIME_TMP = os.path.join(ROOT, ".runtime-tmp")


def banner(step: str, title: str) -> None:
    print("\n" + "#" * 74)
    print(f"# {step}  {title}")
    print("#" * 74, flush=True)


def run(cmd: list[str], env: dict | None = None, cwd: str = ROOT, check: bool = True,
        echo: bool = True) -> int:
    """Run a step, streaming output when `echo`, else capturing it to a log.

    PyInstaller is extremely chatty and writes progress to stderr, which makes
    PowerShell wrap the whole run as an error. Noisy steps are logged to
    build-logs/ instead so the build report stays readable.
    """
    label = os.path.basename(cmd[0])
    if cmd[0] == PY:
        rest = cmd[1:]
        if rest and rest[0] == "-m" and len(rest) > 1:
            label = rest[1]                        # python -m PyInstaller -> PyInstaller
        elif rest:
            label = os.path.basename(rest[0])      # python tools/foo.py -> foo.py
    if echo:
        print(f"$ {' '.join(cmd)}\n", flush=True)
        proc = subprocess.run(cmd, cwd=cwd, env=env)
    else:
        log_dir = os.path.join(ROOT, "build-logs")
        os.makedirs(log_dir, exist_ok=True)
        log = os.path.join(log_dir, f"{label}.log")
        print(f"$ {' '.join(cmd)}")
        print(f"    输出记录到 build-logs/{os.path.basename(log)} ...", flush=True)
        with open(log, "w", encoding="utf-8", errors="replace") as fh:
            proc = subprocess.run(cmd, cwd=cwd, env=env, stdout=fh, stderr=subprocess.STDOUT)
        tail = _tail(log, 6)
        if proc.returncode != 0 and tail:
            print(tail, flush=True)
    if check and proc.returncode != 0:
        raise SystemExit(f"step failed with exit code {proc.returncode}: {cmd[0]}")
    return proc.returncode


def _tail(path: str, n: int) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            lines = [ln.rstrip() for ln in fh if ln.strip()]
        return "\n".join(f"    | {ln}" for ln in lines[-n:])
    except OSError:
        return ""


def base_env() -> dict:
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONPATH"] = BUILDTOOLS + os.pathsep + env.get("PYTHONPATH", "")
    os.makedirs(RUNTIME_TMP, exist_ok=True)
    env["TMP"] = env["TEMP"] = RUNTIME_TMP
    return env


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true", help="skip slow content-index checks")
    ap.add_argument("--zip", action="store_true", help="produce a distributable zip")
    ap.add_argument("--skip-build", action="store_true", help="reuse existing dist/*.exe")
    args = ap.parse_args()
    env = base_env()
    started = time.time()
    results: list[tuple[str, bool, float]] = []

    def step(name: str, fn) -> None:
        t0 = time.perf_counter()
        try:
            fn()
            ok = True
        except SystemExit as exc:
            print(f"\n!! {name}: {exc}", flush=True)
            ok = False
        results.append((name, ok, time.perf_counter() - t0))

    banner("1/6", "生成图标与版本资源")
    step("assets", lambda: run([PY, os.path.join(HERE, "make_assets.py")], env, echo=False))

    def ensure_testdata() -> None:
        corpus = os.path.join(ROOT, "testdata")
        if os.path.isdir(corpus) and any(
            f.endswith(".xlsx") for _b, _d, fs in os.walk(corpus) for f in fs
        ):
            print(f"复用已有测试数据：{corpus}")
            return
        run([PY, os.path.join(HERE, "make_testdata.py"), corpus,
             "--files", "400", "--hardlinks", "3000", "--rows", "25", "--clean"], env)

    banner("2/6", "准备测试数据与内置运行时")
    step("testdata", ensure_testdata)

    def prepare_runtime() -> None:
        """Prune the bundled Python: 1000+ mostly-useless files otherwise."""
        target = os.path.join(ROOT, "embed-python")
        if not os.path.isdir(target):
            print("embed-python/ 不存在，跳过（分发包仍可用，只是少了备用启动方式）")
            return
        run([PY, os.path.join(HERE, "prune_embed.py"), target], env, echo=False)

    step("prune-runtime", prepare_runtime)

    if not args.skip_build:
        banner("3/6", "PyInstaller 打包（GUI + 命令行）")
        # Always --clean: PyInstaller caches its analysis keyed on the script
        # path, so edits to an *imported* module (excelfinder_core) can be
        # silently ignored, shipping a stale binary that looks freshly built.
        step("exe-gui", lambda: run([PY, "-m", "PyInstaller", "--noconfirm", "--clean",
                                     "--distpath", "dist", "--workpath", "build",
                                     "ExcelFinder.spec"], env, echo=False))
        step("exe-cli", lambda: run([PY, "-m", "PyInstaller", "--noconfirm", "--clean",
                                     "--distpath", "dist", "--workpath", "build-cli",
                                     "ExcelFinder-cli.spec"], env, echo=False))
        for name in ("ExcelFinder.exe", "ExcelFinder-cli.exe"):
            path = os.path.join(ROOT, "dist", name)
            print(f"    {name:<24} {os.path.getsize(path)/1024/1024:6.2f} MB"
                  if os.path.exists(path) else f"    {name} 缺失")
    else:
        results.append(("exe-gui", True, 0.0))
        results.append(("exe-cli", True, 0.0))

    banner("4/6", "组装 U盘便携包")
    pkg_cmd = [PY, os.path.join(HERE, "make_portable.py")]
    if args.zip:
        pkg_cmd.append("--zip")
    step("package", lambda: run(pkg_cmd, env))

    banner("5/6", "验证")
    step("engine-bench", lambda: run([PY, os.path.join(HERE, "bench.py")], env))
    gui_cmd = [PY, os.path.join(HERE, "gui_integration.py"), os.path.join(ROOT, "testdata")]
    if not args.quick:
        gui_cmd.append("--content")
    step("gui-integration", lambda: run(gui_cmd, env))
    step("refresh-button", lambda: run([PY, os.path.join(HERE, "test_refresh.py")], env))
    step("portable", lambda: run([PY, os.path.join(HERE, "verify_portable.py")], env))

    print("\n" + "=" * 74)
    print("构建报告")
    print("=" * 74)
    for name, ok, secs in results:
        print(f"  [{'OK ' if ok else 'BAD'}] {name:<18} {secs:7.1f}s")
    failed = [n for n, ok, _ in results if not ok]
    print(f"\n总耗时 {time.time() - started:.1f}s")
    if failed:
        print(f"失败步骤：{failed}")
        return 1
    print("全部通过。可分发产物：")
    print(f"  {os.path.join(ROOT, 'portable')}          （整个文件夹拷到U盘）")
    if args.zip:
        print(f"  {os.path.join(ROOT, 'dist', 'ExcelFinder-portable.zip')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
