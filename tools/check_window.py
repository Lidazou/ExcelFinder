# -*- coding: utf-8 -*-
"""Launch a process, then enumerate its top-level Win32 windows.

Proves whether a GUI build really created a visible window, instead of relying
on ``MainWindowTitle`` (which is empty whenever the window has no caption) or on
the process staying alive (which background threads can fake).

Usage: python tools/check_window.py <exe> [args...]
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os
import subprocess
import sys
import time

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

EnumWindowsProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

user32.EnumWindows.argtypes = [EnumWindowsProc, wt.LPARAM]
user32.GetWindowThreadProcessId.argtypes = [wt.HWND, ctypes.POINTER(wt.DWORD)]
user32.GetWindowTextLengthW.argtypes = [wt.HWND]
user32.GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.IsWindowVisible.argtypes = [wt.HWND]
user32.GetClassNameW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
user32.GetWindowRect.argtypes = [wt.HWND, ctypes.POINTER(wt.RECT)]

# PyInstaller's onefile bootloader re-executes the app in a CHILD process, so
# the real window belongs to a descendant pid, not the one we spawned.
TH32CS_SNAPPROCESS = 0x00000002
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value


class PROCESSENTRY32(ctypes.Structure):
    _fields_ = [
        ("dwSize", wt.DWORD),
        ("cntUsage", wt.DWORD),
        ("th32ProcessID", wt.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wt.DWORD),
        ("cntThreads", wt.DWORD),
        ("th32ParentProcessID", wt.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wt.DWORD),
        ("szExeFile", ctypes.c_char * 260),
    ]


def descendants(root_pid: int) -> set[int]:
    """All live pids whose ancestry includes root_pid (inclusive)."""
    snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if snap == INVALID_HANDLE_VALUE:
        return {root_pid}
    parents: dict[int, int] = {}
    try:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
        ok = kernel32.Process32First(snap, ctypes.byref(entry))
        while ok:
            parents[int(entry.th32ProcessID)] = int(entry.th32ParentProcessID)
            ok = kernel32.Process32Next(snap, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snap)

    result = {root_pid}
    changed = True
    while changed:                      # transitive closure, tree is tiny
        changed = False
        for pid, ppid in parents.items():
            if pid not in result and ppid in result:
                result.add(pid)
                changed = True
    return result


def windows_of(pids: set[int]) -> list[dict]:
    """Top-level windows owned by any pid in `pids`."""
    found: list[dict] = []

    def cb(hwnd, _lparam):
        wpid = wt.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(wpid))
        if wpid.value not in pids:
            return True
        n = user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(n + 2)
        user32.GetWindowTextW(hwnd, buf, n + 2)
        cls = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls, 256)
        rect = wt.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        found.append(
            {
                "hwnd": hwnd,
                "pid": int(wpid.value),
                "title": buf.value,
                "class": cls.value,
                "visible": bool(user32.IsWindowVisible(hwnd)),
                "size": (rect.right - rect.left, rect.bottom - rect.top),
            }
        )
        return True

    user32.EnumWindows(EnumWindowsProc(cb), 0)
    return found


WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
user32.SendMessageW.restype = ctypes.c_long
user32.GetWindow.argtypes = [wt.HWND, wt.UINT]


def control_texts(parent: int) -> list[str]:
    """Read the text of every child control under `parent`.

    ttk renders Labels and Entry widgets as native STATIC/EDIT controls, so
    WM_GETTEXT reaches the status bar and query box even though tkinter exposes
    nothing through UI Automation.
    """
    out: list[str] = []
    hwnd = user32.GetWindow(parent, 5)          # GW_CHILD
    while hwnd:
        length = user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
        if length and length > 0:
            buf = ctypes.create_unicode_buffer(int(length) + 2)
            user32.SendMessageW(hwnd, WM_GETTEXT, int(length) + 1, ctypes.byref(buf))
            if buf.value.strip():
                out.append(buf.value.strip())
        out.extend(control_texts(hwnd))          # recurse into descendants
        hwnd = user32.GetWindow(hwnd, 2)         # GW_HWNDNEXT
    return out


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    exe = sys.argv[1]
    args = sys.argv[2:]
    env = dict(os.environ)
    env["TMP"] = env["TEMP"] = os.path.join(os.getcwd(), ".runtime-tmp")
    os.makedirs(env["TMP"], exist_ok=True)

    proc = subprocess.Popen([exe, *args], env=env)
    print(f"launched {exe} {args} -> pid {proc.pid}")

    best: list[dict] = []
    for i in range(24):                     # poll for up to 12 s
        time.sleep(0.5)
        code = proc.poll()
        if code is not None:
            print(f"  !! process EXITED early with code {code} after {(i+1)*0.5:.1f}s")
            break
        wins = windows_of(descendants(proc.pid))
        visible = [w for w in wins if w["visible"] and w["size"][0] > 50]
        if visible:
            best = visible
            break

    if best:
        print(f"  OK: {len(best)} visible top-level window(s) after {(i+1)*0.5:.1f}s")
        for w in best:
            print(f"     pid={w['pid']} class={w['class']!r} title={w['title']!r} size={w['size']}")
        time.sleep(float(os.environ.get("CHECK_WINDOW_SETTLE", "0")))
        texts = control_texts(best[0]["hwnd"])
        print(f"  child control texts ({len(texts)}):")
        for t in texts:
            print(f"     | {t[:160]}")
    elif proc.poll() is None:
        allw = windows_of(descendants(proc.pid))
        print(f"  process alive but no visible window; {len(allw)} window(s) total:")
        for w in allw:
            print(f"     pid={w['pid']} class={w['class']!r} title={w['title']!r} "
                  f"visible={w['visible']} size={w['size']}")

    time.sleep(1.0)
    if proc.poll() is None:
        # taskkill /T is required: the PyInstaller bootloader's child owns the
        # real window and survives a plain terminate(), keeping the EXE locked.
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
            capture_output=True, text=True,
        )
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
        print("  process tree terminated by harness")
    return 0 if best else 1


if __name__ == "__main__":
    sys.exit(main())
