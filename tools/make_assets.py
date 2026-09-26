# -*- coding: utf-8 -*-
"""Generate the ExcelFinder icon and Windows version resource (stdlib only).

Draws a small "spreadsheet with a magnifier" glyph and writes:
  assets/excelfinder.ico   multi-size icon (PNG-compressed entry for 256x256)
  assets/version_info.txt  VS_VERSIONINFO consumed by PyInstaller
"""
from __future__ import annotations

import os
import struct
import zlib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")

BG = (0x1F, 0x6F, 0x3C, 255)      # Excel-ish green frame
PAPER = (0xFF, 0xFF, 0xFF, 255)
GRID = (0xC8, 0xD8, 0xC8, 255)
ACCENT = (0xE8, 0x8A, 0x1A, 255)  # magnifier handle
LENS = (0x2A, 0x7F, 0xD4, 255)
LENS_IN = (0xEC, 0xF5, 0xFF, 255)


def blank(size: int) -> list[list[tuple[int, int, int, int]]]:
    return [[(0, 0, 0, 0) for _ in range(size)] for _ in range(size)]


def rect(px, x0, y0, x1, y1, color) -> None:
    n = len(px)
    for y in range(max(0, y0), min(n, y1)):
        row = px[y]
        for x in range(max(0, x0), min(n, x1)):
            row[x] = color


def disc(px, cx: float, cy: float, r: float, color, inner=None, inner_r: float = 0.0) -> None:
    n = len(px)
    r2 = r * r
    for y in range(n):
        for x in range(n):
            dx, dy = x + 0.5 - cx, y + 0.5 - cy
            d2 = dx * dx + dy * dy
            if d2 <= r2:
                if inner is not None and inner_r > 0 and d2 <= inner_r * inner_r:
                    px[y][x] = inner
                else:
                    px[y][x] = color


def thick_line(px, x0, y0, x1, y1, width, color) -> None:
    steps = int(max(abs(x1 - x0), abs(y1 - y0)) * 3) + 1
    for i in range(steps + 1):
        t = i / steps
        cx = x0 + (x1 - x0) * t
        cy = y0 + (y1 - y0) * t
        r = width / 2.0
        n = len(px)
        for y in range(int(cy - r) - 1, int(cy + r) + 2):
            for x in range(int(cx - r) - 1, int(cx + r) + 2):
                if 0 <= x < n and 0 <= y < n:
                    dx, dy = x + 0.5 - cx, y + 0.5 - cy
                    if dx * dx + dy * dy <= r * r:
                        px[y][x] = color


def draw(size: int):
    px = blank(size)
    s = size / 32.0

    def S(v: float) -> int:
        return int(round(v * s))

    # page
    rect(px, S(3), S(2), S(25), S(30), BG)
    rect(px, S(5), S(4), S(23), S(28), PAPER)

    # header band + grid
    rect(px, S(5), S(4), S(23), S(8), (0x2E, 0x8B, 0x57, 255))
    for i in range(1, 5):
        y = int(S(8) + i * S(20 / 5.0))
        rect(px, S(5), y, S(23), y + max(1, S(0.7)), GRID)
    for i in range(1, 4):
        x = int(S(5) + i * S(18 / 4.0))
        rect(px, x, S(8), x + max(1, S(0.6)), S(28), GRID)

    # magnifier over the lower-right
    lens_r = 8.0 * s
    cx, cy = 21.0 * s, 21.0 * s
    disc(px, cx, cy, lens_r + 1.6 * s, (0x14, 0x3A, 0x22, 255))
    disc(px, cx, cy, lens_r, LENS, LENS_IN, lens_r - 1.9 * s)
    thick_line(px, cx + lens_r * 0.72, cy + lens_r * 0.72, 30.0 * s, 30.0 * s, 3.2 * s, ACCENT)
    return px


def bmp_payload(px) -> bytes:
    """BITMAPINFOHEADER + bottom-up BGRA rows + AND mask (classic .ico entry)."""
    n = len(px)
    header = struct.pack("<IiiHHIIiiII", 40, n, n * 2, 1, 32, 0, n * n * 4, 0, 0, 0, 0)
    body = bytearray()
    for y in range(n - 1, -1, -1):
        for x in range(n):
            r, g, b, a = px[y][x]
            body += bytes((b, g, r, a))
    mask_stride = ((n + 31) // 32) * 4
    body += b"\x00" * (mask_stride * n)
    return header + bytes(body)


def png_payload(px) -> bytes:
    """32-bit RGBA PNG, top-down."""
    n = len(px)
    raw = bytearray()
    for y in range(n):
        raw.append(0)
        for x in range(n):
            r, g, b, a = px[y][x]
            raw += bytes((r, g, b, a))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", n, n, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        + chunk(b"IEND", b"")
    )


def write_ico(path: str, sizes=(16, 24, 32, 48, 64, 128, 256)) -> None:
    entries = []
    for size in sizes:
        px = draw(size)
        blob = png_payload(px) if size >= 64 else bmp_payload(px)
        entries.append((size, blob))

    out = bytearray(struct.pack("<HHH", 0, 1, len(entries)))
    offset = 6 + 16 * len(entries)
    for size, blob in entries:
        dim = 0 if size >= 256 else size
        out += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(blob), offset)
        offset += len(blob)
    for _size, blob in entries:
        out += blob
    with open(path, "wb") as fh:
        fh.write(out)


VERSION_TXT = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '080404B0',
        [StringStruct('CompanyName', 'ExcelFinder'),
         StringStruct('FileDescription', 'ExcelFinder - 本地 Excel 极速检索'),
         StringStruct('FileVersion', '1.0.0.0'),
         StringStruct('InternalName', 'ExcelFinder'),
         StringStruct('LegalCopyright', 'Provided as-is'),
         StringStruct('OriginalFilename', 'ExcelFinder.exe'),
         StringStruct('ProductName', 'ExcelFinder'),
         StringStruct('ProductVersion', '1.0.0.0')])
    ]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ]
)
"""


def main() -> int:
    os.makedirs(ASSETS, exist_ok=True)
    ico = os.path.join(ASSETS, "excelfinder.ico")
    write_ico(ico)
    vi = os.path.join(ASSETS, "version_info.txt")
    with open(vi, "w", encoding="utf-8") as fh:
        fh.write(VERSION_TXT)
    print(f"wrote {ico} ({os.path.getsize(ico)} bytes)")
    print(f"wrote {vi}")

    # sanity: the ICO header must parse back
    with open(ico, "rb") as fh:
        data = fh.read()
    reserved, kind, count = struct.unpack("<HHH", data[:6])
    print(f"ICONDIR: reserved={reserved} type={kind} images={count}")
    for i in range(count):
        w, h, _c, _r, _planes, bpp, length, offset = struct.unpack(
            "<BBBBHHII", data[6 + 16 * i: 22 + 16 * i]
        )
        sig = data[offset: offset + 8]
        fmt = "PNG" if sig.startswith(b"\x89PNG") else "BMP"
        print(f"  {w or 256}x{h or 256} {bpp}bpp {fmt} {length} bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
