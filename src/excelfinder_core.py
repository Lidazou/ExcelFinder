# -*- coding: utf-8 -*-
"""
ExcelFinder core engine.

Pure standard library. Provides:
  * ScanEngine   - multi-threaded filesystem indexer for Excel/table files
  * Matcher      - keyword matcher (substring / prefix / exact / wildcard / regex, AND-tokens)
  * Index        - in-memory index + instant query + JSON persistence
  * content extraction for .xlsx/.xlsm/.xltx/.csv/.tsv/.txt (and best-effort .xls/.xlsb)

Designed so the GUI layer stays thin and the engine is unit-testable headless.
"""

from __future__ import annotations

import html
import json
import os
import re
import sys
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Callable, Iterable, Sequence

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

APP_NAME = "ExcelFinder"
APP_VERSION = "1.0.0"
INDEX_FORMAT = 3

#: spreadsheet / table formats treated as "Excel files"
EXCEL_EXTS = frozenset(
    {
        ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam",
        ".xls", ".xlt", ".xla",
        ".xlsb",
        ".csv", ".tsv",
        ".ods",
    }
)

#: formats whose text content we can extract natively
CONTENT_EXTS = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb", ".xls", ".csv", ".tsv", ".txt", ".ods"})

#: files we still index by name even if not a spreadsheet
OTHER_EXTS = frozenset({".txt", ".md", ".log"})

#: everything the walker even bothers to stat
ACCEPT_EXTS = EXCEL_EXTS | OTHER_EXTS

SKIP_DIR_NAMES = frozenset(
    {
        "$recycle.bin", "system volume information", "$windows.~ws", "$windows.~bt",
        "windows", "winsxs", "node_modules", ".git", ".svn", ".hg", "__pycache__",
        ".venv", "venv", ".idea", ".vs", "appdata",
    }
)

# --------------------------------------------------------------------------
# Portable (USB / no-install) mode
# --------------------------------------------------------------------------

PORTABLE_MARKER = "portable.marker"


def app_dir() -> str:
    """Directory the application itself lives in.

    For a frozen build that is the folder holding the .exe (never the onefile
    temp dir); running from source it is the project root.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _writable(path: str) -> bool:
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, ".write_probe")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("1")
        os.remove(probe)
        return True
    except OSError:
        return False


def state_dir() -> str:
    """Where config.json and the index cache live.

    Portable mode (a ``portable.marker`` file, or just a ``data`` folder, beside
    the executable) keeps everything on the stick so the drive letter can change
    freely and nothing is left behind on the host PC. Otherwise settings go to
    %APPDATA%\\ExcelFinder as a normal installed app would.
    """
    base = app_dir()
    portable = os.path.join(base, "data")
    marker = os.path.join(base, PORTABLE_MARKER)
    if os.path.exists(marker) or os.path.isdir(portable):
        if _writable(portable):
            return portable
    roaming = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), APP_NAME)
    try:
        os.makedirs(roaming, exist_ok=True)
    except OSError:
        pass
    return roaming


def is_portable() -> bool:
    return os.path.normcase(state_dir()) == os.path.normcase(os.path.join(app_dir(), "data"))

# Directory walks are metadata-bound and hit a sweet spot around 2-4 threads;
# past that, contention on the filesystem cache costs more than the extra
# parallelism buys. Kept user-tunable in the GUI for network shares.
DEFAULT_THREADS = 4
MAX_THREADS = 64
MAX_TEXT_CHARS = 400_000          # per-file text kept in the index
MAX_CONTENT_FILESIZE = 60 * 1024 * 1024
PREVIEW_CHARS = 120


def fmt_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} TB"


def fmt_time(ts: float) -> str:
    try:
        return time.strftime("%Y-%m-%d %H:%M", time.localtime(ts))
    except (OSError, ValueError, OverflowError):
        return ""


def human_int(n: int) -> str:
    return f"{n:,}"


# --------------------------------------------------------------------------
# Content extraction
# --------------------------------------------------------------------------

# Excel for Windows writes "x:" namespace prefixes; some writers omit them.
# Normalise by accepting an optional prefix on every tag.
def _tag(name: str) -> str:
    return r"(?:[A-Za-z0-9_.\-]+:)?%s\b" % name


RE_SI = re.compile(_tag("si") + r"[^>]*>(.*?)</" + _tag("si") + r">", re.S)
RE_T = re.compile(r"<" + _tag("t") + r"[^>]*>(.*?)</" + _tag("t") + r">", re.S)
RE_CELL = re.compile(
    r"<" + _tag("c") + r"\b([^>]*?)(?:/>|>(.*?)</" + _tag("c") + r">)", re.S
)
RE_CELL_REF = re.compile(r'\br\s*=\s*"([A-Za-z]+\d+)"')
RE_CELL_TYPE = re.compile(r'\bt\s*=\s*"([^"]*)"')
RE_V = re.compile(r"<" + _tag("v") + r"[^>]*>(.*?)</" + _tag("v") + r">", re.S)
RE_IS = re.compile(r"<" + _tag("is") + r"[^>]*>(.*?)</" + _tag("is") + r">", re.S)
RE_SHEETNAME = re.compile(r'<sheetPr[^>]*\bcodeName="([^"]*)"|<' + _tag("sheet") + r'[^>]*\bname="([^"]*)"')
RE_DEFINED = re.compile(r"<" + _tag("definedName") + r"[^>]*>(.*?)</" + _tag("definedName") + r">", re.S)

_ENTITIES = (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'"), ("&amp;", "&"))


def _unescape(text: str) -> str:
    if "&" not in text:
        return text
    for src, dst in _ENTITIES:
        if src in text:
            text = text.replace(src, dst)
    if "&#" in text:
        try:
            text = html.unescape(text)
        except Exception:
            pass
    return text


_WS_RE = re.compile(r"[\s\u00a0\u3000]+")


def _norm(text: str) -> str:
    """Collapse runs of whitespace; keeps text searchable and compact."""
    if not text:
        return ""
    return _WS_RE.sub(" ", text).strip()


def _decode(raw: bytes) -> str:
    """Best-effort text decoding for CSV/TSV/TXT."""
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", "replace")
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        for enc in ("utf-16", "utf-16-le", "utf-16-be"):
            try:
                return raw.decode(enc)
            except UnicodeError:
                continue
    for enc in ("utf-8", "gb18030", "big5", "shift_jis", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", "replace")


def _runs_binary(blob: bytes, min_len: int = 4) -> str:
    """Extract readable string runs from a binary blob (UTF-16LE first, then ASCII).

    Used as a best-effort reader for the legacy .xls / .xlsb containers, which
    have no public pure-Python parser. Catches the vast majority of cell text.
    """
    out: list[str] = []

    # UTF-16LE runs (what Excel writes internally)
    try:
        text16 = blob.decode("utf-16-le", "ignore")
    except Exception:
        text16 = ""
    if text16:
        for run in re.findall(r"[^\x00-\x08\x0b-\x1f\x7f]{%d,}" % min_len, text16):
            out.append(run)

    # ASCII / single-byte runs (compressed BIFF records, BIFF5)
    for run in re.findall(rb"[\x20-\x7e\x80-\xff]{%d,}" % min_len, blob):
        try:
            out.append(run.decode("latin-1"))
        except Exception:
            continue

    return _norm(" ".join(out))


def _shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Ordered shared-string table for an OOXML workbook (table[i] == index i)."""
    try:
        raw = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    except Exception:
        return []
    text = raw.decode("utf-8", "replace")
    if len(text) > 64 * 1024 * 1024:            # pathological file guard
        text = text[: 64 * 1024 * 1024]
    table: list[str] = []
    for chunk in RE_SI.findall(text):
        pieces = RE_T.findall(chunk)
        table.append(_unescape("".join(pieces)))
    return table


def _workbook_meta(zf: zipfile.ZipFile) -> list[str]:
    meta: list[str] = []
    try:
        blob = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    except Exception:
        return meta
    for m in RE_SHEETNAME.finditer(blob):
        name = m.group(1) or m.group(2)
        if name:
            meta.append(_unescape(name))
    for m in RE_DEFINED.finditer(blob):
        val = _norm(_unescape(re.sub(r"<[^>]+>", "", m.group(1))))
        if val:
            meta.append(val)
    return meta


# --------------------------------------------------------------------------
# Cell-level location (which sheet / which cell a keyword sits in)
# --------------------------------------------------------------------------

RE_WB_SHEET = re.compile(r"<" + _tag("sheet") + r"\b([^>]*)/?>")
RE_RID = re.compile(r'\br:id\s*=\s*"([^"]+)"')
RE_SHEET_NAME_ATTR = re.compile(r'\bname\s*=\s*"([^"]*)"')
RE_REL = re.compile(r"<" + _tag("Relationship") + r"\b([^>]*)/?>")
RE_ROW = re.compile(r"<" + _tag("row") + r"\b[^>]*>(.*?)</" + _tag("row") + r">", re.S)
RE_ONE_CELL = re.compile(
    r"<" + _tag("c") + r"\b([^>]*?)(?:/>|>(.*?)</" + _tag("c") + r">)", re.S
)


def _sheet_member_map(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map worksheet member path -> visible sheet name.

    An OOXML workbook reaches its sheets through a relationship file, so
    ``xl/worksheets/sheet1.xml`` is not necessarily the first tab. Reading the
    rels is what lets a result say ``明细!B7`` correctly instead of guessing.
    """
    try:
        wb = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    except Exception:
        return {}
    rid_to_name: dict[str, str] = {}
    order: list[str] = []
    for m in RE_WB_SHEET.finditer(wb):
        attrs = m.group(1)
        name = RE_SHEET_NAME_ATTR.search(attrs)
        rid = RE_RID.search(attrs)
        if not name:
            continue
        label = _unescape(name.group(1))
        order.append(label)
        if rid:
            rid_to_name[rid.group(1)] = label

    rid_to_target: dict[str, str] = {}
    try:
        rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    except Exception:
        rels = ""
    for m in RE_REL.finditer(rels):
        attrs = m.group(1)
        rid = RE_RID.search(attrs)
        tgt = re.search(r'\bTarget\s*=\s*"([^"]+)"', attrs)
        if rid and tgt:
            target = tgt.group(1).lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("./")
            rid_to_target[rid.group(1)] = target

    mapping: dict[str, str] = {}
    for rid, name in rid_to_name.items():
        target = rid_to_target.get(rid)
        if target:
            mapping[target] = name

    # Fall back to positional naming when the rels are missing or unusual.
    members = sorted(
        n for n in zf.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")
    )
    for i, member in enumerate(members):
        if member not in mapping and i < len(order):
            mapping[member] = order[i]
    return mapping


def _cell_text(attrs: str, body: str | None, table: list[str]) -> str:
    """Resolve one ``<c>`` element to its display text ("" when empty)."""
    if not body:
        return ""
    m = RE_CELL_TYPE.search(attrs)
    ctype = m.group(1) if m else "n"
    if ctype == "s":
        mv = RE_V.search(body)
        if not mv:
            return ""
        try:
            return table[int(mv.group(1))]
        except (ValueError, IndexError):
            return ""
    if ctype == "inlineStr":
        mis = RE_IS.search(body)
        if not mis:
            return ""
        return _unescape("".join(RE_T.findall(mis.group(1))))
    mv = RE_V.search(body)
    return _unescape(mv.group(1)) if mv else ""


def iter_cells(path: str):
    """Yield ``(sheet_name, cell_ref, value)`` for every non-empty cell.

    Streaming and bounded: worksheet XML is scanned row by row so a huge sheet
    never has to be held in memory all at once.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".ods"):
        raise ValueError(f"该格式不支持单元格级定位：{ext or '(无扩展名)'}")

    with zipfile.ZipFile(path) as zf:
        table = _shared_strings(zf)
        mapping = _sheet_member_map(zf)
        members = sorted(
            n for n in zf.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml")
        )
        for member in members:
            sheet_name = mapping.get(member) or os.path.basename(member)
            try:
                blob = zf.read(member).decode("utf-8", "replace")
            except Exception:
                continue
            for row_xml in RE_ROW.findall(blob):
                for attrs, body in RE_ONE_CELL.findall(row_xml):
                    ref = RE_CELL_REF.search(attrs)
                    if not ref:
                        continue
                    value = _norm(_cell_text(attrs, body, table))
                    if value:
                        yield sheet_name, ref.group(1).upper(), value


def find_cells(
    path: str,
    keyword: str,
    case: bool = False,
    limit: int = 50,
    want_context: bool = True,
    context_width: int = 160,
) -> list[dict]:
    """Locate `keyword` inside a single workbook, cell by cell.

    Returns a list of dicts::

        {"sheet", "ref", "value", "row", "col", "row_text", "match"}

    This is the second pass that turns "this file matches" into
    "Sheet「明细」B7 == 张伟", which is what makes a result actionable.
    """
    hits: list[dict] = []
    if not keyword:
        return hits
    needle = keyword if case else keyword.casefold()

    for sheet, ref, value in iter_cells(path):
        hay = value if case else value.casefold()
        if needle not in hay:
            continue
        pos = hay.find(needle)
        digits = "".join(ch for ch in ref if ch.isdigit())
        row_no = int(digits) if digits else 0

        # The whole row is the useful citation: it shows the person's other
        # columns, which is how you tell two "张伟" records apart.
        row_text = _row_context_fast(path, sheet, row_no, context_width) if want_context else ""
        hits.append(
            {
                "sheet": sheet,
                "ref": ref,
                "value": value[:200],
                "row": row_no,
                "col": "".join(ch for ch in ref if ch.isalpha()),
                "row_text": row_text,
                "match": value[max(0, pos - 40): pos + len(keyword) + 40],
            }
        )
        if len(hits) >= limit:
            break
    return hits


def _row_context_fast(path: str, sheet: str, row_no: int, width: int) -> str:
    """Single-row text for context, read directly from the worksheet XML.

    The row number comes from the cell refs (``D16`` -> 16) rather than the
    ``<row r=...>`` attribute, which not every writer emits -- Excel for Windows
    does, several libraries do not.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".ods"):
        return ""
    try:
        with zipfile.ZipFile(path) as zf:
            table = _shared_strings(zf)
            mapping = _sheet_member_map(zf)
            member = next((m for m, n in mapping.items() if n == sheet), None)
            if member is None:
                return ""
            blob = zf.read(member).decode("utf-8", "replace")
            for row_xml in RE_ROW.findall(blob):
                values: list[str] = []
                matched = False
                for attrs, body in RE_ONE_CELL.findall(row_xml):
                    ref = RE_CELL_REF.search(attrs)
                    if ref:
                        digits = "".join(ch for ch in ref.group(1) if ch.isdigit())
                        if digits and int(digits) == row_no:
                            matched = True
                    text = _norm(_cell_text(attrs, body, table))
                    if text:
                        values.append(text)
                if matched:
                    return " | ".join(values)[:width]
    except Exception:
        return ""
    return ""


def _extract_xlsx(path: str) -> str:
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        table = _shared_strings(zf)
        buf: list[str] = []
        size = 0

        for meta in _workbook_meta(zf):
            buf.append(meta)
            size += len(meta)

        sheets = [n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml")]
        sheets.sort()
        if not sheets:
            sheets = [n for n in names if n.endswith(".xml") and "worksheet" in n]

        for member in sheets:
            try:
                blob = zf.read(member).decode("utf-8", "replace")
            except Exception:
                continue
            sheet_buf: list[str] = []
            seen: set[str] = set()
            for attrs, body in RE_CELL.findall(blob):
                if not body:
                    continue
                m = RE_CELL_TYPE.search(attrs)
                ctype = m.group(1) if m else "n"
                if ctype == "s":
                    mv = RE_V.search(body)
                    if not mv:
                        continue
                    try:
                        val = table[int(mv.group(1))]
                    except (ValueError, IndexError):
                        continue
                elif ctype == "inlineStr":
                    mis = RE_IS.search(body)
                    if not mis:
                        continue
                    val = _unescape("".join(RE_T.findall(mis.group(1))))
                else:
                    mv = RE_V.search(body)
                    if not mv:
                        continue
                    val = _unescape(mv.group(1))
                val = _norm(val)
                if not val:
                    continue
                # A shared string repeated across 10k rows adds no search value.
                if val in seen:
                    continue
                seen.add(val)
                sheet_buf.append(val)
                size += len(val) + 1
                if size > MAX_TEXT_CHARS:
                    break
            if sheet_buf:
                buf.append(" ".join(sheet_buf))
            if size > MAX_TEXT_CHARS:
                break

        return _norm(" ".join(buf))


def extract_text(path: str, size: int | None = None) -> str:
    """Extract searchable plain text from a spreadsheet file.

    Raises on hard failures so the caller can record the reason.
    """
    ext = os.path.splitext(path)[1].lower()
    if size is None:
        try:
            size = os.path.getsize(path)
        except OSError:
            size = 0

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".ods"):
        return _extract_xlsx(path)

    if ext in (".csv", ".tsv", ".txt", ".md", ".log"):
        with open(path, "rb") as fh:
            raw = fh.read(MAX_TEXT_CHARS * 4)
        return _norm(_decode(raw))[:MAX_TEXT_CHARS]

    if ext in (".xls", ".xlsb", ".xlt", ".xla"):
        if size > MAX_CONTENT_FILESIZE:
            raise ValueError("文件过大，跳过内容解析")
        with open(path, "rb") as fh:
            blob = fh.read(MAX_CONTENT_FILESIZE)
        return _runs_binary(blob)[:MAX_TEXT_CHARS]

    raise ValueError(f"不支持内容解析的格式：{ext or '(无扩展名)'}")


# --------------------------------------------------------------------------
# Index records
# --------------------------------------------------------------------------

@dataclass(slots=True)
class Record:
    path: str
    name: str
    size: int
    mtime: float
    ext: str
    content: str | None = None      # None = not extracted yet
    error: str = ""
    content_loaded: bool = False

    @property
    def folder(self) -> str:
        return os.path.dirname(self.path)

    def folded(self) -> str:
        """Case-insensitive form of `content`, computed once and reused.

        Case-folding megabytes of text on every keystroke is the single most
        expensive part of a content search, so it is cached and invalidated
        whenever `content` is replaced.
        """
        cached = self._folded
        if cached is not None and self._folded_for is self.content:
            return cached
        self._folded = (self.content or "").casefold()
        self._folded_for = self.content
        return self._folded

    def set_content(self, text: str | None, error: str = "") -> None:
        self.content = text
        self.error = error
        self.content_loaded = True
        self._folded = None
        self._folded_for = None

    _folded: str | None = field(default=None, repr=False, compare=False)
    _folded_for: str | None = field(default=None, repr=False, compare=False)

    def to_json(self) -> list:
        return [self.path, self.size, self.mtime, self.error, self.content]


# --------------------------------------------------------------------------
# Matcher
# --------------------------------------------------------------------------

@dataclass(slots=True)
class MatchOpts:
    query: str = ""
    mode: str = "contains"      # contains | prefix | exact | wildcard | regex
    case: bool = False
    ext_filter: str = ""        # e.g. ".xlsx,.xlsm"; empty = all indexed
    min_size: int = 0           # bytes
    mtime_after: float = 0.0
    mtime_before: float = 0.0
    req_name: bool = False      # keyword must appear in the FILE NAME
    req_content: bool = False   # keyword must appear in the CONTENT
    exclude: str = ""           # tokens that must not appear in the path


class Matcher:
    """Compiles a user query into a fast callable.

    The query is split on spaces into tokens which are AND-ed together, so
    "2023 销售" finds names containing both parts -- closer to how people
    actually search than a single literal string.
    """

    def __init__(self, opts: MatchOpts):
        self.opts = opts
        flags = 0 if opts.case else re.IGNORECASE
        raw = (opts.query or "").strip()
        self.tokens = [t for t in re.split(r"\s+", raw) if t] if raw else []
        self.needles = [t if opts.case else t.casefold() for t in self.tokens]
        self.regexes: list[re.Pattern[str]] = []
        self.excludes: list[str] = [
            t if opts.case else t.casefold() for t in re.split(r"\s+", (opts.exclude or "").strip()) if t
        ]

        for token in self.tokens:
            if opts.mode == "contains":
                continue
            if opts.mode == "prefix":
                self.regexes.append(re.compile(r"^" + re.escape(token), flags))
            elif opts.mode == "exact":
                self.regexes.append(re.compile(r"^" + re.escape(token) + r"$", flags))
            elif opts.mode == "wildcard":
                pattern = re.escape(token).replace(r"\*", ".*").replace(r"\?", ".")
                self.regexes.append(re.compile(pattern, flags))
            elif opts.mode == "regex":
                try:
                    self.regexes.append(re.compile(token, flags))
                except re.error:
                    # Invalid pattern: fall back to a literal match rather than
                    # silently returning nothing.
                    self.regexes.append(re.compile(re.escape(token), flags))

        self.exts: set[str] = set()
        for part in (opts.ext_filter or "").replace(";", ",").split(","):
            part = part.strip().lower()
            if not part:
                continue
            if not part.startswith("."):
                part = "." + part
            self.exts.add(part)

        self.use_regex = opts.mode != "contains"
        self.has_filters = bool(
            self.exts or opts.min_size or opts.mtime_after or opts.mtime_before or raw
        )

    # -- name matching ----------------------------------------------------
    def match_name(self, name: str) -> bool:
        if not self.tokens:
            return True
        hay = name if self.opts.case else name.casefold()
        if self.use_regex:
            return all(rx.search(name) for rx in self.regexes)
        return all(n in hay for n in self.needles)

    # -- content matching -------------------------------------------------
    def match_content(self, content: str) -> bool:
        if not self.tokens:
            return True
        hay = content if self.opts.case else content.casefold()
        if self.use_regex:
            return all(rx.search(content) for rx in self.regexes)
        return all(n in hay for n in self.needles)

    def match_content_folded(self, folded: str, raw: str | None = None) -> bool:
        """Match against an already case-folded string (avoids re-folding)."""
        if not self.tokens:
            return True
        if self.use_regex:
            return all(rx.search(raw if raw is not None else folded) for rx in self.regexes)
        return all(n in folded for n in self.needles)

    def estimate_cost(self) -> int:
        """Rough work estimate used to decide whether a pre-pass pays off."""
        if not self.tokens:
            return 0
        if self.use_regex:
            return 3000 * len(self.tokens)
        return sum(len(n) for n in self.needles) * 20

    def match_meta(self, rec: Record) -> bool:
        o = self.opts
        if self.exts and rec.ext not in self.exts:
            return False
        if o.min_size and rec.size < o.min_size:
            return False
        if o.mtime_after and rec.mtime < o.mtime_after:
            return False
        if o.mtime_before and rec.mtime > o.mtime_before:
            return False
        if self.excludes:
            hay = rec.path if o.case else rec.path.casefold()
            if any(x in hay for x in self.excludes):
                return False
        return True

    def scope(self) -> str:
        if self.opts.req_name and self.opts.req_content:
            return "both"
        if self.opts.req_content:
            return "content"
        if self.opts.req_name:
            return "name"
        # No explicit scope: match name first, content only when asked for.
        return "name"


# --------------------------------------------------------------------------
# Scoring / result ranking
# --------------------------------------------------------------------------

def name_score(name: str, matcher: Matcher) -> int:
    """Higher is better. Ranked: equals > starts with > word boundary > anywhere."""
    if not matcher.tokens:
        return 10
    stem = os.path.splitext(name)[0]
    a = name if matcher.opts.case else name.casefold()
    b = stem if matcher.opts.case else stem.casefold()
    score = 0
    for needle in matcher.needles:
        if not needle:
            continue
        if b == needle:
            score += 400
        elif a == needle:
            score += 350
        elif b.startswith(needle):
            score += 260
        elif a.startswith(needle):
            score += 230
        elif needle in b:
            score += 160
        elif needle in a:
            score += 120
        else:
            score += 40
    return score


def make_preview(text: str, needle: str, width: int = PREVIEW_CHARS) -> str:
    """Short '...keyword...' context window around the first hit."""
    if not text:
        return ""
    if not needle:
        return text[:width]
    pos = text.casefold().find(needle.casefold())
    if pos < 0:
        return text[:width]
    start = max(0, pos - width // 3)
    end = min(len(text), pos + len(needle) + width)
    return ("..." if start else "") + text[start:end].replace("\n", " ") + ("..." if end < len(text) else "")


# --------------------------------------------------------------------------
# Index container
# --------------------------------------------------------------------------

class Index:
    """In-memory index. All queries are pure RAM operations (microseconds)."""

    def __init__(self) -> None:
        self.records: list[Record] = []
        self.root: str = ""
        self.built_at: float = 0.0
        self.scan_seconds: float = 0.0
        self.content_done: int = 0
        self._by_path: dict[str, Record] = {}
        self._lock = threading.Lock()

    # -- population -------------------------------------------------------
    def add(self, rec: Record) -> None:
        self.records.append(rec)
        self._by_path[rec.path] = rec

    def replace_all(self, records: Iterable[Record], root: str, seconds: float) -> None:
        self.records = list(records)
        self._by_path = {r.path: r for r in self.records}
        self.root = root
        self.built_at = time.time()
        self.scan_seconds = seconds

    def get(self, path: str) -> Record | None:
        return self._by_path.get(path)

    # -- stats ------------------------------------------------------------
    def stats(self) -> dict:
        by_ext: dict[str, int] = {}
        total = 0
        for r in self.records:
            by_ext[r.ext] = by_ext.get(r.ext, 0) + 1
            total += r.size
        with_content = sum(1 for r in self.records if r.content_loaded)
        return {
            "count": len(self.records),
            "bytes": total,
            "by_ext": dict(sorted(by_ext.items(), key=lambda kv: -kv[1])),
            "with_content": with_content,
            "root": self.root,
            "built_at": self.built_at,
            "scan_seconds": self.scan_seconds,
        }

    # -- query ------------------------------------------------------------
    def search(self, matcher: Matcher, limit: int = 20000) -> tuple[list[tuple[Record, int, str]], int]:
        """Return ([(record, score, preview)], total_matches).

        `total_matches` can exceed the returned list when `limit` truncates it.
        """
        scope = matcher.scope()
        needle = matcher.tokens[0] if matcher.tokens else ""
        has_tokens = bool(matcher.tokens)
        hits: list[tuple[Record, int, str]] = []
        total = 0

        # For content-heavy queries, walking the records that have already been
        # stemmed is cheaper than scanning the whole index and bailing out on a
        # missing content index for each one.
        pool: Iterable[Record] = self.records
        if scope in ("content", "both") and matcher.estimate_cost() >= 60:
            pool = [r for r in self.records if r.content_loaded]

        for rec in pool:
            if not matcher.match_meta(rec):
                continue

            if scope == "content":
                if not rec.content_loaded:
                    continue
                folded = rec.folded()
                if has_tokens and not matcher.match_content_folded(folded, rec.content):
                    continue
                # Only the rows that will actually be shown need a snippet --
                # building context windows for 100k matches is wasted work.
                preview = make_preview(rec.content or "", needle) if len(hits) < limit else ""
            else:  # "name" or "both"
                if not matcher.match_name(rec.name):
                    continue
                preview = ""
                if scope == "both":
                    if not rec.content_loaded:
                        continue
                    folded = rec.folded()
                    if has_tokens and not matcher.match_content_folded(folded, rec.content):
                        continue
                    preview = make_preview(rec.content or "", needle) if len(hits) < limit else ""

            if not has_tokens:
                score = 10
            else:
                score = name_score(rec.name, matcher)
                if scope == "content":
                    score += 500          # content hits are rarer => more valuable
            total += 1
            if len(hits) < limit:
                hits.append((rec, score, preview))

        hits.sort(key=lambda t: (-t[1], -t[0].mtime, t[0].name.casefold()))
        return hits, total

    def paths_needing_content(self, matcher: Matcher | None = None) -> list[Record]:
        out = []
        for rec in self.records:
            if rec.content_loaded:
                continue
            if rec.ext not in CONTENT_EXTS:
                continue
            if matcher is not None and not matcher.match_meta(rec):
                continue
            out.append(rec)
        return out

    # -- persistence ------------------------------------------------------
    def save(self, path: str) -> None:
        payload = {
            "app": APP_NAME,
            "format": INDEX_FORMAT,
            "version": APP_VERSION,
            "root": self.root,
            "built_at": self.built_at,
            "scan_seconds": self.scan_seconds,
            "records": [r.to_json() for r in self.records],
        }
        tmp = path + ".tmp"
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)
        os.replace(tmp, path)

    def load(self, path: str) -> bool:
        try:
            with open(path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
        except (OSError, ValueError):
            return False
        if payload.get("format") != INDEX_FORMAT or payload.get("app") != APP_NAME:
            return False
        records: list[Record] = []
        for row in payload.get("records", []):
            try:
                p, size, mtime, error, content = (list(row) + [None] * 5)[:5]
            except (TypeError, ValueError):
                continue
            rec = Record(
                path=p,
                name=os.path.basename(p),
                size=int(size or 0),
                mtime=float(mtime or 0),
                ext=os.path.splitext(p)[1].lower(),
                content=content,
                error=error or "",
                content_loaded=content is not None,
            )
            records.append(rec)
        self.records = records
        self._by_path = {r.path: r for r in records}
        self.root = payload.get("root", "")
        self.built_at = float(payload.get("built_at", 0))
        self.scan_seconds = float(payload.get("scan_seconds", 0))
        return True


# --------------------------------------------------------------------------
# Scan engine
# --------------------------------------------------------------------------

class ScanEngine:
    """Parallel directory walker.

    A single `os.walk` on a large tree is bottlenecked on one core and one
    syscall stream. Here the top level is partitioned into independent
    sub-trees, each walked by its own thread, so an NVMe SSD or a network share
    is saturated instead of the CPU waiting.

    `reuse` enables incremental scans: cached records whose (size, mtime) still
    match are carried over untouched, so a re-scan of a huge tree is fast.
    """

    def __init__(
        self,
        root: str,
        threads: int = DEFAULT_THREADS,
        on_file: Callable[[Record, int], None] | None = None,
        skip_hidden: bool = False,
        follow_links: bool = False,
        reuse: Iterable[Record] | None = None,
    ) -> None:
        self.root = os.path.abspath(root)
        self.threads = max(1, int(threads))
        self.on_file = on_file
        self.skip_hidden = skip_hidden
        self.follow_links = follow_links
        self._cancel = threading.Event()
        self._seen = 0
        self._lock = threading.Lock()
        self.reused = 0
        self.errors: list[str] = []

        self._cache: dict[str, Record] = {}
        for rec in reuse or ():
            self._cache[rec.path] = rec

    def cancel(self) -> None:
        self._cancel.set()

    @property
    def cancelled(self) -> bool:
        return self._cancel.is_set()

    def count(self) -> int:
        with self._lock:
            return self._seen

    # ------------------------------------------------------------------
    def _walk(self, top: str, out: list[Record]) -> None:
        stack = [top]
        skip = SKIP_DIR_NAMES
        skip_hidden = self.skip_hidden
        follow = self.follow_links
        cache = self._cache
        on_file = self.on_file
        while stack:
            if self._cancel.is_set():
                return
            current = stack.pop()
            try:
                with os.scandir(current) as it:
                    for entry in it:
                        try:
                            if entry.is_dir(follow_symlinks=follow):
                                if skip_hidden and entry.name.startswith("."):
                                    continue
                                if entry.name.casefold() in skip:
                                    continue
                                stack.append(entry.path)
                                continue
                            if not entry.is_file(follow_symlinks=follow):
                                continue
                        except OSError:
                            continue

                        ext = os.path.splitext(entry.name)[1].lower()
                        if ext not in ACCEPT_EXTS:
                            continue
                        if skip_hidden and entry.name.startswith(("~$", ".")):
                            continue

                        try:
                            st = entry.stat(follow_symlinks=follow)
                            size, mtime = st.st_size, st.st_mtime
                        except OSError as exc:
                            self._record_error(f"{entry.path}: {exc}")
                            continue

                        with self._lock:
                            self._seen += 1
                            n = self._seen

                        old = cache.get(entry.path)
                        if old is not None and old.size == size and old.mtime == mtime:
                            rec = old
                            self.reused += 1
                            if ext in EXCEL_EXTS:
                                out.append(rec)
                            # A callback per file costs more than the scan itself
                            # on 100k-file trees, so report every 64th entry.
                            if on_file is not None and not (n & 63):
                                on_file(rec, n)
                            continue

                        rec = Record(path=entry.path, name=entry.name, size=size, mtime=mtime, ext=ext)
                        if ext in EXCEL_EXTS:
                            out.append(rec)
                        if on_file is not None and not (n & 63):
                            on_file(rec, n)
            except OSError as exc:
                self._record_error(f"{current}: {exc}")
            except Exception as exc:                      # pragma: no cover - defensive
                self._record_error(f"{current}: {exc!r}")

    def _record_error(self, msg: str) -> None:
        with self._lock:
            if len(self.errors) < 200:
                self.errors.append(msg)

    # ------------------------------------------------------------------
    def run(self) -> tuple[list[Record], float, int]:
        started = time.perf_counter()
        out: list[Record] = []
        root = self.root

        # Partition: root-level files + each top-level directory becomes a job.
        jobs: list[str] = []
        try:
            with os.scandir(root) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=self.follow_links):
                            if self.skip_hidden and entry.name.startswith("."):
                                continue
                            if entry.name.casefold() in SKIP_DIR_NAMES:
                                continue
                            jobs.append(entry.path)
                        elif entry.is_file(follow_symlinks=self.follow_links):
                            jobs.append(entry.path)
                    except OSError:
                        continue
        except OSError as exc:
            raise FileNotFoundError(f"无法读取目录：{root}（{exc}）") from exc

        if not jobs:
            return out, time.perf_counter() - started, 0

        # Small trees: one thread. Big trees: one thread per top-level branch.
        workers = max(1, min(self.threads, len(jobs)))
        if workers == 1:
            for job in jobs:
                if self._cancel.is_set():
                    break
                if os.path.isdir(job):
                    self._walk(job, out)
                else:
                    self._walk_file(job, out)
        else:
            with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="scan") as pool:
                futures = []
                for job in jobs:
                    if os.path.isdir(job):
                        futures.append(pool.submit(self._walk, job, out))
                    else:
                        futures.append(pool.submit(self._walk_file, job, out))
                for fut in futures:
                    try:
                        fut.result()
                    except Exception as exc:              # pragma: no cover - defensive
                        self._record_error(repr(exc))

        out.sort(key=lambda r: r.path.casefold())
        return out, time.perf_counter() - started, self.reused

    def _walk_file(self, path: str, out: list[Record]) -> None:
        name = os.path.basename(path)
        ext = os.path.splitext(name)[1].lower()
        if ext not in ACCEPT_EXTS:
            return
        try:
            st = os.stat(path)
        except OSError as exc:
            self._record_error(f"{path}: {exc}")
            return
        with self._lock:
            self._seen += 1
            n = self._seen
        old = self._cache.get(path)
        if old is not None and old.size == st.st_size and old.mtime == st.st_mtime:
            rec = old
            self.reused += 1
        else:
            rec = Record(path=path, name=name, size=st.st_size, mtime=st.st_mtime, ext=ext)
        if ext in EXCEL_EXTS:
            out.append(rec)
        if self.on_file is not None and not (n & 63):
            self.on_file(rec, n)


# --------------------------------------------------------------------------
# Content indexing
# --------------------------------------------------------------------------

def index_content(
    records: Sequence[Record],
    threads: int = DEFAULT_THREADS,
    on_progress: Callable[[int, int, Record], None] | None = None,
    cancel: threading.Event | None = None,
    force: bool = False,
) -> tuple[int, int]:
    """Extract text for records lacking it. Returns (done, failed)."""
    todo = [r for r in records if force or not r.content_loaded]
    if not todo:
        return 0, 0

    cancel = cancel or threading.Event()
    total = len(todo)
    done = 0
    failed = 0
    lock = threading.Lock()

    def work(rec: Record) -> None:
        nonlocal done, failed
        if cancel.is_set():
            return
        if rec.size > MAX_CONTENT_FILESIZE:
            rec.set_content("", "文件过大，跳过内容解析")
            with lock:
                failed += 1
                done += 1
                if on_progress:
                    on_progress(done, total, rec)
            return
        try:
            rec.set_content(extract_text(rec.path, rec.size))
        except Exception as exc:
            rec.set_content("", f"{type(exc).__name__}: {exc}")
            with lock:
                failed += 1
        with lock:
            done += 1
            if on_progress:
                on_progress(done, total, rec)

    workers = max(1, min(threads, len(todo)))
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="content") as pool:
        list(pool.map(work, todo))

    return done, failed


# --------------------------------------------------------------------------
# Export helpers
# --------------------------------------------------------------------------

def export_csv(
    path: str,
    rows: Sequence[tuple[Record, int, str]],
    matcher: Matcher | None = None,
    with_cells: bool = False,
    case: bool = False,
) -> None:
    """Write results to CSV, optionally resolving the matching cell address.

    `with_cells` re-reads each workbook, so it is opt-in: a 3000-row export
    would otherwise spend ten seconds hunting cell references nobody reads.
    """
    import csv

    keyword = matcher.tokens[0] if (matcher and matcher.tokens) else ""
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["文件名", "所在目录", "扩展名", "大小(字节)", "修改时间",
                         "命中位置", "工作表", "单元格", "该行内容", "命中片段"])
        for rec, _score, preview in rows:
            where = "文件名" if matcher is None or matcher.match_name(rec.name) else ""
            content_hit = bool(
                matcher is not None and rec.content_loaded
                and matcher.tokens and matcher.match_content_folded(rec.folded(), rec.content)
            )
            if content_hit:
                where = (where + "+内容") if where else "内容"

            sheets: list[str] = []
            refs: list[str] = []
            row_texts: list[str] = []
            if with_cells and keyword and content_hit:
                try:
                    located = find_cells(rec.path, keyword, case=case, limit=20)
                except Exception:
                    located = []
                for c in located:
                    sheets.append(c["sheet"])
                    refs.append(c["ref"])
                    row_texts.append(c.get("row_text", ""))

            writer.writerow(
                [
                    rec.name,
                    rec.folder,
                    rec.ext,
                    rec.size,
                    fmt_time(rec.mtime),
                    where,
                    " / ".join(dict.fromkeys(sheets))[:200],
                    " ".join(dict.fromkeys(refs))[:200],
                    (row_texts[0] if row_texts else "")[:300],
                    preview.replace("\n", " ")[:200],
                ]
            )


def find_root_via_dialog() -> str:  # pragma: no cover - GUI helper
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        chosen = filedialog.askdirectory(title="选择要搜索的目录")
        root.destroy()
        return chosen or ""
    except Exception:
        return ""


def open_in_explorer(path: str, select: bool = False) -> None:
    """Reveal a file in Explorer (or open it with its default app)."""
    if not os.path.exists(path):
        parent = os.path.dirname(path)
        if os.path.isdir(parent):
            os.startfile(parent)  # noqa: S606
        return
    if select:
        import subprocess

        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        os.startfile(path)  # noqa: S606


def setup_console() -> None:
    """Make console output UTF-8 and legible on a Chinese Windows.

    Python picks its stdout encoding from the *ANSI* codepage (cp936 on a zh-CN
    box) while the console is often decoding as UTF-8 (chcp 65001), so Chinese
    turns into mojibake. Forcing UTF-8 on both sides fixes interactive use and
    piped output alike.
    """
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError, OSError):
            pass


#: PowerShell that opens a workbook and selects a cell. Written as a template so
#: Excel and WPS can both be tried by swapping the ProgID.
_EXCEL_AT_CELL_PS = r"""
$ErrorActionPreference = 'Stop'
$path = '__PATH__'
$sheet = '__SHEET__'
$cell = '__CELL__'
$excel = $null
$owned = $false
try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('__PROGID__') }
catch {
  try { $excel = New-Object -ComObject '__PROGID__'; $owned = $true } catch { exit 3 }
}
if ($null -eq $excel) { exit 3 }
$excel.Visible = $true
$wb = $null
foreach ($b in @($excel.Workbooks)) { if ($b.FullName -eq $path) { $wb = $b; break } }
if ($null -eq $wb) { $wb = $excel.Workbooks.Open($path) }
if ($sheet -ne '') {
  $ws = $null
  foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $sheet) { $ws = $s; break } }
  if ($null -eq $ws) { $ws = $wb.Worksheets.Item(1) }
} else { $ws = $wb.Worksheets.Item(1) }
$ws.Activate()
if ($cell -ne '') {
  $r = $ws.Range($cell)
  $r.Select() | Out-Null
  $excel.Goto($r) | Out-Null
} else { $ws.Range('A1').Select() | Out-Null }
$wb.Activate() | Out-Null
# Release the COM instance when we created it, otherwise an orphaned host keeps
# the workbook locked and lingers in Task Manager. Excel survives this (it keeps
# a separate process alive); WPS would close the document we just opened, so it
# is left to exit with its own window.
if ($owned -and __OWNED_QUIT__) { $excel.DisplayAlerts = $false; $excel.Quit() }
exit 0
"""


#: Candidate COM ProgIDs per spreadsheet program, in preference order.
#: WPS Office registers Ket.Application for its ET spreadsheets component;
#: some older builds used ET.Application, so that is probed too.
SPREADSHEET_APPS: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("excel", "Microsoft Excel", ("Excel.Application",)),
    ("wps", "WPS 表格", ("Ket.Application", "KET.Application", "ET.Application", "et.Application")),
)

#: GUI choices -> (label, id, app ids to try in order).
OPEN_WITH_CHOICES: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("自动（Office 优先，无则用 WPS）", "auto", ("excel", "wps")),
    ("只用 Microsoft Excel 打开", "excel", ("excel",)),
    ("只用 WPS 表格打开", "wps", ("wps",)),
)


_INSTALLED_CACHE: dict[str, str] | None = None


def _progid_registered(prog_id: str) -> bool:
    """True when a COM ProgID exists in HKCR, i.e. the program is installed."""
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID"):
            return True
    except OSError:
        return False


def installed_apps(refresh: bool = False) -> dict[str, str]:
    """Map app id -> the ProgID actually registered on this PC.

    Detected from the registry rather than by launching: launching is slow, and
    a sandboxed COM failure is indistinguishable from "not installed", which
    would report the wrong answer.
    """
    global _INSTALLED_CACHE
    if _INSTALLED_CACHE is not None and not refresh:
        return dict(_INSTALLED_CACHE)
    found: dict[str, str] = {}
    for app_id, _label, prog_ids in SPREADSHEET_APPS:
        for prog_id in prog_ids:
            if _progid_registered(prog_id):
                found[app_id] = prog_id
                break
    _INSTALLED_CACHE = found
    return dict(found)


def _app_label(app_id: str) -> str:
    return next((lb for aid, lb, _p in SPREADSHEET_APPS if aid == app_id), app_id)


def open_excel_at_cell(
    path: str,
    sheet: str = "",
    cell: str = "",
    timeout: int = 90,
    prefer: str = "auto",
) -> tuple[bool, str]:
    """Open a workbook with `sheet`/`cell` selected.

    `prefer` is one of the OPEN_WITH_CHOICES ids: "auto" tries Office then WPS,
    "excel"/"wps" pin a specific program.

    Returns ``(ok, description)`` naming the program that actually worked, so
    the UI can report it instead of silently doing something else.
    """
    import subprocess

    order: tuple[str, ...] = ("excel", "wps")
    for _label, choice_id, app_ids in OPEN_WITH_CHOICES:
        if choice_id == prefer:
            order = app_ids
            break

    exe = os.path.join(
        os.environ.get("SystemRoot", r"C:\Windows"),
        "System32", "WindowsPowerShell", "v1.0", "powershell.exe",
    )
    if not os.path.exists(exe):
        exe = "powershell"

    available = installed_apps()
    tried: list[str] = []
    for app_id in order:
        label = _app_label(app_id)
        prog_id = available.get(app_id)
        if not prog_id:
            tried.append(f"{label} 未安装")
            continue
        script = (
            _EXCEL_AT_CELL_PS
            .replace("__PROGID__", prog_id)
            .replace("__PATH__", path.replace("'", "''"))
            .replace("__SHEET__", (sheet or "").replace("'", "''"))
            .replace("__CELL__", (cell or "").replace("'", "''"))
            # Only Excel tolerates an immediate Quit() after opening; WPS would
            # close the document before the user ever sees it.
            .replace("__OWNED_QUIT__", "$true" if app_id == "excel" else "$false")
        )
        try:
            proc = subprocess.run(
                [exe, "-NoProfile", "-NonInteractive", "-STA",
                 "-WindowStyle", "Hidden", "-Command", script],
                capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=timeout,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            tried.append(f"{label}（{type(exc).__name__}）")
            continue
        if proc.returncode == 0:
            return True, f"已用 {label} 打开"
        tried.append(f"{label}（错误码 {proc.returncode}）")
    return False, "无法打开：" + "、".join(tried)


# --------------------------------------------------------------------------
# Self-test (headless) -- shared by the CLI tool and the frozen binary
# --------------------------------------------------------------------------

def run_selftest(argv: Sequence[str] | None = None) -> int:  # pragma: no cover
    import argparse

    parser = argparse.ArgumentParser(
        prog="ExcelFinder --selftest",
        description="Headless ExcelFinder engine check (no GUI).",
    )
    parser.add_argument("root", help="directory to scan")
    parser.add_argument("--query", default="", help="keyword to search for")
    parser.add_argument("--content", action="store_true", help="also parse and search file contents")
    parser.add_argument("--threads", type=int, default=DEFAULT_THREADS)
    parser.add_argument("--limit", type=int, default=20, help="max rows to print")
    parser.add_argument(
        "--cells",
        type=int,
        nargs="?",
        const=5,
        default=0,
        metavar="N",
        help="also locate each hit inside the workbook (sheet + cell address), up to N per file",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    if not os.path.isdir(args.root):
        print(f"错误：目录不存在：{args.root}")
        return 2

    print(f"{APP_NAME} v{APP_VERSION} 自检")
    print(f"Python    : {sys.version.split()[0]}")
    print(f"线程数    : {args.threads}")
    print("-" * 68)

    engine = ScanEngine(args.root, threads=args.threads)
    t0 = time.perf_counter()
    records, seconds, reused = engine.run()
    print(f"扫描      : {len(records):,} 个表格文件，用时 {seconds:.3f}s")
    total = sum(r.size for r in records)
    rate = (total / 1024 / 1024 / seconds) if seconds > 0 else 0.0
    print(f"数据量    : {fmt_size(total)}   吞吐 {rate:,.1f} MB/s")
    if engine.errors:
        print(f"跳过/错误 : {len(engine.errors)}（首个：{engine.errors[0][:90]}）")

    if not records:
        print("提示：该目录下没有找到 Excel/CSV 文件，请换一个目录再试。")
        return 1

    idx = Index()
    idx.replace_all(records, args.root, seconds)

    if args.content:
        t0 = time.perf_counter()
        done, failed = index_content(records, threads=args.threads)
        el = time.perf_counter() - t0
        print(f"内容解析  : {done:,} 成功 / {failed:,} 失败，用时 {el:.2f}s ({done/max(el,1e-9):,.0f} 文件/秒)")
        chars = sum(len(r.content or "") for r in records)
        print(f"文本总量  : {chars:,} 字符")
        if failed:
            for rec in records:
                if rec.error:
                    print(f"   ! {rec.name}: {rec.error}")
                    break

    query = args.query.strip()
    if not query:
        print("\n（未提供 --query，跳过检索演示）")
        return 0

    matcher = Matcher(MatchOpts(query=query, req_content=args.content))
    t0 = time.perf_counter()
    hits, total_hits = idx.search(matcher)
    ms = (time.perf_counter() - t0) * 1000
    scope = "文件名+内容" if args.content else "文件名"
    print(f"\n检索      : {query!r}（范围：{scope}）")
    print(f"结果      : {total_hits:,} 个命中，耗时 {ms:.2f} ms")
    for rec, score, preview in hits[: args.limit]:
        print(f"  [{score:>4}] {rec.name}")
        print(f"          {rec.folder}")
        if args.cells:
            # Second pass on just this file: turn "matches" into "B7 of 明细".
            try:
                located = find_cells(rec.path, query, case=False, limit=args.cells)
            except Exception as exc:
                located = []
                print(f"          定位失败: {exc}")
            if located:
                for c in located:
                    print(f"          → 工作表「{c['sheet']}」{c['ref']} = {c['value'][:60]}")
                    if c.get("row_text"):
                        print(f"            该行: {c['row_text'][:110]}")
            else:
                print("          → 文件名命中（内容中未找到该关键词）")
        elif preview:
            print(f"          内容: {preview[:110]}")
    if total_hits > args.limit:
        print(f"  … 其余 {total_hits - args.limit:,} 条已省略")
    print("\n自检完成：引擎工作正常。")
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(run_selftest())
