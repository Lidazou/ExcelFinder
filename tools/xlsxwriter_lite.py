# -*- coding: utf-8 -*-
"""Minimal standard-library XLSX writer.

Enough of the OOXML spec to produce real, Excel-openable workbooks for testing
ExcelFinder without pulling in openpyxl/xlsxwriter. Also handy as a smoke test:
if ExcelFinder can read these files, its parser is namespace-correct, because
this writer emits the same ``x:`` prefixes Excel for Windows uses.
"""
from __future__ import annotations

import re
import zipfile
from xml.sax.saxutils import escape

_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _t(value) -> str:
    return _ILLEGAL.sub("", str(value))


def _col(index: int) -> str:
    """0-based column index -> A, B, ... AA."""
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
{sheet_overrides}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def write_xlsx(path: str, sheets: list[tuple[str, list[list]]]) -> None:
    """Write ``sheets`` = [(sheet_name, [[cell, cell, ...], ...]), ...] to ``path``."""
    shared: list[str] = []
    shared_index: dict[str, int] = {}

    def sid(text: str) -> int:
        i = shared_index.get(text)
        if i is None:
            i = len(shared)
            shared.append(text)
            shared_index[text] = i
        return i

    sheet_xml: list[str] = []
    for _name, rows in sheets:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
            "<sheetData>",
        ]
        for r, row in enumerate(rows, start=1):
            parts.append(f'<row r="{r}">')
            for c, value in enumerate(row):
                ref = f"{_col(c)}{r}"
                if value is None or value == "":
                    continue
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    parts.append(f'<c r="{ref}"><v>{value}</v></c>')
                else:
                    # Mirrors Excel: strings live in sharedStrings.xml and the
                    # cell carries only an index with t="s".
                    parts.append(f'<c r="{ref}" t="s"><v>{sid(_t(value))}</v></c>')
            parts.append("</row>")
        parts.append("</sheetData></worksheet>")
        sheet_xml.append("".join(parts))

    overrides = "\n".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(sheets) + 1)
    )

    workbook = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    for i, (name, _rows) in enumerate(sheets, start=1):
        safe_name = escape(_t(name), {'"': "&quot;"})
        workbook.append(
            f'<sheet name="{safe_name}" sheetId="{i}" r:id="rId{i}"/>'
        )
    workbook.append("</sheets>")
    workbook.append(
        '<definedNames><definedName name="_xlnm.Print_Titles">明细!$1:$1</definedName></definedNames>'
    )
    workbook.append("</workbook>")

    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>',
    ]
    for i in range(1, len(sheets) + 1):
        wb_rels.append(
            f'<Relationship Id="rId{i}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
        )
    wb_rels.append("</Relationships>")

    strings_xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared)}" uniqueCount="{len(shared)}">',
    ]
    for s in shared:
        strings_xml.append(f"<si><t>{escape(s)}</t></si>")
    strings_xml.append("</sst>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES.format(sheet_overrides=overrides))
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("xl/workbook.xml", "".join(workbook))
        zf.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        zf.writestr("xl/sharedStrings.xml", "".join(strings_xml))
        for i, xml in enumerate(sheet_xml, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", xml)
