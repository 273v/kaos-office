"""Native lxml XLSX writer — TabularDocument to SpreadsheetML.

Uses the same OPC infrastructure and namespace constants as the native
reader. Zero external dependencies beyond lxml (already a core
kaos-office dependency). Consistent with the kaos-office philosophy:
full control over the XML layer, no format-library dependency.

The writer produces the exact inverse of what ``native.py`` reads:
- ``xl/workbook.xml`` — sheet definitions
- ``xl/worksheets/sheetN.xml`` — row/cell data with typed values
- ``xl/sharedStrings.xml`` — deduplicated string table
- ``xl/styles.xml`` — number formats for dates/times/money
- OPC packaging: ``[Content_Types].xml``, ``.rels`` files

Usage::

    from kaos_office.xlsx.writer import write_xlsx, write_xlsx_bytes

    write_xlsx(tabular_doc, "output.xlsx")
    xlsx_bytes = write_xlsx_bytes(tabular_doc)
"""

from __future__ import annotations

import datetime
import zipfile
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Any

from kaos_core.logging import get_logger
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import (
    SML,
    SML_CELL,
    SML_MERGE_CELL,
    SML_MERGE_CELLS,
    SML_ROW,
    SML_SHEET,
    SML_SHEET_DATA,
    SML_SHEETS,
    SML_SI,
    SML_SST,
    SML_T,
    SML_VALUE,
    SML_WORKBOOK,
    SML_WORKSHEET,
    R,
)
from kaos_office.xlsx.cell_ref import index_to_col_letters
from kaos_office.xlsx.styles import date_to_serial

logger = get_logger(__name__)

# OPC content types
_CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_CT_SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
_CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
_CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"

# Relationship types
_RT_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
_RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_RT_SHARED_STRINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)
_RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

# Namespace maps for serialization
_SML_NSMAP = {None: SML}
_R_NSMAP = {None: SML, "r": R}
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_RELS_NSMAP = {None: _RELS_NS}
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_CT_NSMAP = {None: _CT_NS}

# Built-in date numFmtIds
_NUMFMT_DATE = 14  # mm-dd-yy
_NUMFMT_DATETIME = 22  # m/d/yy h:mm
_NUMFMT_TIME = 21  # h:mm:ss
_NUMFMT_DURATION = 46  # [h]:mm:ss
_NUMFMT_MONEY = 164  # custom: $#,##0.00
_NUMFMT_DECIMAL = 165  # custom: #,##0.00


def write_xlsx(
    doc: Any,
    path: str | Path,
) -> Path:
    """Write a TabularDocument to an XLSX file using native lxml.

    Args:
        doc: A ``TabularDocument`` from kaos-content.
        path: Output file path.

    Returns:
        The output path.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    data = write_xlsx_bytes(doc)
    path.write_bytes(data)

    logger.info(
        "xlsx.writer: wrote %s, tables=%d, size=%d, path=%s",
        doc.metadata.title or "untitled",
        len(doc.tables),
        len(data),
        path,
    )
    return path


def write_xlsx_bytes(doc: Any) -> bytes:
    """Write a TabularDocument to XLSX bytes (in-memory).

    Produces a standards-compliant XLSX ZIP archive with:
    - Workbook + worksheets with typed cells
    - Shared string table for text deduplication
    - Styles for date/time/money/decimal formatting
    - Proper OPC packaging (content types, relationships)
    """
    from kaos_content.model.tabular import ColumnType

    buf = BytesIO()

    # Collect all unique strings across all tables for the SST
    sst: dict[str, int] = {}  # string → index

    def _intern_string(s: str) -> int:
        if s not in sst:
            sst[s] = len(sst)
        return sst[s]

    # Pre-scan all tables to build the SST
    for table in doc.tables:
        for col in table.columns:
            _intern_string(col.name)
        for row in table.rows:
            for col_idx, value in enumerate(row):
                col_type = (
                    table.columns[col_idx].column_type
                    if col_idx < len(table.columns)
                    else ColumnType.TEXT
                )
                if _is_string_type(value, col_type):
                    _intern_string(_to_string(value))

    # Build style index: map ColumnType → style index (s attribute)
    # Style 0 = General, Style 1-N = typed formats
    style_map = {
        ColumnType.DATE: 1,
        ColumnType.DATETIME: 2,
        ColumnType.TIME: 3,
        ColumnType.DURATION: 4,
        ColumnType.MONEY: 5,
        ColumnType.DECIMAL: 6,
    }

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1. Content types
        zf.writestr("[Content_Types].xml", _build_content_types(doc.tables))

        # 2. Root relationships
        zf.writestr("_rels/.rels", _build_root_rels())

        # 3. Workbook
        zf.writestr("xl/workbook.xml", _build_workbook(doc.tables))

        # 4. Workbook relationships
        zf.writestr("xl/_rels/workbook.xml.rels", _build_workbook_rels(doc.tables))

        # 5. Shared strings
        zf.writestr("xl/sharedStrings.xml", _build_shared_strings(sst))

        # 6. Styles
        zf.writestr("xl/styles.xml", _build_styles())

        # 7. Worksheets
        for i, table in enumerate(doc.tables):
            sheet_xml = _build_worksheet(table, sst, style_map)
            zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet_xml)

    return buf.getvalue()


def _is_string_type(value: Any, col_type: Any) -> bool:
    """Check if a value should be stored as a shared string."""
    from kaos_content.model.tabular import ColumnType

    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float, Decimal)):
        return False
    if isinstance(value, (datetime.date, datetime.datetime, datetime.time, datetime.timedelta)):
        return False
    return not (col_type in (ColumnType.MONEY,) and isinstance(value, dict))


def _to_string(value: Any) -> str:
    """Convert a value to its display string for the SST.

    Lists are formatted as semicolon-separated values (not JSON).
    Dicts are formatted as "key: value" pairs.
    """
    if isinstance(value, list):
        # Semicolon-separated, no brackets — readable in spreadsheets
        return "; ".join(str(v) for v in value) if value else ""
    if isinstance(value, dict):
        return "; ".join(f"{k}: {v}" for k, v in value.items())
    return str(value)


def _xml_decl(root: etree._Element) -> bytes:
    """Serialize an lxml element to bytes with XML declaration."""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _build_content_types(tables: tuple) -> bytes:
    """Build [Content_Types].xml."""
    root = etree.Element(f"{{{_CT_NS}}}Types", nsmap=_CT_NSMAP)

    etree.SubElement(root, "Default", Extension="rels", ContentType=_CT_RELS)
    etree.SubElement(root, "Default", Extension="xml", ContentType="application/xml")
    etree.SubElement(root, "Override", PartName="/xl/workbook.xml", ContentType=_CT_WORKBOOK)
    etree.SubElement(
        root,
        "Override",
        PartName="/xl/sharedStrings.xml",
        ContentType=_CT_SHARED_STRINGS,
    )
    etree.SubElement(root, "Override", PartName="/xl/styles.xml", ContentType=_CT_STYLES)

    for i in range(len(tables)):
        etree.SubElement(
            root,
            "Override",
            PartName=f"/xl/worksheets/sheet{i + 1}.xml",
            ContentType=_CT_WORKSHEET,
        )

    return _xml_decl(root)


def _build_root_rels() -> bytes:
    """Build _rels/.rels — root relationships."""
    root = etree.Element("Relationships", nsmap=_RELS_NSMAP)
    etree.SubElement(
        root,
        "Relationship",
        Id="rId1",
        Type=_RT_OFFICE_DOCUMENT,
        Target="xl/workbook.xml",
    )
    return _xml_decl(root)


def _build_workbook(tables: tuple) -> bytes:
    """Build xl/workbook.xml — sheet definitions."""
    root = etree.Element(SML_WORKBOOK, nsmap=_R_NSMAP)
    sheets_el = etree.SubElement(root, SML_SHEETS)

    for i, table in enumerate(tables):
        name = _safe_sheet_name(table.name) or f"Sheet{i + 1}"
        etree.SubElement(
            sheets_el,
            SML_SHEET,
            name=name,
            sheetId=str(i + 1),
            attrib={f"{{{R}}}id": f"rId{i + 1}"},
        )

    return _xml_decl(root)


def _build_workbook_rels(tables: tuple) -> bytes:
    """Build xl/_rels/workbook.xml.rels — sheet + styles + SST relationships."""
    root = etree.Element("Relationships", nsmap=_RELS_NSMAP)

    for i in range(len(tables)):
        etree.SubElement(
            root,
            "Relationship",
            Id=f"rId{i + 1}",
            Type=_RT_WORKSHEET,
            Target=f"worksheets/sheet{i + 1}.xml",
        )

    next_id = len(tables) + 1
    etree.SubElement(
        root,
        "Relationship",
        Id=f"rId{next_id}",
        Type=_RT_STYLES,
        Target="styles.xml",
    )
    etree.SubElement(
        root,
        "Relationship",
        Id=f"rId{next_id + 1}",
        Type=_RT_SHARED_STRINGS,
        Target="sharedStrings.xml",
    )

    return _xml_decl(root)


def _build_shared_strings(sst: dict[str, int]) -> bytes:
    """Build xl/sharedStrings.xml — deduplicated string table."""
    root = etree.Element(SML_SST, nsmap=_SML_NSMAP)
    root.set("count", str(len(sst)))
    root.set("uniqueCount", str(len(sst)))

    # Sort by index to maintain insertion order
    for s, _idx in sorted(sst.items(), key=lambda x: x[1]):
        si = etree.SubElement(root, SML_SI)
        t = etree.SubElement(si, SML_T)
        t.text = s

    return _xml_decl(root)


def _build_styles() -> bytes:
    """Build xl/styles.xml — number formats for typed columns.

    Style indices:
    0 = General (numbers, integers, text)
    1 = Date (numFmtId=14: mm-dd-yy)
    2 = DateTime (numFmtId=22: m/d/yy h:mm)
    3 = Time (numFmtId=21: h:mm:ss)
    4 = Duration (numFmtId=46: [h]:mm:ss)
    5 = Money (custom numFmtId=164: $#,##0.00)
    6 = Decimal (custom numFmtId=165: #,##0.00)
    7 = Bold (for header row)
    """
    root = etree.Element(f"{{{SML}}}styleSheet", nsmap=_SML_NSMAP)

    # Custom number formats
    numFmts = etree.SubElement(root, f"{{{SML}}}numFmts", count="2")
    etree.SubElement(numFmts, f"{{{SML}}}numFmt", numFmtId="164", formatCode="$#,##0.00")
    etree.SubElement(numFmts, f"{{{SML}}}numFmt", numFmtId="165", formatCode="#,##0.00")

    # Fonts: 0=default, 1=bold
    fonts = etree.SubElement(root, f"{{{SML}}}fonts", count="2")
    font0 = etree.SubElement(fonts, f"{{{SML}}}font")
    etree.SubElement(font0, f"{{{SML}}}sz", val="11")
    etree.SubElement(font0, f"{{{SML}}}name", val="Calibri")
    font1 = etree.SubElement(fonts, f"{{{SML}}}font")
    etree.SubElement(font1, f"{{{SML}}}b")
    etree.SubElement(font1, f"{{{SML}}}sz", val="11")
    etree.SubElement(font1, f"{{{SML}}}name", val="Calibri")

    # Fills (required minimum: none + gray125)
    fills = etree.SubElement(root, f"{{{SML}}}fills", count="2")
    fill0 = etree.SubElement(fills, f"{{{SML}}}fill")
    etree.SubElement(fill0, f"{{{SML}}}patternFill", patternType="none")
    fill1 = etree.SubElement(fills, f"{{{SML}}}fill")
    etree.SubElement(fill1, f"{{{SML}}}patternFill", patternType="gray125")

    # Borders (required minimum: one empty border)
    borders = etree.SubElement(root, f"{{{SML}}}borders", count="1")
    border0 = etree.SubElement(borders, f"{{{SML}}}border")
    for side in ("left", "right", "top", "bottom", "diagonal"):
        etree.SubElement(border0, f"{{{SML}}}{side}")

    # Cell style XFs (base)
    cellStyleXfs = etree.SubElement(root, f"{{{SML}}}cellStyleXfs", count="1")
    etree.SubElement(
        cellStyleXfs,
        f"{{{SML}}}xf",
        numFmtId="0",
        fontId="0",
        fillId="0",
        borderId="0",
    )

    # Cell XFs (actual cell formats)
    # 0=General, 1=Date, 2=DateTime, 3=Time, 4=Duration, 5=Money, 6=Decimal, 7=Bold
    cellXfs = etree.SubElement(root, f"{{{SML}}}cellXfs", count="8")
    etree.SubElement(
        cellXfs, f"{{{SML}}}xf", numFmtId="0", fontId="0", fillId="0", borderId="0", xfId="0"
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_DATE),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_DATETIME),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_TIME),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_DURATION),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_MONEY),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId=str(_NUMFMT_DECIMAL),
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
        applyNumberFormat="1",
    )
    etree.SubElement(
        cellXfs,
        f"{{{SML}}}xf",
        numFmtId="0",
        fontId="1",
        fillId="0",
        borderId="0",
        xfId="0",
        applyFont="1",
    )

    # Cell styles (required for openpyxl compatibility — "Normal" default style)
    cellStyles = etree.SubElement(root, f"{{{SML}}}cellStyles", count="1")
    etree.SubElement(cellStyles, f"{{{SML}}}cellStyle", name="Normal", xfId="0", builtinId="0")

    return _xml_decl(root)


_BOLD_STYLE_IDX = "7"


_MIN_COL_WIDTH = 8
_MAX_COL_WIDTH = 60


def _build_worksheet(table: Any, sst: dict[str, int], style_map: dict) -> bytes:
    """Build xl/worksheets/sheetN.xml — cell data with auto-sized columns."""

    root = etree.Element(SML_WORKSHEET, nsmap=_SML_NSMAP)

    n_cols = len(table.columns)
    n_rows = len(table.rows) + 1  # +1 for header

    # Dimension
    if n_cols > 0 and n_rows > 0:
        last_col = index_to_col_letters(n_cols - 1)
        etree.SubElement(root, f"{{{SML}}}dimension", ref=f"A1:{last_col}{n_rows}")

    # Track column widths for auto-sizing
    col_widths = [len(col.name) + 2 for col in table.columns]

    # Pre-scan data for column widths
    for row in table.rows:
        for col_idx in range(min(len(row), n_cols)):
            value = row[col_idx]
            if value is None:
                continue
            display_len = len(_display_string(value, table.columns[col_idx].column_type))
            if col_idx < len(col_widths) and display_len > col_widths[col_idx]:
                col_widths[col_idx] = display_len

    # Write column widths (<cols> element must come before <sheetData>)
    if col_widths:
        cols_el = etree.SubElement(root, f"{{{SML}}}cols")
        for i, width in enumerate(col_widths):
            clamped = max(_MIN_COL_WIDTH, min(width + 2, _MAX_COL_WIDTH))
            etree.SubElement(
                cols_el,
                f"{{{SML}}}col",
                min=str(i + 1),
                max=str(i + 1),
                width=str(clamped),
                customWidth="1",
            )

    # Sheet data
    sheet_data = etree.SubElement(root, SML_SHEET_DATA)

    # Header row (bold, style 7)
    header_row = etree.SubElement(sheet_data, SML_ROW, r="1")
    for col_idx, col in enumerate(table.columns):
        cell_ref = f"{index_to_col_letters(col_idx)}1"
        c = etree.SubElement(header_row, SML_CELL, r=cell_ref, t="s", s=_BOLD_STYLE_IDX)
        v = etree.SubElement(c, SML_VALUE)
        v.text = str(sst[col.name])

    # Freeze header row
    sv = etree.SubElement(root, f"{{{SML}}}sheetViews")
    sheet_view = etree.SubElement(sv, f"{{{SML}}}sheetView", workbookViewId="0", tabSelected="1")
    etree.SubElement(
        sheet_view,
        f"{{{SML}}}pane",
        ySplit="1",
        topLeftCell="A2",
        activePane="bottomLeft",
        state="frozen",
    )

    # Data rows
    for row_idx, row in enumerate(table.rows, start=2):
        row_el = etree.SubElement(sheet_data, SML_ROW, r=str(row_idx))

        for col_idx in range(min(len(row), n_cols)):
            value = row[col_idx]
            if value is None:
                continue

            col_type = table.columns[col_idx].column_type
            cell_ref = f"{index_to_col_letters(col_idx)}{row_idx}"
            style_idx = style_map.get(col_type)

            _write_cell_xml(row_el, cell_ref, value, col_type, style_idx, sst)

    # Merged cells
    merged = table.metadata.get("merged_ranges", [])
    if merged:
        mc_el = etree.SubElement(root, SML_MERGE_CELLS, count=str(len(merged)))
        for ref in merged:
            etree.SubElement(mc_el, SML_MERGE_CELL, ref=ref)

    return _xml_decl(root)


def _display_string(value: Any, col_type: Any) -> str:
    """Estimate the display width string for a cell value."""
    if value is None:
        return ""
    if isinstance(value, list):
        return "; ".join(str(v) for v in value)
    if isinstance(value, dict):
        return str(value)
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime.date, datetime.datetime)):
        return "2025-01-15"  # Fixed width for dates
    return str(value)


def _write_cell_xml(
    row_el: etree._Element,
    cell_ref: str,
    value: Any,
    col_type: Any,
    style_idx: int | None,
    sst: dict[str, int],
) -> None:
    """Write a single cell element to the row."""
    from kaos_content.model.tabular import ColumnType

    attribs: dict[str, str] = {"r": cell_ref}
    if style_idx is not None:
        attribs["s"] = str(style_idx)

    # Boolean
    if isinstance(value, bool):
        attribs["t"] = "b"
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        v = etree.SubElement(c, SML_VALUE)
        v.text = "1" if value else "0"
        return

    # Date/time → serial number with style
    # Coerce ISO date strings to actual dates when column type is DATE
    if col_type in (ColumnType.DATE, ColumnType.DATETIME, ColumnType.TIME):
        date_value = value
        if isinstance(value, str) and col_type == ColumnType.DATE:
            date_value = _try_parse_date(value)
        if isinstance(date_value, (datetime.date, datetime.datetime, datetime.time)):
            c = etree.SubElement(row_el, SML_CELL, **attribs)
            v = etree.SubElement(c, SML_VALUE)
            v.text = str(date_to_serial(date_value))
            return

    # Duration → fractional days
    if col_type == ColumnType.DURATION and isinstance(value, datetime.timedelta):
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        v = etree.SubElement(c, SML_VALUE)
        v.text = str(value.total_seconds() / 86400)
        return

    # Money → extract amount
    if col_type == ColumnType.MONEY and isinstance(value, dict):
        amount = value.get("amount")
        if amount is not None:
            c = etree.SubElement(row_el, SML_CELL, **attribs)
            v = etree.SubElement(c, SML_VALUE)
            v.text = str(float(amount))
            return

    # Decimal
    if isinstance(value, Decimal):
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        v = etree.SubElement(c, SML_VALUE)
        v.text = str(float(value))
        return

    # Numeric
    if isinstance(value, (int, float)):
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        v = etree.SubElement(c, SML_VALUE)
        v.text = str(value)
        return

    # String (shared string table)
    s = _to_string(value)
    if s in sst:
        attribs["t"] = "s"
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        v = etree.SubElement(c, SML_VALUE)
        v.text = str(sst[s])
    else:
        # Inline string fallback
        attribs["t"] = "inlineStr"
        c = etree.SubElement(row_el, SML_CELL, **attribs)
        is_el = etree.SubElement(c, f"{{{SML}}}is")
        t_el = etree.SubElement(is_el, SML_T)
        t_el.text = s


def _try_parse_date(s: str) -> datetime.date | str:
    """Try to parse an ISO date string. Returns original string on failure."""
    try:
        return datetime.date.fromisoformat(s)
    except (ValueError, TypeError):
        pass
    # Try common formats
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return s


def _safe_sheet_name(name: str) -> str:
    """Sanitize a table name for use as an Excel sheet name."""
    if not name:
        return "Sheet1"
    # Strip version suffixes like "-v2" for cleaner display
    clean = name
    if clean.endswith(("-v1", "-v2", "-v3")):
        clean = clean[:-3]
    # Replace hyphens with spaces, title case
    clean = clean.replace("-", " ").replace("_", " ").title()
    for ch in r"[]:*?/\\":
        clean = clean.replace(ch, "_")
    return clean[:31]
