"""XLSX writer — reference implementation kept for cross-validation.

Production code uses :mod:`kaos_office.xlsx.writer` (native lxml,
zero extra deps). This module wraps ``xlsxwriter`` (BSD-2-Clause) and
is kept as an independent encoder against which the native writer's
output is differentially validated in tests. It is not part of the
published surface.

``xlsxwriter`` is a dev-only dependency (declared in
``[dependency-groups].dev``). It used to ship as the public
``[xlsx-write]`` extra; that extra was removed in 0.1.0a1 because
production never depended on this code path.
"""

from __future__ import annotations

import contextlib
import datetime
import io
import json
from decimal import Decimal
from pathlib import Path
from typing import Any

from kaos_core.logging import get_logger

from kaos_office.xlsx.styles import date_to_serial

logger = get_logger(__name__)


def write_xlsx(
    doc: Any,
    path: str | Path,
    *,
    bold_headers: bool = True,
    auto_width: bool = True,
    freeze_header: bool = True,
) -> Path:
    """Write a TabularDocument to an XLSX file.

    Args:
        doc: A ``TabularDocument`` from kaos-content.
        path: Output file path.
        bold_headers: Apply bold formatting to the header row.
        auto_width: Auto-size column widths based on content.
        freeze_header: Freeze the top row for scrolling.

    Returns:
        The output path.

    Raises:
        ImportError: If xlsxwriter is not installed.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    try:
        import xlsxwriter
    except ImportError as exc:
        msg = (
            "xlsxwriter is not installed. This module is the dev-only reference "
            "writer used for cross-validation; it is not in any published extra. "
            "Install via the dev group: `uv sync --group dev`. "
            "For production use, call kaos_office.xlsx.writer (native lxml, no "
            "extra deps) or use serialize_csv() from kaos-content for CSV output."
        )
        raise ImportError(msg) from exc

    wb = xlsxwriter.Workbook(str(path))
    try:
        _write_workbook(
            wb,
            doc,
            bold_headers=bold_headers,
            auto_width=auto_width,
            freeze_header=freeze_header,
        )
    finally:
        wb.close()

    logger.info(
        "xlsx.writer: wrote %s, tables=%d, path=%s",
        doc.metadata.title or "untitled",
        len(doc.tables),
        path,
    )
    return path


def write_xlsx_bytes(
    doc: Any,
    *,
    bold_headers: bool = True,
    auto_width: bool = True,
    freeze_header: bool = True,
) -> bytes:
    """Write a TabularDocument to XLSX bytes (in-memory).

    Same as ``write_xlsx`` but returns bytes instead of writing to disk.
    Useful for VFS/artifact flow.
    """
    try:
        import xlsxwriter
    except ImportError as exc:
        msg = (
            "xlsxwriter is not installed. This module is the dev-only reference "
            "writer used for cross-validation; it is not in any published extra. "
            "Install via the dev group: `uv sync --group dev`. "
            "For production use, call kaos_office.xlsx.writer (native lxml, no "
            "extra deps) or use serialize_csv() from kaos-content for CSV output."
        )
        raise ImportError(msg) from exc

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    try:
        _write_workbook(
            wb,
            doc,
            bold_headers=bold_headers,
            auto_width=auto_width,
            freeze_header=freeze_header,
        )
    finally:
        wb.close()

    return buf.getvalue()


def _write_workbook(
    wb: Any,
    doc: Any,
    *,
    bold_headers: bool,
    auto_width: bool,
    freeze_header: bool,
) -> None:
    """Write all tables from a TabularDocument into the workbook."""
    from kaos_content.model.tabular import ColumnType

    # Create formats
    bold_fmt = wb.add_format({"bold": True})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd"})
    datetime_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
    time_fmt = wb.add_format({"num_format": "hh:mm:ss"})
    money_fmt = wb.add_format({"num_format": "$#,##0.00"})
    decimal_fmt = wb.add_format({"num_format": "#,##0.00"})
    duration_fmt = wb.add_format({"num_format": "[h]:mm:ss"})

    format_map = {
        ColumnType.DATE: date_fmt,
        ColumnType.DATETIME: datetime_fmt,
        ColumnType.TIME: time_fmt,
        ColumnType.MONEY: money_fmt,
        ColumnType.DECIMAL: decimal_fmt,
        ColumnType.DURATION: duration_fmt,
    }

    for table in doc.tables:
        sheet_name = _safe_sheet_name(table.name)
        ws = wb.add_worksheet(sheet_name)

        columns = table.columns
        n_cols = len(columns)

        # Track max widths for auto-sizing
        col_widths = [len(c.name) for c in columns] if auto_width else []

        # Write header row
        for col_idx, col in enumerate(columns):
            fmt = bold_fmt if bold_headers else None
            ws.write(0, col_idx, col.name, fmt)

        # Freeze header
        if freeze_header:
            ws.freeze_panes(1, 0)

        # Write data rows
        for row_idx, row in enumerate(table.rows, start=1):
            for col_idx in range(min(len(row), n_cols)):
                value = row[col_idx]
                col_type = (
                    columns[col_idx].column_type if col_idx < len(columns) else ColumnType.TEXT
                )
                cell_fmt = format_map.get(col_type)

                _write_cell(ws, row_idx, col_idx, value, col_type, cell_fmt)

                # Track width
                if auto_width and col_idx < len(col_widths):
                    display = _display_width(value)
                    if display > col_widths[col_idx]:
                        col_widths[col_idx] = display

        # Apply column widths
        if auto_width:
            for col_idx, width in enumerate(col_widths):
                ws.set_column(col_idx, col_idx, min(width + 2, 60))

        # Restore merged cells from metadata
        merged = table.metadata.get("merged_ranges", [])
        for merge_ref in merged:
            with contextlib.suppress(Exception):
                ws.merge_range(merge_ref, "")

        logger.debug(
            "xlsx.writer: wrote sheet %s, rows=%d, cols=%d",
            sheet_name,
            len(table.rows),
            n_cols,
        )


def _write_cell(
    ws: Any,
    row: int,
    col: int,
    value: Any,
    col_type: Any,
    fmt: Any,
) -> None:
    """Write a single cell with appropriate type handling."""
    from kaos_content.model.tabular import ColumnType

    if value is None:
        return  # Skip null cells

    # Date/time types → serial number with format
    if col_type in (ColumnType.DATE, ColumnType.DATETIME, ColumnType.TIME) and isinstance(
        value, (datetime.date, datetime.datetime, datetime.time)
    ):
        serial = date_to_serial(value)
        ws.write_number(row, col, serial, fmt)
        return

    # Duration → fractional days
    if col_type == ColumnType.DURATION and isinstance(value, datetime.timedelta):
        ws.write_number(row, col, value.total_seconds() / 86400, fmt)
        return

    # Money → extract amount
    if col_type == ColumnType.MONEY and isinstance(value, dict):
        amount = value.get("amount")
        if amount is not None:
            ws.write_number(row, col, float(amount), fmt)
            return
        ws.write_string(row, col, str(value))
        return

    # Decimal → float
    if isinstance(value, Decimal):
        ws.write_number(row, col, float(value), fmt)
        return

    # Boolean
    if isinstance(value, bool):
        ws.write_boolean(row, col, value)
        return

    # Numbers
    if isinstance(value, (int, float)):
        ws.write_number(row, col, value, fmt)
        return

    # Lists/dicts → JSON string
    if isinstance(value, (list, dict)):
        ws.write_string(row, col, json.dumps(value, default=str))
        return

    # Default: string
    ws.write_string(row, col, str(value))


def _safe_sheet_name(name: str) -> str:
    """Sanitize a table name for use as an Excel sheet name.

    Excel sheet names: max 31 chars, no []:*?/\\ characters.
    """
    if not name:
        return "Sheet1"
    # Remove invalid chars
    for ch in r"[]:*?/\\":
        name = name.replace(ch, "_")
    return name[:31]


_MAX_DISPLAY_WIDTH = 50


def _display_width(value: Any) -> int:
    """Estimate the display width of a cell value for auto-sizing."""
    if value is None:
        return 0
    s = str(value)
    return min(len(s), _MAX_DISPLAY_WIDTH)
