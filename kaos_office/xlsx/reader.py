"""XLSX reader — dispatch to native (lxml+OPC) or calamine engine.

Default engine is ``"native"`` — uses lxml and our OPC infrastructure.
No external dependencies. Includes ZIP bomb, path traversal, and XML
bomb protection.

The ``"calamine"`` engine uses python-calamine (Rust, 7-28x faster)
and requires the ``[xlsx-calamine]`` optional extra.

Both engines produce identical ``TabularDocument`` output.

Usage::

    from kaos_office.xlsx.reader import parse_xlsx, list_sheets

    # Default (native lxml)
    doc = parse_xlsx("report.xlsx")

    # Calamine (faster for large files)
    doc = parse_xlsx("report.xlsx", engine="calamine")
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from kaos_content.model.tabular import TabularDocument


def parse_xlsx(
    path: str | Path,
    *,
    sheets: list[str] | None = None,
    max_rows: int | None = None,
    header_row: int = 0,
    include_formulas: bool = False,
    engine: Literal["native", "calamine"] = "native",
) -> TabularDocument:
    """Parse an XLSX file into a TabularDocument.

    Args:
        path: Path to the XLSX file.
        sheets: Specific sheet names. None = all visible sheets.
        max_rows: Maximum data rows per sheet. None = all.
        header_row: 0-based row index for column headers.
        include_formulas: Extract formula text into metadata.
        engine: ``"native"`` (default, lxml+OPC) or ``"calamine"``
            (Rust, faster, requires ``[xlsx-calamine]`` extra).

    Returns:
        TabularDocument with one Table per extracted sheet.
    """
    if engine == "calamine":
        from kaos_office.xlsx.calamine_reader import parse_xlsx_calamine

        return parse_xlsx_calamine(
            path,
            sheets=sheets,
            max_rows=max_rows,
            header_row=header_row,
            include_formulas=include_formulas,
        )

    from kaos_office.xlsx.native import parse_xlsx_native

    return parse_xlsx_native(
        path,
        sheets=sheets,
        max_rows=max_rows,
        header_row=header_row,
        include_formulas=include_formulas,
    )


def list_sheets(path: str | Path) -> list[dict[str, Any]]:
    """List all sheets in an XLSX file with metadata.

    Uses the native reader (no external dependencies needed for metadata).

    Returns a list of dicts with name, visible, rows, columns.
    """
    # Use OPC to read workbook.xml directly for sheet metadata
    from kaos_office.ooxml.namespace import SML_DIMENSION, SML_SHEET, SML_SHEETS, qn
    from kaos_office.opc.package import OPCPackage
    from kaos_office.xlsx.native import parse_xlsx_native as _

    p = Path(path).resolve()
    if not p.is_file():
        msg = f"File not found: {p}"
        raise FileNotFoundError(msg)

    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    result = []

    with OPCPackage.open(p) as pkg:
        wb_xml = pkg.read_xml("xl/workbook.xml")
        wb_rels = pkg.relationships("xl/workbook.xml")

        sheets_el = wb_xml.find(SML_SHEETS)
        if sheets_el is None:
            return result

        for sheet_el in sheets_el.iterchildren(SML_SHEET):
            name = sheet_el.get("name", "")
            state = sheet_el.get("state", "visible")
            rid = sheet_el.get(qn(R_NS, "id"), "")

            rows = 0
            cols = 0
            # Try to read dimensions from the sheet
            rel = wb_rels.get(rid)
            if rel is not None:
                sheet_path = f"xl/{rel.target}"
                if pkg.has_part(sheet_path):
                    sheet_xml = pkg.read_xml(sheet_path)
                    dim = sheet_xml.find(SML_DIMENSION)
                    if dim is not None:
                        ref = dim.get("ref", "")
                        if ":" in ref:
                            from kaos_office.xlsx.cell_ref import parse_cell_ref

                            _, end = ref.split(":")
                            r, c = parse_cell_ref(end)
                            rows = r + 1
                            cols = c + 1

            result.append(
                {
                    "name": name,
                    "type": "WorkSheet",
                    "visible": state == "visible",
                    "rows": rows,
                    "columns": cols,
                }
            )

    return result
