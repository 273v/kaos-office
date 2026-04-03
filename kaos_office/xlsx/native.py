"""Native lxml XLSX reader — parse SpreadsheetML XML via OPC package.

Uses our shared OPC infrastructure (ZIP bomb detection, path traversal
prevention, XML bomb protection) and lxml for XML parsing. Zero external
dependencies beyond lxml (already a core kaos-office dependency).

Same two-pass architecture as the DOCX reader:
1. Metadata load: workbook.xml, sharedStrings.xml, styles.xml
2. Data walk: for each sheet, walk rows → cells → typed values
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from kaos_content.model.attr import Provenance, SourceRef
from kaos_content.model.metadata import DocumentMetadata
from kaos_content.model.tabular import (
    Column,
    Table,
    TabularDocument,
    infer_column_type,
)
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import (
    SML_CELL,
    SML_FORMULA,
    SML_INLINE_STR,
    SML_MERGE_CELL,
    SML_MERGE_CELLS,
    SML_ROW,
    SML_SHEET,
    SML_SHEET_DATA,
    SML_SHEETS,
    SML_T,
    SML_VALUE,
    SML_WORKBOOK_PR,
    XLSX_MIME_TYPE,
    R,
    qn,
)
from kaos_office.opc.package import OPCPackage
from kaos_office.xlsx.cell_ref import parse_cell_ref
from kaos_office.xlsx.shared_strings import SharedStringTable
from kaos_office.xlsx.styles import StyleTable, serial_to_date


def parse_xlsx_native(
    path: str | Path,
    *,
    sheets: list[str] | None = None,
    max_rows: int | None = None,
    header_row: int = 0,
    include_formulas: bool = False,
) -> TabularDocument:
    """Parse XLSX using native lxml + OPC. No external dependencies.

    Args:
        path: Path to the XLSX file.
        sheets: Specific sheet names. None = all visible.
        max_rows: Max data rows per sheet. None = all.
        header_row: 0-based row index for headers.
        include_formulas: Store formula text in metadata.

    Returns:
        TabularDocument with one Table per sheet.
    """
    p = Path(path).resolve()
    if not p.is_file():
        msg = f"File not found: {p}"
        raise FileNotFoundError(msg)

    with OPCPackage.open(p) as pkg:
        # --- Pass 1: Metadata ---

        # Parse workbook.xml → sheet names and relationship IDs
        wb_xml = pkg.read_xml("xl/workbook.xml")
        sheet_info = _parse_workbook(wb_xml)

        # Check for 1904 date system
        date1904 = _is_date1904(wb_xml)

        # Resolve sheet paths via relationships
        wb_rels = pkg.relationships("xl/workbook.xml")

        # Parse shared strings
        sst: SharedStringTable | None = None
        if pkg.has_part("xl/sharedStrings.xml"):
            sst = SharedStringTable(pkg.read_xml("xl/sharedStrings.xml"))

        # Parse styles for date detection
        style_table: StyleTable | None = None
        if pkg.has_part("xl/styles.xml"):
            style_table = StyleTable(pkg.read_xml("xl/styles.xml"))

        # --- Pass 2: Sheet data ---
        selected = _select_sheets(sheet_info, sheets)
        tables: list[Table] = []

        for name, rid, _visible in selected:
            rel = wb_rels.get(rid)
            if rel is None:
                continue
            # Resolve relative path from workbook location
            sheet_path = f"xl/{rel.target}"
            if not pkg.has_part(sheet_path):
                continue

            sheet_xml = pkg.read_xml(sheet_path)
            table = _parse_sheet(
                sheet_xml,
                name=name,
                sst=sst,
                style_table=style_table,
                date1904=date1904,
                header_row=header_row,
                max_rows=max_rows,
                include_formulas=include_formulas,
            )
            tables.append(table)

    source = SourceRef(uri=p.as_uri(), mime_type=XLSX_MIME_TYPE)
    return TabularDocument(
        metadata=DocumentMetadata(
            title=p.stem,
            source=source,
            document_type="xlsx",
            extra={"sheet_count": len(tables)},
        ),
        tables=tuple(tables),
        provenance=Provenance(source=source, extractor="kaos-office/xlsx/native"),
    )


# ---------------------------------------------------------------------------
# Workbook parsing
# ---------------------------------------------------------------------------

_R_ID = qn(R, "id")


def _parse_workbook(
    wb_xml: etree._Element,
) -> list[tuple[str, str, bool]]:
    """Extract sheet info from workbook.xml.

    Returns list of (name, rId, visible).
    """
    result = []
    sheets_el = wb_xml.find(SML_SHEETS)
    if sheets_el is None:
        return result
    for sheet_el in sheets_el.iterchildren(SML_SHEET):
        name = sheet_el.get("name", "")
        rid = sheet_el.get(_R_ID, "")
        state = sheet_el.get("state", "visible")
        result.append((name, rid, state == "visible"))
    return result


def _is_date1904(wb_xml: etree._Element) -> bool:
    """Check if workbook uses the 1904 date system."""
    pr = wb_xml.find(SML_WORKBOOK_PR)
    if pr is not None:
        return pr.get("date1904", "0") in ("1", "true")
    return False


def _select_sheets(
    sheet_info: list[tuple[str, str, bool]],
    selected: list[str] | None,
) -> list[tuple[str, str, bool]]:
    """Filter sheets by name or visibility."""
    if selected is not None:
        return [(n, r, v) for n, r, v in sheet_info if n in selected]
    return [(n, r, v) for n, r, v in sheet_info if v]


# ---------------------------------------------------------------------------
# Sheet parsing
# ---------------------------------------------------------------------------


def _parse_sheet(
    sheet_xml: etree._Element,
    *,
    name: str,
    sst: SharedStringTable | None,
    style_table: StyleTable | None,
    date1904: bool,
    header_row: int,
    max_rows: int | None,
    include_formulas: bool,
) -> Table:
    """Parse a single worksheet XML into a Table."""
    sheet_data = sheet_xml.find(SML_SHEET_DATA)
    if sheet_data is None:
        return Table(name=name)

    # Collect all rows as sparse data
    all_rows: list[tuple[int, dict[int, Any]]] = []
    formulas: dict[str, str] = {}
    max_col = 0

    for row_el in sheet_data.iterchildren(SML_ROW):
        row_num = int(row_el.get("r", "0")) - 1  # 0-based
        cells: dict[int, Any] = {}

        for cell_el in row_el.iterchildren(SML_CELL):
            ref = cell_el.get("r", "")
            if not ref:
                continue
            _, col_idx = parse_cell_ref(ref)
            if col_idx > max_col:
                max_col = col_idx

            value = _extract_cell_value(
                cell_el, sst=sst, style_table=style_table, date1904=date1904
            )
            cells[col_idx] = value

            # Extract formula if requested
            if include_formulas:
                f_el = cell_el.find(SML_FORMULA)
                if f_el is not None and f_el.text:
                    formulas[ref] = f"={f_el.text}"

        if cells:
            all_rows.append((row_num, cells))

    if not all_rows:
        return Table(name=name)

    # Convert sparse rows to dense
    n_cols = max_col + 1
    dense_rows: list[list[Any]] = []
    for _, cells in sorted(all_rows, key=lambda x: x[0]):
        row = [cells.get(c) for c in range(n_cols)]
        dense_rows.append(row)

    # Split header and data
    if header_row >= len(dense_rows):
        headers = [f"column_{i}" for i in range(n_cols)]
        data_rows = dense_rows
    else:
        header_vals = dense_rows[header_row]
        headers = [str(v) if v is not None else f"column_{i}" for i, v in enumerate(header_vals)]
        data_rows = dense_rows[header_row + 1 :]

    if not headers:
        return Table(name=name)

    # Apply max_rows
    total_rows = len(data_rows)
    if max_rows is not None and len(data_rows) > max_rows:
        data_rows = data_rows[:max_rows]

    # Normalize to header width
    n_cols = len(headers)
    normalized = [tuple((row + [None] * n_cols)[:n_cols]) for row in data_rows]

    # Infer column types
    columns: list[Column] = []
    for i, hdr in enumerate(headers):
        col_vals = [row[i] for row in normalized if i < len(row)]
        ct = infer_column_type(col_vals)
        nullable = any(v is None for v in col_vals)
        columns.append(Column(name=hdr, column_type=ct, nullable=nullable))

    # Merged cells metadata
    meta: dict[str, Any] = {}
    merge_cells_el = sheet_xml.find(SML_MERGE_CELLS)
    if merge_cells_el is not None:
        ranges = []
        for mc in merge_cells_el.iterchildren(SML_MERGE_CELL):
            ref = mc.get("ref", "")
            if ref:
                ranges.append(ref)
        if ranges:
            meta["merged_ranges"] = ranges

    if formulas:
        meta["formulas"] = formulas

    return Table(
        name=name,
        columns=tuple(columns),
        rows=tuple(normalized),
        row_count=total_rows,
        metadata=meta,
    )


# ---------------------------------------------------------------------------
# Cell value extraction
# ---------------------------------------------------------------------------


def _extract_cell_value(
    cell_el: etree._Element,
    *,
    sst: SharedStringTable | None,
    style_table: StyleTable | None,
    date1904: bool,
) -> Any:
    """Extract a typed Python value from a cell XML element."""
    cell_type = cell_el.get("t")
    style_idx = int(cell_el.get("s", "0"))

    # Inline string: <c t="inlineStr"><is><t>text</t></is></c>
    if cell_type == "inlineStr":
        is_el = cell_el.find(SML_INLINE_STR)
        if is_el is not None:
            t_el = is_el.find(SML_T)
            return t_el.text if t_el is not None else ""
        return ""

    v_el = cell_el.find(SML_VALUE)
    if v_el is None or v_el.text is None:
        return None

    raw = v_el.text

    # Shared string: index into SST
    if cell_type == "s":
        if sst is None:
            return raw
        try:
            return sst.get(int(raw))
        except (ValueError, IndexError):
            return raw

    # Boolean
    if cell_type == "b":
        return raw == "1"

    # Error
    if cell_type == "e":
        return None  # Treat errors as null

    # Formula string result
    if cell_type == "str":
        return raw

    # Number (explicit "n" or absent type — the default)
    try:
        # Check if it's a date via style
        if style_table is not None and style_table.is_date(style_idx):
            serial = float(raw)
            return serial_to_date(serial, date1904=date1904)

        # Integer or float
        if "." in raw or "E" in raw or "e" in raw:
            return float(raw)
        return int(raw)
    except (ValueError, OverflowError):
        return raw
