"""XLSX extraction and generation — TabularDocument ↔ SpreadsheetML.

Reader: :func:`parse_xlsx` (native lxml by default; ``engine="calamine"``
for the Rust fast-path) and :func:`list_sheets` for cheap workbook
metadata.

Writer: :func:`write_xlsx` and :func:`write_xlsx_bytes` (native lxml,
no extra dependencies).
"""

from kaos_office.xlsx.reader import list_sheets, parse_xlsx
from kaos_office.xlsx.writer import write_xlsx, write_xlsx_bytes

__all__ = [
    "list_sheets",
    "parse_xlsx",
    "write_xlsx",
    "write_xlsx_bytes",
]
