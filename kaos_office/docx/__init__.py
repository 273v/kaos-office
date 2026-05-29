"""WordprocessingML (DOCX) extraction and generation."""

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.redline import compare_docx, write_redline
from kaos_office.docx.writer import write_docx, write_docx_bytes

__all__ = [
    "compare_docx",
    "parse_docx",
    "write_docx",
    "write_docx_bytes",
    "write_redline",
]
