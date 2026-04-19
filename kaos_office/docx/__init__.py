"""WordprocessingML (DOCX) extraction and generation."""

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx, write_docx_bytes

__all__ = ["parse_docx", "write_docx", "write_docx_bytes"]
