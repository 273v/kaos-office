"""kaos-office: Office document extraction for KAOS.

Extracts DOCX files into kaos-content ContentDocument AST with provenance.
Uses lxml for XML parsing — no python-docx dependency.
"""

from kaos_office._version import __version__
from kaos_office.docx.reader import parse_docx
from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    ParseDocxTool,
    SearchDocxTool,
    register_office_tools,
)

__all__ = [
    "DocxMetadataTool",
    "GetDocxMarkdownTool",
    "GetDocxTextTool",
    "ParseDocxTool",
    "SearchDocxTool",
    "__version__",
    "parse_docx",
    "register_office_tools",
]
