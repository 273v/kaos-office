"""kaos-office: Office document extraction for KAOS.

Extracts DOCX and PPTX files into kaos-content ContentDocument AST with provenance.
Uses lxml for XML parsing. PPTX uses python-pptx for shape traversal with OPC fallback
for SmartArt.
"""

from kaos_content.search import SearchResult, SearchResults, search_document

from kaos_office._version import __version__
from kaos_office.docx.reader import parse_docx
from kaos_office.pptx.reader import parse_pptx
from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    GetSlideTool,
    ListSlidesTool,
    ParseDocxTool,
    ParsePptxTool,
    SearchDocxTool,
    register_office_tools,
)

__all__ = [
    "DocxMetadataTool",
    "GetDocxMarkdownTool",
    "GetDocxTextTool",
    "GetSlideTool",
    "ListSlidesTool",
    "ParseDocxTool",
    "ParsePptxTool",
    "SearchDocxTool",
    "SearchResult",
    "SearchResults",
    "__version__",
    "parse_docx",
    "parse_pptx",
    "register_office_tools",
    "search_document",
]
