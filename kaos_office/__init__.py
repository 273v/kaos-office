"""kaos-office: Office document extraction for KAOS.

Extracts DOCX, PPTX, and XLSX files into kaos-content AST models:
- DOCX, PPTX → ContentDocument (flow content: headings, paragraphs, lists, tables)
- XLSX → TabularDocument (typed columns, row data, multi-sheet)

All three formats produce markdown via a unified entry point::

    from kaos_office import extract_to_markdown

    md = extract_to_markdown("report.docx")
    md = extract_to_markdown("slides.pptx")
    md = extract_to_markdown("data.xlsx")
"""

from __future__ import annotations

from pathlib import Path

from kaos_content.search import SearchResult, SearchResults, search_document

from kaos_office._version import __version__
from kaos_office.docx.reader import parse_docx
from kaos_office.errors import (
    DocxExtractionError,
    KaosOfficeError,
    PptxExtractionError,
    XlsxExtractionError,
)
from kaos_office.pptx.reader import parse_pptx
from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    GetSheetXlsxTool,
    GetSlideTool,
    ListSheetsXlsxTool,
    ListSlidesTool,
    ParseDocxTool,
    ParsePptxTool,
    ParseXlsxTool,
    SearchDocxTool,
    WriteDocxTool,
    WritePptxTool,
    WriteXlsxTool,
    XlsxMetadataTool,
    register_office_tools,
)


def extract_to_markdown(path: str | Path, **kwargs: object) -> str:
    """Extract any Office document to markdown.

    Dispatches by file extension:
    - ``.docx`` → ContentDocument → ``serialize_markdown()``
    - ``.pptx`` → ContentDocument → ``serialize_markdown()``
    - ``.xlsx`` → TabularDocument → ``serialize_tabular_markdown()``

    Args:
        path: Path to the Office document.
        **kwargs: Passed to the format-specific parser (e.g., ``header_row``
            for XLSX, ``sheets`` for specific sheet selection).

    Returns:
        Markdown string.

    Raises:
        ValueError: If the file extension is not supported.
        FileNotFoundError: If the file does not exist.
    """
    p = Path(path).resolve()
    ext = p.suffix.lower()

    if ext == ".docx":
        from kaos_content.serializers import serialize_markdown

        doc = parse_docx(p)
        return serialize_markdown(doc)

    if ext == ".pptx":
        from kaos_content.serializers import serialize_markdown

        doc = parse_pptx(p)
        return serialize_markdown(doc)

    if ext in (".xlsx", ".xlsm", ".xls"):
        from kaos_content.serializers.tabular import serialize_tabular_markdown

        from kaos_office.xlsx.reader import parse_xlsx

        doc = parse_xlsx(p, **kwargs)  # type: ignore[arg-type]  # ty: ignore[invalid-argument-type]
        return serialize_tabular_markdown(doc)

    supported = ".docx, .pptx, .xlsx"
    msg = f"Unsupported file extension: {ext!r}. Supported: {supported}"
    raise ValueError(msg)


__all__ = [
    "DocxExtractionError",
    "DocxMetadataTool",
    "GetDocxMarkdownTool",
    "GetDocxTextTool",
    "GetSheetXlsxTool",
    "GetSlideTool",
    "KaosOfficeError",
    "ListSheetsXlsxTool",
    "ListSlidesTool",
    "ParseDocxTool",
    "ParsePptxTool",
    "ParseXlsxTool",
    "PptxExtractionError",
    "SearchDocxTool",
    "SearchResult",
    "SearchResults",
    "WriteDocxTool",
    "WritePptxTool",
    "WriteXlsxTool",
    "XlsxExtractionError",
    "XlsxMetadataTool",
    "__version__",
    "extract_to_markdown",
    "parse_docx",
    "parse_pptx",
    "register_office_tools",
    "search_document",
]
