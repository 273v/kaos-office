"""MCP tool definitions for Office document extraction.

KaosTool implementations for DOCX and PPTX parsing, text extraction, search,
and metadata. Registered via register_office_tools(runtime).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from kaos_core import KaosContext, KaosRuntime, KaosTool, ToolMetadata, ToolResult
from kaos_core.types.annotations import ToolAnnotations
from kaos_core.types.enums import ToolCapability, ToolCategory
from kaos_core.types.parameters import ParameterSchema

_MODULE = "kaos-office"
_VERSION = "0.1.0"

# All Office tools are read-only, idempotent, and local-only.
_OFFICE_ANNOTATIONS = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)


def _validate_docx_path(path_str: str) -> Path | None:
    """Validate a DOCX file path exists and return it, or None."""
    p = Path(path_str)
    if not p.exists():
        return None
    return p


def _validate_pptx_path(path_str: str) -> Path | None:
    """Validate a PPTX file path exists and return it, or None."""
    p = Path(path_str)
    if not p.exists():
        return None
    return p


class ParseDocxTool(KaosTool):
    """Parse a DOCX file into a structured ContentDocument."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-parse-docx",
            display_name="Parse DOCX",
            description=(
                "Parse a DOCX file into a structured document with paragraphs, "
                "headings, tables, lists, and provenance. "
                "Returns a summary and resource link to the full document."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the DOCX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_docx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. "
                "Verify the path is correct. Use an absolute path if relative doesn't work."
            )

        try:
            from kaos_office.docx.reader import parse_docx

            doc = parse_docx(path)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to parse DOCX: {exc}. "
                "The file may be corrupted or password-protected. "
                "Try kaos-office-metadata to check file validity first."
            )

        # Store as artifact if context available
        if context and context.runtime:
            from kaos_content.artifacts import (
                document_outline,
                document_to_summary,
                store_document,
            )
            from kaos_content.views import DocumentView

            manifest = await store_document(doc, context.runtime, context, name=path.stem)
            summary = document_to_summary(doc, max_length=500)
            outline = document_outline(doc)
            view = DocumentView(doc)

            return manifest.to_tool_result(
                summary=summary,
                structured_content={
                    "artifact_id": manifest.artifact_id,
                    "title": doc.metadata.title,
                    "block_count": len(doc.body),
                    "has_sections": view.has_sections,
                    "outline": outline[:10],
                    "section_count": len(view.flat_sections),
                    "body_uri": manifest.body_uri,
                    "sections_uri": f"kaos://content/{manifest.artifact_id}/sections",
                },
            )

        # No runtime — return inline summary
        from kaos_content.serializers.text import serialize_text

        text = serialize_text(doc)
        blocks = len(doc.body)
        return ToolResult.create_success(
            f"Parsed {blocks} blocks from {path.name}.\n\n{text[:2000]}"
        )


class GetDocxTextTool(KaosTool):
    """Extract plain text from a DOCX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-text",
            display_name="Get DOCX Text",
            description=(
                "Extract plain text from a DOCX file. "
                "Returns the full document text without formatting."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the DOCX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_docx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_content.serializers.text import serialize_text

            from kaos_office.docx.reader import parse_docx

            doc = parse_docx(path)
            text = serialize_text(doc)
            return ToolResult.create_success(text)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to extract text: {exc}. "
                "Try kaos-office-parse-docx for more detailed error info."
            )


class GetDocxMarkdownTool(KaosTool):
    """Extract markdown from a DOCX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-markdown",
            display_name="Get DOCX Markdown",
            description=(
                "Extract content from a DOCX file as markdown. "
                "Preserves formatting (bold, italic, links, lists, tables)."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the DOCX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_docx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_content.serializers.markdown import serialize_markdown

            from kaos_office.docx.reader import parse_docx

            doc = parse_docx(path)
            md = serialize_markdown(doc)
            return ToolResult.create_success(md)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to extract markdown: {exc}. "
                "Try kaos-office-get-text for plain text extraction."
            )


class DocxMetadataTool(KaosTool):
    """Extract metadata from a DOCX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-metadata",
            display_name="DOCX Metadata",
            description=(
                "Extract metadata from a DOCX file: title, author, dates, "
                "word count, page count, and application info."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the DOCX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_docx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_office.docx.metadata import DocxMetadata
            from kaos_office.opc.package import OPCPackage

            with OPCPackage.open(path) as pkg:
                core_xml = (
                    pkg.read_part("docProps/core.xml")
                    if pkg.has_part("docProps/core.xml")
                    else None
                )
                app_xml = (
                    pkg.read_part("docProps/app.xml") if pkg.has_part("docProps/app.xml") else None
                )
                meta = DocxMetadata.from_xml(core_xml, app_xml)

            meta_dict = meta.to_dict()
            title = meta_dict.get("title", Path(path).name)
            summary = f"Metadata for {title}"
            return ToolResult.create_success(output=meta_dict, summary=summary)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to extract metadata: {exc}. "
                "The file may not be a valid DOCX. Check that it opens in Word or LibreOffice."
            )


class SearchDocxTool(KaosTool):
    """Search within a DOCX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-search",
            display_name="Search DOCX",
            description=(
                "Search for content within a DOCX file using BM25 ranking. "
                "Returns matching paragraphs with relevance scores and block references."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.QUERY,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the DOCX file.",
                ),
                ParameterSchema(
                    name="query",
                    type="string",
                    description="Search query text.",
                ),
                ParameterSchema(
                    name="top_k",
                    type="integer",
                    description="Maximum number of results to return.",
                    required=False,
                    default=10,
                ),
                ParameterSchema(
                    name="level",
                    type="string",
                    description="Search granularity: 'paragraph' or 'sentence'.",
                    required=False,
                    default="paragraph",
                    constraints={"enum": ["paragraph", "sentence"]},
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        query = inputs.get("query", "")
        top_k = inputs.get("top_k", 10)
        level = inputs.get("level", "paragraph")

        if not query:
            return ToolResult.create_error(
                "Query is required. Provide a search term to find in the document. "
                "Example: kaos-office-search path='doc.docx' query='force majeure'"
            )

        path = _validate_docx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_content.search import search_document

            from kaos_office.docx.reader import parse_docx

            doc = parse_docx(path)
            results = search_document(doc, query, top_k=top_k, level=level)

            result_data = {
                "query": results.query,
                "total_matches": results.total_matches,
                "has_more": results.has_more,
                "results": [
                    {
                        "text": r.text,
                        "score": round(r.score, 4),
                        "block_ref": r.block_ref,
                        "section_title": r.section_title,
                    }
                    for r in results.results
                ],
            }
            more = " (has more)" if results.has_more else ""
            summary = f"Found {results.total_matches} matches for '{results.query}'{more}"
            return ToolResult.create_success(output=result_data, summary=summary)
        except Exception as exc:
            return ToolResult.create_error(
                f"Search failed: {exc}. "
                "Try kaos-office-get-text to verify the document has extractable content."
            )


class ParsePptxTool(KaosTool):
    """Parse a PPTX file into a structured ContentDocument."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-parse-pptx",
            display_name="Parse PPTX",
            description=(
                "Parse a PPTX file into a structured document. Each slide becomes "
                "a section with headings, paragraphs, tables (including chart data), "
                "images, SmartArt text, and speaker notes."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the PPTX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_pptx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. "
                "Verify the path is correct. Use an absolute path if relative doesn't work."
            )

        try:
            from kaos_office.pptx.reader import parse_pptx

            doc = parse_pptx(path)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to parse PPTX: {exc}. "
                "The file may be corrupted or password-protected. "
                "Try kaos-office-list-slides for basic file validation."
            )

        # Store as artifact if context available
        if context and context.runtime:
            from kaos_content.artifacts import (
                document_outline,
                document_to_summary,
                store_document,
            )
            from kaos_content.views import DocumentView

            manifest = await store_document(doc, context.runtime, context, name=path.stem)
            summary = document_to_summary(doc, max_length=500)
            outline = document_outline(doc)
            view = DocumentView(doc)

            return manifest.to_tool_result(
                summary=summary,
                structured_content={
                    "artifact_id": manifest.artifact_id,
                    "title": doc.metadata.title,
                    "slide_count": len(doc.body),
                    "block_count": len(doc.body),
                    "has_sections": view.has_sections,
                    "outline": outline[:10],
                    "section_count": len(view.flat_sections),
                    "body_uri": manifest.body_uri,
                    "sections_uri": f"kaos://content/{manifest.artifact_id}/sections",
                },
            )

        # No runtime — return inline summary
        from kaos_content.serializers.text import serialize_text

        text = serialize_text(doc)
        blocks = len(doc.body)
        return ToolResult.create_success(
            f"Parsed {blocks} slides from {path.name}.\n\n{text[:2000]}"
        )


class ListSlidesTool(KaosTool):
    """List slides in a PPTX file with titles and metadata."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-list-slides",
            display_name="List PPTX Slides",
            description=(
                "List all slides in a PPTX file with their titles, shape counts, "
                "and whether they have speaker notes."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the PPTX file.",
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        path = _validate_pptx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_office.pptx.reader import list_slides

            slides = list_slides(path)
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to list slides: {exc}. "
                "The file may be corrupted. Try opening it in PowerPoint or LibreOffice."
            )

        summary = f"Found {len(slides)} slides"
        return ToolResult.create_success(output={"slides": slides}, summary=summary)


class GetSlideTool(KaosTool):
    """Extract text from a specific slide in a PPTX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-slide",
            display_name="Get PPTX Slide",
            description=(
                "Extract text content from a specific slide in a PPTX file. "
                "Uses 1-based slide numbering. Use kaos-office-list-slides first "
                "to see available slides."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the PPTX file.",
                ),
                ParameterSchema(
                    name="slide_number",
                    type="integer",
                    description="Slide number (1-based).",
                    constraints={"minimum": 1},
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        slide_number = inputs.get("slide_number", 1)

        path = _validate_pptx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_office.pptx.reader import get_slide_text

            text = get_slide_text(path, slide_number)
            return ToolResult.create_success(text)
        except ValueError as exc:
            return ToolResult.create_error(
                f"{exc}. Use kaos-office-list-slides to see available slide numbers."
            )
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to extract slide: {exc}. "
                "Try kaos-office-parse-pptx for full document extraction."
            )


class SearchPptxTool(KaosTool):
    """Search within a PPTX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-search-pptx",
            display_name="Search PPTX",
            description=(
                "Search for content within a PPTX file using BM25 ranking. "
                "Returns matching text with relevance scores, slide numbers, "
                "and block references. Use kaos-office-get-slide to read full slide content."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.QUERY,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the PPTX file.",
                ),
                ParameterSchema(
                    name="query",
                    type="string",
                    description="Search query text.",
                ),
                ParameterSchema(
                    name="top_k",
                    type="integer",
                    description="Maximum number of results to return.",
                    required=False,
                    default=10,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        query = inputs.get("query", "")
        top_k = inputs.get("top_k", 10)

        if not query:
            return ToolResult.create_error(
                "Query is required. Provide a search term to find in the presentation. "
                "Example: kaos-office-search-pptx path='slides.pptx' query='revenue growth'"
            )

        path = _validate_pptx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_content.search import search_document

            from kaos_office.pptx.reader import parse_pptx

            doc = parse_pptx(path)
            results = search_document(doc, query, top_k=top_k)

            result_data = {
                "query": results.query,
                "total_matches": results.total_matches,
                "has_more": results.has_more,
                "results": [
                    {
                        "text": r.text,
                        "score": round(r.score, 4),
                        "block_ref": r.block_ref,
                        "section_title": r.section_title,
                    }
                    for r in results.results
                ],
            }
            more = " (has more)" if results.has_more else ""
            summary = f"Found {results.total_matches} matches for '{results.query}'{more}"
            return ToolResult.create_success(output=result_data, summary=summary)
        except Exception as exc:
            return ToolResult.create_error(
                f"Search failed: {exc}. "
                "Try kaos-office-parse-pptx to verify the presentation has extractable content."
            )


class GetSlideNotesTool(KaosTool):
    """Extract speaker notes from a specific slide in a PPTX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-slide-notes",
            display_name="Get PPTX Slide Notes",
            description=(
                "Extract speaker notes from a specific slide in a PPTX file. "
                "Uses 1-based slide numbering. Use kaos-office-list-slides first "
                "to see which slides have notes."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(
                    name="path",
                    type="string",
                    description="Path to the PPTX file.",
                ),
                ParameterSchema(
                    name="slide",
                    type="integer",
                    description="Slide number (1-based).",
                    constraints={"minimum": 1},
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        slide_number = inputs.get("slide", 1)

        path = _validate_pptx_path(path_str)
        if path is None:
            return ToolResult.create_error(
                f"File not found: {path_str}. Verify the path is correct and the file exists."
            )

        try:
            from kaos_office.pptx.reader import get_slide_notes

            notes = get_slide_notes(path, slide_number)
            if notes is None:
                return ToolResult.create_success(
                    f"Slide {slide_number} has no speaker notes. "
                    "Use kaos-office-list-slides to see which slides have notes."
                )
            return ToolResult.create_success(notes)
        except ValueError as exc:
            return ToolResult.create_error(
                f"{exc}. Use kaos-office-list-slides to see available slide numbers."
            )
        except Exception as exc:
            return ToolResult.create_error(
                f"Failed to extract slide notes: {exc}. "
                "Try kaos-office-parse-pptx for full document extraction."
            )


def _validate_xlsx_path(path_str: str) -> Path | None:
    """Validate an XLSX file path exists and return it, or None."""
    p = Path(path_str)
    if not p.exists():
        return None
    return p


class ParseXlsxTool(KaosTool):
    """Parse an XLSX file into a TabularDocument."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-parse-xlsx",
            display_name="Parse XLSX",
            description=(
                "Parse an Excel workbook into a structured TabularDocument with typed columns. "
                "Each worksheet becomes a table. Use kaos-office-list-sheets-xlsx first."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(name="path", type="string", description="Path to the XLSX file."),
                ParameterSchema(
                    name="sheets",
                    type="array",
                    description="Sheet names to extract. Default: all visible.",
                    required=False,
                ),
                ParameterSchema(
                    name="header_row",
                    type="integer",
                    description="0-based header row index. Default: 0.",
                    required=False,
                    default=0,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs["path"]
        path = _validate_xlsx_path(path_str)
        if path is None:
            return ToolResult.create_error(f"File not found: {path_str}.")

        try:
            from kaos_content.serializers.tabular import serialize_tabular_summary

            from kaos_office.xlsx.reader import parse_xlsx

            doc = parse_xlsx(
                path, sheets=inputs.get("sheets"), header_row=inputs.get("header_row", 0)
            )
            summary = serialize_tabular_summary(doc)

            if context is not None and context.runtime is not None:
                from kaos_content.artifacts import store_tabular

                manifest = await store_tabular(doc, context.runtime, context, name=path.stem)
                return manifest.to_tool_result(
                    summary=summary,
                    structured_content={
                        "artifact_id": manifest.artifact_id,
                        "table_count": len(doc.tables),
                        "tables": [{"name": t.name, "row_count": t.row_count} for t in doc.tables],
                    },
                )
            return ToolResult.create_text(summary)
        except ImportError:
            return ToolResult.create_error(
                "XLSX requires python-calamine. pip install kaos-office[xlsx]"
            )
        except Exception as exc:
            return ToolResult.create_error(f"XLSX extraction failed: {exc}.")


class ListSheetsXlsxTool(KaosTool):
    """List sheets in an XLSX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-list-sheets-xlsx",
            display_name="List XLSX Sheets",
            description="List all sheets in an Excel workbook with dimensions.",
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(name="path", type="string", description="Path to the XLSX file."),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path = _validate_xlsx_path(inputs["path"])
        if path is None:
            return ToolResult.create_error(f"File not found: {inputs['path']}.")
        try:
            from kaos_office.xlsx.reader import list_sheets

            return ToolResult.create_success(output={"sheets": list_sheets(path)})
        except ImportError:
            return ToolResult.create_error(
                "XLSX requires python-calamine. pip install kaos-office[xlsx]"
            )
        except Exception as exc:
            return ToolResult.create_error(f"Failed to list sheets: {exc}.")


class GetSheetXlsxTool(KaosTool):
    """Get a single sheet's data as TSV."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-sheet-xlsx",
            display_name="Get XLSX Sheet",
            description="Extract a single sheet's data. Use list-sheets-xlsx to see available sheets.",
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(name="path", type="string", description="Path to the XLSX file."),
                ParameterSchema(name="sheet", type="string", description="Sheet name."),
                ParameterSchema(
                    name="max_rows",
                    type="integer",
                    description="Max rows. Default: 100.",
                    required=False,
                    default=100,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path = _validate_xlsx_path(inputs["path"])
        if path is None:
            return ToolResult.create_error(f"File not found: {inputs['path']}.")
        try:
            from kaos_content.serializers.tabular import serialize_tsv

            from kaos_office.xlsx.reader import parse_xlsx

            doc = parse_xlsx(path, sheets=[inputs["sheet"]], max_rows=inputs.get("max_rows", 100))
            if not doc.tables:
                return ToolResult.create_error(f"Sheet '{inputs['sheet']}' not found.")
            return ToolResult.create_text(serialize_tsv(doc.tables[0]))
        except ImportError:
            return ToolResult.create_error(
                "XLSX requires python-calamine. pip install kaos-office[xlsx]"
            )
        except Exception as exc:
            return ToolResult.create_error(f"Failed to extract sheet: {exc}.")


class XlsxMetadataTool(KaosTool):
    """Show XLSX workbook metadata."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-xlsx-metadata",
            display_name="XLSX Metadata",
            description="Show workbook metadata: sheet names, dimensions, column types.",
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(name="path", type="string", description="Path to the XLSX file."),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path = _validate_xlsx_path(inputs["path"])
        if path is None:
            return ToolResult.create_error(f"File not found: {inputs['path']}.")
        try:
            from kaos_content.artifacts import tabular_summary

            from kaos_office.xlsx.reader import parse_xlsx

            doc = parse_xlsx(path)
            return ToolResult.create_success(output=tabular_summary(doc))
        except ImportError:
            return ToolResult.create_error(
                "XLSX requires python-calamine. pip install kaos-office[xlsx]"
            )
        except Exception as exc:
            return ToolResult.create_error(f"Failed to read XLSX metadata: {exc}.")


def register_office_tools(runtime: KaosRuntime) -> int:
    """Register all Office MCP tools with a runtime."""
    tools: list[KaosTool] = [
        ParseDocxTool(),
        GetDocxTextTool(),
        GetDocxMarkdownTool(),
        DocxMetadataTool(),
        SearchDocxTool(),
        ParsePptxTool(),
        ListSlidesTool(),
        GetSlideTool(),
        SearchPptxTool(),
        GetSlideNotesTool(),
        ParseXlsxTool(),
        ListSheetsXlsxTool(),
        GetSheetXlsxTool(),
        XlsxMetadataTool(),
    ]
    for tool in tools:
        runtime.tools.register_tool(tool)
    return len(tools)
