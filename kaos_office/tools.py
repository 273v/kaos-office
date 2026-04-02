"""MCP tool definitions for Office document extraction.

KaosTool implementations for DOCX parsing, text extraction, search,
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
            from kaos_content.artifacts import document_to_summary, store_document

            manifest = await store_document(doc, context.runtime, context, name=path.stem)
            summary = document_to_summary(doc)
            return ToolResult.create_success(
                f"{summary}\n\nFull document stored: {manifest.artifact_id}"
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

            import json

            return ToolResult.create_success(json.dumps(meta.to_dict(), indent=2))
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

            import json

            output = {
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
            return ToolResult.create_success(json.dumps(output, indent=2))
        except Exception as exc:
            return ToolResult.create_error(
                f"Search failed: {exc}. "
                "Try kaos-office-get-text to verify the document has extractable content."
            )


def register_office_tools(runtime: KaosRuntime) -> int:
    """Register all Office MCP tools with a runtime.

    Args:
        runtime: KaosRuntime to register tools with.

    Returns:
        Number of tools registered.
    """
    tools: list[KaosTool] = [
        ParseDocxTool(),
        GetDocxTextTool(),
        GetDocxMarkdownTool(),
        DocxMetadataTool(),
        SearchDocxTool(),
    ]
    for tool in tools:
        runtime.tools.register_tool(tool)
    return len(tools)
