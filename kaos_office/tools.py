"""MCP tool definitions for Office document extraction.

KaosTool implementations for DOCX and PPTX parsing, text extraction, search,
and metadata. Registered via register_office_tools(runtime).

Every file-input tool routes its ``path`` parameter through
:func:`kaos_office._path_resolver.resolve_office_input` which delegates
to :func:`kaos_core.path_resolver.resolve_input_path`. That covers four
input shapes the agent might pass:

* ``kaos://artifacts/<uuid>`` — artifact-store lookup scoped to the
  caller's ``session_id``.
* ``kaos://<scheme>/<path>`` — VFS read scoped to ``context.session_id``.
* relative path that exists in the session VFS — VFS read.
* absolute filesystem path — direct disk read (CLI / tests / trusted
  callers).

Without that routing, files uploaded into ``KaosRuntime.vfs`` by a UI
host (e.g. ``kaos-ui``'s single-user-chat SPA) are invisible to the
tools because raw ``Path(p).exists()`` resolves against the CWD, not
the session VFS. See
``kaos-modules/docs/plans/vfs-blind-tools-audit-and-fix-plan.md`` for
the production post-mortem.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from kaos_core import KaosContext, KaosRuntime, KaosTool, ToolMetadata, ToolResult
from kaos_core.path_resolver import ResolvedOrigin
from kaos_core.types.annotations import ToolAnnotations
from kaos_core.types.enums import ToolCapability, ToolCategory
from kaos_core.types.parameters import ParameterSchema

from kaos_office._path_resolver import (
    InputPathResolutionError,
    ResolvedInput,
    resolve_office_input,
)

_MODULE = "kaos-office"
_VERSION = "0.1.0"

# All Office read tools are read-only, idempotent, and local-only.
_OFFICE_ANNOTATIONS = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)

# Write tools produce new files and register artifacts. Not idempotent — same
# input yields a new artifact id each call. `destructiveHint=False` because we
# refuse silent overwrites; callers must opt in via `force=True`.
_OFFICE_WRITE_ANNOTATIONS = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=True,
)

# Shared path-parameter description fragments. The "what to pass" line
# is what tells the agent that artifact URIs and VFS paths are
# first-class inputs — without it, agents default to absolute
# filesystem paths and miss SPA-uploaded files entirely.
_DOCX_PATH_DESC = (
    "Path to the DOCX file. Accepts an absolute filesystem path, a "
    "kaos://artifacts/<id> URI returned by a previous tool, or a "
    "relative path that resolves inside the session VFS (e.g. files "
    "uploaded through the host UI)."
)
_PPTX_PATH_DESC = (
    "Path to the PPTX file. Accepts an absolute filesystem path, a "
    "kaos://artifacts/<id> URI returned by a previous tool, or a "
    "relative path that resolves inside the session VFS (e.g. files "
    "uploaded through the host UI)."
)
_XLSX_PATH_DESC = (
    "Path to the XLSX file. Accepts an absolute filesystem path, a "
    "kaos://artifacts/<id> URI returned by a previous tool, or a "
    "relative path that resolves inside the session VFS (e.g. files "
    "uploaded through the host UI)."
)


def _origin_extras(resolved: ResolvedInput) -> dict[str, Any]:
    """Return structured-content fields that thread provenance from an artifact input.

    When the input originated in the artifact store, the SPA (or any
    downstream tool that chains on the result) should see the *original*
    artifact id / body URI rather than the random temp-file path the
    resolver materialised. Parse-* tools call this and merge the
    returned dict into their response so the ArtifactCard renders the
    upload the user actually attached.

    Returns an empty dict for VFS / filesystem inputs — those have no
    artifact id of their own and the Parse-* tools register a fresh
    derived artifact for them via ``store_document`` / ``store_tabular``.
    """
    if resolved.origin is not ResolvedOrigin.ARTIFACT:
        return {}
    extras: dict[str, Any] = {}
    if resolved.artifact_id is not None:
        extras["source_artifact_id"] = resolved.artifact_id
    if resolved.body_uri is not None:
        extras["source_body_uri"] = resolved.body_uri
    return extras


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
                    description=_DOCX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="docx") as resolved:
                try:
                    from kaos_office.docx.reader import parse_docx

                    doc = parse_docx(resolved.path)
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

                    name = resolved.path.stem or "document"
                    manifest = await store_document(doc, context.runtime, context, name=name)
                    summary = document_to_summary(doc, max_length=500)
                    outline = document_outline(doc)
                    view = DocumentView(doc)

                    structured: dict[str, Any] = {
                        "artifact_id": manifest.artifact_id,
                        "title": doc.metadata.title,
                        "block_count": len(doc.body),
                        "has_sections": view.has_sections,
                        "outline": outline[:10],
                        "section_count": len(view.flat_sections),
                        "body_uri": manifest.body_uri,
                        "sections_uri": f"kaos://content/{manifest.artifact_id}/sections",
                    }
                    structured.update(_origin_extras(resolved))
                    return manifest.to_tool_result(
                        summary=summary,
                        structured_content=structured,
                    )

                # No runtime — return inline summary
                from kaos_content.serializers.text import serialize_text

                text = serialize_text(doc)
                blocks = len(doc.body)
                return ToolResult.create_success(
                    f"Parsed {blocks} blocks from {resolved.path.name}.\n\n{text[:2000]}"
                )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_DOCX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="docx") as resolved:
                try:
                    from kaos_content.serializers.text import serialize_text

                    from kaos_office.docx.reader import parse_docx

                    doc = parse_docx(resolved.path)
                    text = serialize_text(doc)
                    return ToolResult.create_success(text)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to extract text: {exc}. "
                        "Try kaos-office-parse-docx for more detailed error info."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_DOCX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="docx") as resolved:
                try:
                    from kaos_content.serializers.markdown import serialize_markdown

                    from kaos_office.docx.reader import parse_docx

                    doc = parse_docx(resolved.path)
                    md = serialize_markdown(doc)
                    return ToolResult.create_success(md)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to extract markdown: {exc}. "
                        "Try kaos-office-get-text for plain text extraction."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_DOCX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="docx") as resolved:
                try:
                    from kaos_office.docx.metadata import DocxMetadata
                    from kaos_office.opc.package import OPCPackage

                    with OPCPackage.open(resolved.path) as pkg:
                        core_xml = (
                            pkg.read_part("docProps/core.xml")
                            if pkg.has_part("docProps/core.xml")
                            else None
                        )
                        app_xml = (
                            pkg.read_part("docProps/app.xml")
                            if pkg.has_part("docProps/app.xml")
                            else None
                        )
                        meta = DocxMetadata.from_xml(core_xml, app_xml)

                    meta_dict = meta.to_dict()
                    title = meta_dict.get("title", resolved.path.name)
                    summary = f"Metadata for {title}"
                    return ToolResult.create_success(output=meta_dict, summary=summary)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to extract metadata: {exc}. "
                        "The file may not be a valid DOCX. Check that it opens in Word "
                        "or LibreOffice."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_DOCX_PATH_DESC,
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

        try:
            async with resolve_office_input(path_str, context, format="docx") as resolved:
                try:
                    from kaos_content.search import search_document

                    from kaos_office.docx.reader import parse_docx

                    doc = parse_docx(resolved.path)
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
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_PPTX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="pptx") as resolved:
                try:
                    from kaos_office.pptx.reader import parse_pptx

                    doc = parse_pptx(resolved.path)
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

                    name = resolved.path.stem or "presentation"
                    manifest = await store_document(doc, context.runtime, context, name=name)
                    summary = document_to_summary(doc, max_length=500)
                    outline = document_outline(doc)
                    view = DocumentView(doc)

                    structured: dict[str, Any] = {
                        "artifact_id": manifest.artifact_id,
                        "title": doc.metadata.title,
                        "slide_count": len(doc.body),
                        "block_count": len(doc.body),
                        "has_sections": view.has_sections,
                        "outline": outline[:10],
                        "section_count": len(view.flat_sections),
                        "body_uri": manifest.body_uri,
                        "sections_uri": f"kaos://content/{manifest.artifact_id}/sections",
                    }
                    structured.update(_origin_extras(resolved))
                    return manifest.to_tool_result(
                        summary=summary,
                        structured_content=structured,
                    )

                # No runtime — return inline summary
                from kaos_content.serializers.text import serialize_text

                text = serialize_text(doc)
                blocks = len(doc.body)
                return ToolResult.create_success(
                    f"Parsed {blocks} slides from {resolved.path.name}.\n\n{text[:2000]}"
                )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_PPTX_PATH_DESC,
                ),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        path_str = inputs.get("path", "")
        try:
            async with resolve_office_input(path_str, context, format="pptx") as resolved:
                try:
                    from kaos_office.pptx.reader import list_slides

                    slides = list_slides(resolved.path)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to list slides: {exc}. "
                        "The file may be corrupted. Try opening it in PowerPoint or LibreOffice."
                    )

                summary = f"Found {len(slides)} slides"
                return ToolResult.create_success(output={"slides": slides}, summary=summary)
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_PPTX_PATH_DESC,
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

        try:
            async with resolve_office_input(path_str, context, format="pptx") as resolved:
                try:
                    from kaos_office.pptx.reader import get_slide_text

                    text = get_slide_text(resolved.path, slide_number)
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
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_PPTX_PATH_DESC,
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

        try:
            async with resolve_office_input(path_str, context, format="pptx") as resolved:
                try:
                    from kaos_content.search import search_document

                    from kaos_office.pptx.reader import parse_pptx

                    doc = parse_pptx(resolved.path)
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
                        "Try kaos-office-parse-pptx to verify the presentation has "
                        "extractable content."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                    description=_PPTX_PATH_DESC,
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

        try:
            async with resolve_office_input(path_str, context, format="pptx") as resolved:
                try:
                    from kaos_office.pptx.reader import get_slide_notes

                    notes = get_slide_notes(resolved.path, slide_number)
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
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


# Shared XLSX-tool error copy. The native lxml reader needs no extras;
# the install hint stays so the message is consistent across the four
# XLSX tools when an optional engine (calamine / openpyxl) is missing.
_XLSX_IMPORT_ERROR = (
    "XLSX optional dependency missing. The default native reader needs no extras; "
    "install `kaos-office[xlsx-calamine]` for the Rust fast-path, "
    "`kaos-office[xlsx-formulas]` for openpyxl-backed formula extraction, "
    "or the aggregate `kaos-office[xlsx]` for both. "
    'Alternative: drop `engine="calamine"` and `include_formulas=True` to '
    "fall back to the native lxml reader, which has no extra deps."
)


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
                ParameterSchema(name="path", type="string", description=_XLSX_PATH_DESC),
                ParameterSchema(
                    name="sheets",
                    type="array",
                    description="Sheet names to extract. Default: all visible.",
                    required=False,
                    constraints={"items": {"type": "string"}},
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
        try:
            async with resolve_office_input(path_str, context, format="xlsx") as resolved:
                try:
                    from kaos_content.serializers.tabular import serialize_tabular_summary

                    from kaos_office.xlsx.reader import parse_xlsx

                    doc = parse_xlsx(
                        resolved.path,
                        sheets=inputs.get("sheets"),
                        header_row=inputs.get("header_row", 0),
                    )
                    summary = serialize_tabular_summary(doc)

                    if context is not None and context.runtime is not None:
                        from kaos_content.artifacts import store_tabular

                        name = resolved.path.stem or "workbook"
                        manifest = await store_tabular(doc, context.runtime, context, name=name)
                        structured: dict[str, Any] = {
                            "artifact_id": manifest.artifact_id,
                            "table_count": len(doc.tables),
                            "tables": [
                                {"name": t.name, "row_count": t.row_count} for t in doc.tables
                            ],
                        }
                        structured.update(_origin_extras(resolved))
                        return manifest.to_tool_result(
                            summary=summary,
                            structured_content=structured,
                        )
                    return ToolResult.create_text(summary)
                except ImportError:
                    return ToolResult.create_error(_XLSX_IMPORT_ERROR)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"XLSX extraction failed for {resolved.path}: {exc}. "
                        "The file may be password-protected, corrupted, or use an unsupported "
                        "feature. Try kaos-office-xlsx-metadata to inspect the workbook, or "
                        "kaos-office-list-sheets-xlsx to verify it opens at all."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                ParameterSchema(name="path", type="string", description=_XLSX_PATH_DESC),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        try:
            async with resolve_office_input(inputs["path"], context, format="xlsx") as resolved:
                try:
                    from kaos_office.xlsx.reader import list_sheets

                    return ToolResult.create_success(output={"sheets": list_sheets(resolved.path)})
                except ImportError:
                    return ToolResult.create_error(_XLSX_IMPORT_ERROR)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to list sheets in {resolved.path}: {exc}. "
                        "Verify the file is a valid XLSX (not XLS, CSV, or a renamed file). "
                        "Try kaos-office-xlsx-metadata if you only need workbook properties, "
                        "or open the file directly to confirm it isn't corrupted or "
                        "password-protected."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


class GetSheetXlsxTool(KaosTool):
    """Get a single sheet's data as TSV."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-get-sheet-xlsx",
            display_name="Get XLSX Sheet",
            description=(
                "Extract a single sheet's data. Use list-sheets-xlsx to see available sheets."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.EXTRACT,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_ANNOTATIONS,
            input_schema=[
                ParameterSchema(name="path", type="string", description=_XLSX_PATH_DESC),
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
        try:
            async with resolve_office_input(inputs["path"], context, format="xlsx") as resolved:
                try:
                    from kaos_content.serializers.tabular import serialize_tsv

                    from kaos_office.xlsx.reader import parse_xlsx

                    doc = parse_xlsx(
                        resolved.path,
                        sheets=[inputs["sheet"]],
                        max_rows=inputs.get("max_rows", 100),
                    )
                    if not doc.tables:
                        return ToolResult.create_error(
                            f"Sheet '{inputs['sheet']}' not found in {resolved.path}. "
                            "Sheet names are case-sensitive and may include trailing whitespace. "
                            "Call kaos-office-list-sheets-xlsx first to see the exact names "
                            "available."
                        )
                    return ToolResult.create_text(serialize_tsv(doc.tables[0]))
                except ImportError:
                    return ToolResult.create_error(_XLSX_IMPORT_ERROR)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to extract sheet '{inputs['sheet']}' from {resolved.path}: {exc}. "
                        "If the sheet uses array formulas or external links, the native reader "
                        "may not support them — retry with kaos-office-parse-xlsx "
                        '(engine="calamine") or use kaos-office-list-sheets-xlsx to confirm '
                        "the sheet exists."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


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
                ParameterSchema(name="path", type="string", description=_XLSX_PATH_DESC),
            ],
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        try:
            async with resolve_office_input(inputs["path"], context, format="xlsx") as resolved:
                try:
                    from kaos_content.artifacts import tabular_summary

                    from kaos_office.xlsx.reader import parse_xlsx

                    doc = parse_xlsx(resolved.path)
                    return ToolResult.create_success(output=tabular_summary(doc))
                except ImportError:
                    return ToolResult.create_error(_XLSX_IMPORT_ERROR)
                except Exception as exc:
                    return ToolResult.create_error(
                        f"Failed to read XLSX metadata for {resolved.path}: {exc}. "
                        "Confirm the file is a valid .xlsx workbook (not .xls / .xlsm with "
                        "macros). Try kaos-office-list-sheets-xlsx for a lighter probe that "
                        "returns just the sheet inventory without parsing every cell."
                    )
        except InputPathResolutionError as exc:
            return ToolResult.create_error(exc.to_agent_message())


# ---------------------------------------------------------------------------
# Writer tools — ContentDocument / TabularDocument -> DOCX / PPTX / XLSX
# ---------------------------------------------------------------------------


async def _resolve_content_document(
    inputs: dict[str, Any], context: KaosContext | None
) -> tuple[Any | None, str | None]:
    """Resolve a ContentDocument from inline JSON or an artifact id.

    Returns ``(doc, None)`` on success or ``(None, error_message)`` on failure.
    """
    from kaos_content.model.document import ContentDocument

    inline_raw = inputs.get("document_json")
    artifact_id = inputs.get("document_id")

    if not inline_raw and not artifact_id:
        return None, (
            "Missing document: provide either `document_json` (inline JSON) or "
            "`document_id` (an artifact id from kaos-office-parse-docx / parse-pptx)."
        )

    if artifact_id:
        if context is None or context.runtime is None:
            return None, "`document_id` requires a KaosRuntime; pass `document_json` instead."
        from kaos_content.artifacts import load_document

        try:
            return await load_document(artifact_id, context.runtime), None
        except Exception as exc:  # artifact missing, JSON corrupt, wrong type
            return None, f"Failed to load document artifact {artifact_id!r}: {exc}."

    if not isinstance(inline_raw, str | bytes | bytearray):
        return None, "`document_json` must be a JSON string."
    try:
        return ContentDocument.model_validate_json(inline_raw), None
    except Exception as exc:
        return None, (
            f"`document_json` is not a valid ContentDocument: {exc}. "
            "Use the JSON body from kaos-office-parse-docx or kaos-office-parse-pptx."
        )


async def _resolve_tabular_document(
    inputs: dict[str, Any], context: KaosContext | None
) -> tuple[Any | None, str | None]:
    """Resolve a TabularDocument from inline JSON or an artifact id."""
    from kaos_content.model.tabular import TabularDocument

    inline_raw = inputs.get("document_json")
    artifact_id = inputs.get("document_id")

    if not inline_raw and not artifact_id:
        return None, (
            "Missing document: provide either `document_json` (inline JSON) or "
            "`document_id` (artifact id from kaos-office-parse-xlsx)."
        )

    if artifact_id:
        if context is None or context.runtime is None:
            return None, "`document_id` requires a KaosRuntime; pass `document_json` instead."
        from kaos_content.artifacts import load_tabular

        try:
            return await load_tabular(artifact_id, context.runtime), None
        except Exception as exc:
            return None, f"Failed to load tabular artifact {artifact_id!r}: {exc}."

    if not isinstance(inline_raw, str | bytes | bytearray):
        return None, "`document_json` must be a JSON string."
    try:
        return TabularDocument.model_validate_json(inline_raw), None
    except Exception as exc:
        return None, (
            f"`document_json` is not a valid TabularDocument: {exc}. "
            "Use the JSON body from kaos-office-parse-xlsx."
        )


def _check_output_path(output_path: str, force: bool) -> tuple[Path | None, str | None]:
    """Validate an output path. Refuses to overwrite unless ``force`` is set."""
    p = Path(output_path)
    if p.exists() and not force:
        return None, (
            f"Refusing to overwrite existing file: {p}. "
            "Pass `force=true` to overwrite, or choose a different `output_path`."
        )
    return p, None


_WRITER_INPUT_SCHEMA_COMMON: list[ParameterSchema] = [
    ParameterSchema(
        name="output_path",
        type="string",
        description="Filesystem path where the output file will be written.",
    ),
    ParameterSchema(
        name="document_json",
        type="string",
        description="Full JSON body of the document (inline).",
        required=False,
    ),
    ParameterSchema(
        name="document_id",
        type="string",
        description="Artifact id produced by a parse tool. Loaded from VFS.",
        required=False,
    ),
    ParameterSchema(
        name="force",
        type="boolean",
        description="Overwrite output_path if it already exists.",
        required=False,
        default=False,
    ),
]


_FORMAT_MIME: dict[str, str] = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}


async def _register_output_as_artifact(
    written_path: Path,
    format_name: str,
    context: KaosContext | None,
) -> dict[str, Any] | None:
    """Copy the just-written file into the VFS and register as an artifact.

    Returns ``{"artifact_id": ..., "body_uri": ..., "manifest_uri": ...}`` on
    success, or ``None`` when no runtime is available or registration fails.
    Registration failures are silent (the local file is the source of truth);
    the returned dict is advisory metadata for chaining into other MCP tools.
    """
    if context is None or context.runtime is None:
        return None
    try:
        from kaos_core.artifacts.models import ArtifactRole

        data = written_path.read_bytes()
        vfs_path = f"office_output/{written_path.name}"
        ctx_path = context.get_vfs_path(vfs_path)
        await ctx_path.write_bytes(data)
        manifest = await context.runtime.artifacts.create_from_path(
            vfs_path,
            context_id=context.session_id,
            session_id=context.session_id,
            name=written_path.stem,
            mime_type=_FORMAT_MIME.get(format_name),
            role=ArtifactRole.BODY,
            provenance={"format": format_name, "source_path": str(written_path)},
        )
    except Exception:
        logger = __import__("kaos_core.logging", fromlist=["get_logger"]).get_logger(__name__)
        logger.debug("Failed to register output as artifact", exc_info=True)
        return None
    return {
        "artifact_id": manifest.artifact_id,
        "body_uri": manifest.body_uri,
        "manifest_uri": f"kaos://artifacts/{manifest.artifact_id}/manifest",
    }


def _write_success_result(
    written_path: Path,
    format_name: str,
    extra_structured: dict[str, Any] | None = None,
) -> ToolResult:
    size = written_path.stat().st_size
    structured: dict[str, Any] = {
        "path": str(written_path),
        "format": format_name,
        "size_bytes": size,
    }
    if extra_structured:
        structured.update(extra_structured)
    summary = f"Wrote {size} bytes to {written_path} ({format_name})."
    return ToolResult.create_success(output=structured, summary=summary)


class WriteDocxTool(KaosTool):
    """Write a ContentDocument to a DOCX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-write-docx",
            display_name="Write DOCX",
            description=(
                "Serialize a ContentDocument (inline JSON or artifact id) to a DOCX file. "
                "Call kaos-office-parse-docx on an existing file to obtain a starting JSON body."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.TRANSFORM,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_WRITE_ANNOTATIONS,
            input_schema=_WRITER_INPUT_SCHEMA_COMMON,
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        doc, err = await _resolve_content_document(inputs, context)
        if err is not None or doc is None:
            return ToolResult.create_error(err or "Failed to resolve ContentDocument.")
        out, err = _check_output_path(inputs.get("output_path", ""), bool(inputs.get("force")))
        if err is not None or out is None:
            return ToolResult.create_error(err or "Missing output_path.")

        try:
            from kaos_office.docx.writer import write_docx

            write_docx(doc, out)
        except Exception as exc:
            return ToolResult.create_error(
                f"DOCX write failed: {exc}. "
                "Verify the ContentDocument has a `body` and a writable `output_path`."
            )

        extra: dict[str, Any] = {"block_count": len(doc.body)}
        artifact_meta = await _register_output_as_artifact(out, "docx", context)
        if artifact_meta is not None:
            extra.update(artifact_meta)
        return _write_success_result(out, "docx", extra)


class WritePptxTool(KaosTool):
    """Write a ContentDocument to a PPTX file."""

    @property
    def metadata(self) -> ToolMetadata:
        schema = list(_WRITER_INPUT_SCHEMA_COMMON)
        schema.append(
            ParameterSchema(
                name="template_path",
                type="string",
                description=(
                    "Optional path to a .pptx file to use as a template. "
                    "Accepts an absolute filesystem path, a kaos://artifacts/<id> "
                    "URI, or a session-VFS path."
                ),
                required=False,
            )
        )
        return ToolMetadata(
            name="kaos-office-write-pptx",
            display_name="Write PPTX",
            description=(
                "Serialize a ContentDocument (inline JSON or artifact id) to a PPTX file. "
                "Each `Heading(depth=1)` starts a new slide; tables get their own slide. "
                "Use an optional `template_path` to apply a branded theme."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.TRANSFORM,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_WRITE_ANNOTATIONS,
            input_schema=schema,
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        doc, err = await _resolve_content_document(inputs, context)
        if err is not None or doc is None:
            return ToolResult.create_error(err or "Failed to resolve ContentDocument.")
        out, err = _check_output_path(inputs.get("output_path", ""), bool(inputs.get("force")))
        if err is not None or out is None:
            return ToolResult.create_error(err or "Missing output_path.")

        template_path = inputs.get("template_path") or None

        # Resolve the optional template through the same VFS-aware helper so
        # SPA-uploaded .pptx templates work the same as filesystem paths. The
        # writer call happens inside the `async with` so the resolver's temp
        # file stays alive for the duration of `write_pptx`.
        if template_path is not None:
            try:
                async with resolve_office_input(
                    template_path, context, format="pptx"
                ) as resolved_template:
                    return await self._write(
                        doc=doc,
                        out=out,
                        template_path=resolved_template.path,
                        template_input=template_path,
                        context=context,
                    )
            except InputPathResolutionError as exc:
                return ToolResult.create_error(
                    f"Template not found: {template_path}. "
                    "Omit `template_path` to use the default blank deck. "
                    f"(resolver: {exc.to_agent_message()})"
                )

        return await self._write(
            doc=doc,
            out=out,
            template_path=None,
            template_input=None,
            context=context,
        )

    async def _write(
        self,
        *,
        doc: Any,
        out: Path,
        template_path: Path | None,
        template_input: str | None,
        context: KaosContext | None,
    ) -> ToolResult:
        try:
            from kaos_office.pptx.writer import write_pptx

            write_pptx(doc, out, template=template_path)
        except ImportError:
            return ToolResult.create_error(
                "PPTX writing requires python-pptx. pip install kaos-office[pptx]"
            )
        except Exception as exc:
            return ToolResult.create_error(
                f"PPTX write failed: {exc}. "
                "Verify the ContentDocument has a `body` and any `template_path` is valid."
            )

        extra: dict[str, Any] = {
            "block_count": len(doc.body),
            "template_path": template_input,
        }
        artifact_meta = await _register_output_as_artifact(out, "pptx", context)
        if artifact_meta is not None:
            extra.update(artifact_meta)
        return _write_success_result(out, "pptx", extra)


class WriteXlsxTool(KaosTool):
    """Write a TabularDocument to an XLSX file."""

    @property
    def metadata(self) -> ToolMetadata:
        return ToolMetadata(
            name="kaos-office-write-xlsx",
            display_name="Write XLSX",
            description=(
                "Serialize a TabularDocument (inline JSON or artifact id) to an XLSX file. "
                "Each `Table` becomes a worksheet. Call kaos-office-parse-xlsx to obtain "
                "a starting JSON body."
            ),
            category=ToolCategory.DOCUMENT,
            capability=ToolCapability.TRANSFORM,
            module_name=_MODULE,
            version=_VERSION,
            annotations=_OFFICE_WRITE_ANNOTATIONS,
            input_schema=_WRITER_INPUT_SCHEMA_COMMON,
        )

    async def execute(
        self, inputs: dict[str, Any], context: KaosContext | None = None
    ) -> ToolResult:
        doc, err = await _resolve_tabular_document(inputs, context)
        if err is not None or doc is None:
            return ToolResult.create_error(err or "Failed to resolve TabularDocument.")
        out, err = _check_output_path(inputs.get("output_path", ""), bool(inputs.get("force")))
        if err is not None or out is None:
            return ToolResult.create_error(err or "Missing output_path.")

        try:
            from kaos_office.xlsx.writer import write_xlsx

            write_xlsx(doc, out)
        except Exception as exc:
            return ToolResult.create_error(
                f"XLSX write failed: {exc}. "
                "Verify the TabularDocument has `tables` and a writable `output_path`."
            )

        extra: dict[str, Any] = {"table_count": len(doc.tables)}
        artifact_meta = await _register_output_as_artifact(out, "xlsx", context)
        if artifact_meta is not None:
            extra.update(artifact_meta)
        return _write_success_result(out, "xlsx", extra)


def register_office_documents_tools(runtime: KaosRuntime) -> int:
    """Register the 14 read-only Office MCP tools.

    DOCX, PPTX, and XLSX parsers, metadata inspectors, listers, and
    BM25 searchers — every tool that takes an Office document and
    returns text, structure, metadata, or a derived view *without*
    producing a new artifact. Pins the SessionToolSet ``documents``
    group entry point.
    """
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


def register_office_authoring_tools(runtime: KaosRuntime) -> int:
    """Register the 3 Office authoring (writer) MCP tools.

    DOCX, PPTX, and XLSX writers — they serialize a
    ``ContentDocument`` / ``TabularDocument`` to a new Office
    artifact. Pins the SessionToolSet ``authoring`` group entry
    point: denied by default at the ceiling and opted into
    per-session for drafting workflows.
    """
    tools: list[KaosTool] = [
        WriteDocxTool(),
        WritePptxTool(),
        WriteXlsxTool(),
    ]
    for tool in tools:
        runtime.tools.register_tool(tool)
    return len(tools)


def register_office_tools(runtime: KaosRuntime) -> int:
    """Register all Office MCP tools with a runtime.

    Backward-compatible union of
    :func:`register_office_documents_tools` (14 read-only) and
    :func:`register_office_authoring_tools` (3 writers). Existing
    callers continue to see the same 17 tools with the same names
    and schemas.
    """
    count = register_office_documents_tools(runtime)
    count += register_office_authoring_tools(runtime)
    return count
