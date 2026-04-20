"""DOCX Reader — parse_docx() → ContentDocument.

Two-pass architecture:
  Pass 1: Load metadata (styles, numbering, relationships, document metadata)
  Pass 2: Walk document body with tag dispatch → DocumentBuilder

Supports: paragraphs, headings, lists, tables, images, hyperlinks,
footnotes, endnotes, comments, track changes (accept/skip), page breaks.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from kaos_content import ContentDocument
from kaos_content.builders import DocumentBuilder
from kaos_content.model.annotation import AnnotationType
from kaos_content.model.attr import Attr
from kaos_content.model.blocks import Table
from kaos_content.model.inlines import (
    Emphasis,
    Image,
    Inline,
    LineBreak,
    Link,
    Span,
    Strikethrough,
    Strong,
    Subscript,
    Superscript,
    Text,
    Underline,
)
from kaos_content.model.metadata import PageSetup, Section, SectionBreakType
from kaos_content.model.table import Cell, Row, TableSection
from kaos_core.logging import get_logger
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.docx.metadata import DocxMetadata
from kaos_office.docx.numbering import NumberingResolver
from kaos_office.docx.styles import StyleResolver
from kaos_office.ooxml.namespace import (
    DOCX_MIME_TYPE,
    R_EMBED,
    R_ID,
    R_ID_ATTR,
    RT_COMMENTS,
    RT_ENDNOTES,
    RT_FOOTER,
    RT_FOOTNOTES,
    RT_HEADER,
    RT_OFFICE_DOCUMENT,
    W_ANCHOR,
    W_B,
    W_BODY,
    W_BOOKMARK_START,
    W_BR,
    W_COMMENT,
    W_DEL,
    W_DEL_TEXT,
    W_DRAWING,
    W_ENDNOTE,
    W_ENDNOTE_REFERENCE,
    W_ENDNOTES,
    W_FOOTER_REFERENCE,
    W_FOOTNOTE,
    W_FOOTNOTE_REFERENCE,
    W_FOOTNOTES,
    W_GRIDCOL,
    W_GRIDSPAN,
    W_HEADER_REFERENCE,
    W_HYPERLINK,
    W_I,
    W_ID,
    W_ILVL,
    W_INS,
    W_MOVE_FROM,
    W_MOVE_TO,
    W_NUMID,
    W_NUMPR,
    W_P,
    W_PGMAR,
    W_PGSZ,
    W_PPR,
    W_PSTYLE,
    W_R,
    W_RPR,
    W_SDT,
    W_SDTCONTENT,
    W_SECTPR,
    W_STRIKE,
    W_T,
    W_TAB,
    W_TBL,
    W_TBLGRID,
    W_TC,
    W_TCPR,
    W_TR,
    W_TYPE,
    W_U,
    W_VAL,
    W_VERTALING,
    W_VMERGE,
    WP,
    A,
    W,
    emu_to_pt,
    qn,
    twips_to_pt,
)
from kaos_office.opc.package import OPCPackage

logger = get_logger(__name__)

_EXTRACTOR = "kaos-office/docx"


@dataclass  # Mutable: item_open toggled during list processing
class ListState:
    """Track open list state for proper begin/end nesting."""

    num_id: str
    ilvl: int
    ordered: bool
    item_open: bool = False  # Whether a list item is currently open


@dataclass  # Mutable: accumulates state (bookmarks, list_stack) during document parsing
class ParseContext:
    """Mutable state passed through the parse tree."""

    builder: DocumentBuilder
    styles: StyleResolver
    numbering: NumberingResolver
    rels: dict[str, str]  # rId → target path
    rels_external: dict[str, str]  # rId → external URL
    source_uri: str

    # Footnote/endnote content (parsed from separate XML parts)
    footnotes: dict[str, list[etree._Element]] = field(default_factory=dict)
    endnotes: dict[str, list[etree._Element]] = field(default_factory=dict)

    # Comment content
    comments: dict[str, dict[str, str]] = field(default_factory=dict)

    # List state tracking
    list_stack: list[ListState] = field(default_factory=list)

    # Bookmark tracking for annotations
    bookmarks: dict[str, str] = field(default_factory=dict)  # id → name

    # Revision tracking mode — when True, w:ins/w:del/w:moveFrom/w:moveTo
    # content is wrapped in Span/Div with ``rev-*`` classes and a
    # TRACKED_CHANGE annotation is emitted with metadata.
    track_changes: bool = False

    # Phase 4C: sectPr elements discovered during body walk, captured at
    # the point they close a section so the final block index lines up
    # with whatever the builder actually emitted. Each entry is
    # ``(block_count_at_boundary, sectPr_element)``. Consumed after the
    # walk by :func:`_build_sections`.
    pending_sections: list[tuple[int, etree._Element]] = field(default_factory=list)

    # Phase 6.1: OPC package + URI policy for image extraction. ``pkg``
    # lets ``_extract_image`` read the actual media bytes via
    # ``pkg.read_part(target)``; ``image_src_builder`` converts those
    # bytes + format + 1-based index into the ``Image.src`` value. The
    # default builder inlines a ``data:image/<fmt>;base64,...`` URI so
    # reader-to-writer round-trip is lossless (writer.py Phase 4B.2
    # accepts ``data:`` URIs). Callers can pass their own builder to
    # return bare logical URIs while collecting bytes side-channel —
    # same contract as ``kaos-pdf.extract_pdf.image_src_builder``.
    pkg: Any = None  # OPCPackage | None
    image_src_builder: Any = None  # Callable[[bytes, str, int], str] | None
    image_counter: int = 0


def parse_docx(
    path: str | Path,
    *,
    track_changes: bool = False,
    image_src_builder: Callable[[bytes, str, int], str] | None = None,
) -> ContentDocument:
    """Parse a DOCX file into a ContentDocument.

    Args:
        path: Path to the .docx file.
        track_changes: When True, preserve tracked changes (``w:ins``,
            ``w:del``, ``w:moveFrom``, ``w:moveTo``) by wrapping content
            in ``Span`` / ``Div`` with ``rev-ins`` / ``rev-del`` /
            ``rev-move-from`` / ``rev-move-to`` classes and emitting
            ``AnnotationType.TRACKED_CHANGE`` annotations with the
            author / date / revision-id metadata. When False (default),
            the legacy "accept insertions, skip deletions" behavior
            applies and all revision metadata is discarded.
        image_src_builder: Callable turning an embedded image's
            ``(bytes, fmt, index_1_based)`` into the ``Image.src`` URI
            string. Default is :func:`_inline_data_uri` which emits
            ``data:image/<fmt>;base64,...`` so reader-to-writer
            round-trip is lossless (the DOCX writer's
            ``_decode_image_src`` accepts ``data:`` URIs). Callers who
            need an artifact store or a bare logical URI pass their own
            builder — e.g. write bytes to the VFS and return
            ``kaos://artifacts/{id}/body``, or return ``docx://...``
            while collecting bytes in a side-channel dict.

    Returns:
        ContentDocument with the extracted content.

    Raises:
        OPCPackageError: If the file cannot be opened or parsed.
        OPCSecurityError: If the file fails security validation.
    """
    path = Path(path)
    source_uri = path.as_uri()

    with OPCPackage.open(path) as pkg:
        # --- Pass 1: Load metadata ---
        root_rels = pkg.relationships("/")
        doc_target = root_rels.first_target(RT_OFFICE_DOCUMENT)
        if doc_target is None:
            raise ValueError(
                f"No main document part found in {path}. The file may not be a valid DOCX document."
            )

        # Resolve paths relative to the document part's directory
        doc_dir = str(Path(doc_target).parent)
        if doc_dir == ".":
            doc_dir = ""

        doc_rels = pkg.relationships(doc_target)

        # Build rId → target maps
        rels_map: dict[str, str] = {}
        rels_external: dict[str, str] = {}
        for rel_id in _iter_rel_ids(doc_rels):
            rel = doc_rels.get(rel_id)
            if rel is None:
                continue
            if rel.external:
                rels_external[rel_id] = rel.target
            else:
                # Target is relative to the document part's directory
                if doc_dir:
                    rels_map[rel_id] = f"{doc_dir}/{rel.target}"
                else:
                    rels_map[rel_id] = rel.target

        # Parse styles
        styles_part = _resolve_part(doc_rels, pkg, doc_dir, "styles")
        styles = StyleResolver.from_xml(styles_part)

        # Parse numbering
        numbering_part = _resolve_part(doc_rels, pkg, doc_dir, "numbering")
        numbering = NumberingResolver.from_xml(numbering_part)

        # Parse document metadata
        core_xml = pkg.read_part("docProps/core.xml") if pkg.has_part("docProps/core.xml") else None
        app_xml = pkg.read_part("docProps/app.xml") if pkg.has_part("docProps/app.xml") else None
        docx_meta = DocxMetadata.from_xml(core_xml, app_xml)

        # Parse footnotes
        footnotes = _parse_footnotes(doc_rels, pkg, doc_dir, RT_FOOTNOTES, W_FOOTNOTES, W_FOOTNOTE)

        # Parse endnotes
        endnotes = _parse_footnotes(doc_rels, pkg, doc_dir, RT_ENDNOTES, W_ENDNOTES, W_ENDNOTE)

        # Parse comments
        comments = _parse_comments(doc_rels, pkg, doc_dir)

        # --- Pass 2: Walk document body ---
        doc_xml = pkg.read_xml(doc_target)
        body = doc_xml.find(W_BODY)
        if body is None:
            body = doc_xml  # Some documents have body as the root

        builder = DocumentBuilder(title=docx_meta.title)
        # ``set_source`` attaches provenance to blocks; ``set_metadata`` also
        # populates ``document.metadata.source`` so multi-document corpora
        # (e.g. ``kaos_content.corpus.ContentDocumentCorpus``,
        # ``kaos_ml_core.Corpus.from_documents``) thread ``doc_uri`` without
        # an explicit constructor kwarg. Mirrors ``extract_pdf``
        # (kaos-pdf/extract.py:285) and ``parse_plain_text``
        # (kaos-content/parsers/plain.py).
        from kaos_content.model.attr import SourceRef

        builder.set_source(uri=source_uri, mime_type=DOCX_MIME_TYPE)
        builder.set_metadata(source=SourceRef(uri=source_uri, mime_type=DOCX_MIME_TYPE))
        if docx_meta.creator:
            builder.set_metadata(authors=(docx_meta.creator,))
        if docx_meta.created:
            builder.set_metadata(date=docx_meta.created)

        ctx = ParseContext(
            builder=builder,
            styles=styles,
            numbering=numbering,
            rels=rels_map,
            rels_external=rels_external,
            source_uri=source_uri,
            footnotes=footnotes,
            endnotes=endnotes,
            comments=comments,
            track_changes=track_changes,
            pkg=pkg,
            image_src_builder=image_src_builder
            if image_src_builder is not None
            else _inline_data_uri,
        )

        for child in body:
            _process_body_child(child, ctx)

        # Flush any trailing open lists
        _flush_open_lists(ctx)

        # Add footnote/endnote content
        _add_footnotes(ctx)

        # Add comment annotations
        _add_comment_annotations(ctx)

        # Headers, footers, and page setup (Phase 4) — parsed here after
        # the main body walk so shared state (styles, rels) is available
        # and nothing in the body loop needs to know about them.
        headers, footers = _parse_headers_and_footers(doc_rels, pkg, doc_dir, body, ctx)
        for kind, blocks in headers.items():
            builder.set_header(kind, *blocks)
        for kind, blocks in footers.items():
            builder.set_footer(kind, *blocks)
        page_setup = _parse_page_setup(body)
        if page_setup is not None:
            builder.set_metadata(page_setup=page_setup)

        # Phase 4C: populate per-section layout from every <w:sectPr>
        # captured during the body walk. Empty tuple means "implicit
        # single section" and leaves doc.sections at its default.
        sections = _build_sections(ctx.pending_sections)
        if sections:
            builder.set_sections(sections)

        return builder.build()


# --------------------------------------------------------------------------
# Revision (tracked changes) helpers
# --------------------------------------------------------------------------

# Namespaced attribute names on revision elements (w:ins, w:del, w:moveFrom,
# w:moveTo). Cached at module load for O(1) access.
_W_AUTHOR_ATTR = qn(W, "author")
_W_DATE_ATTR = qn(W, "date")
_W_ID_ATTR = qn(W, "id")
_W_NAME_ATTR = qn(W, "name")

# Revision wrapper class names on Span/Div.
_REV_INS = "rev-ins"
_REV_DEL = "rev-del"
_REV_MOVE_FROM = "rev-move-from"
_REV_MOVE_TO = "rev-move-to"

# Map revision OOXML tag → (class, change_type)
_REV_TAG_MAP = {
    W_INS: (_REV_INS, "insertion"),
    W_DEL: (_REV_DEL, "deletion"),
    W_MOVE_FROM: (_REV_MOVE_FROM, "move_from"),
    W_MOVE_TO: (_REV_MOVE_TO, "move_to"),
}


def _revision_metadata(el: etree._Element) -> dict[str, str]:
    """Extract revision metadata (id, author, date, move name) from an element."""
    kv: dict[str, str] = {}
    rev_id = el.get(_W_ID_ATTR)
    if rev_id:
        kv["rev:id"] = rev_id
    author = el.get(_W_AUTHOR_ATTR)
    if author:
        kv["rev:author"] = author
    date = el.get(_W_DATE_ATTR)
    if date:
        kv["rev:date"] = date
    move_name = el.get(_W_NAME_ATTR)
    if move_name:
        kv["rev:move-name"] = move_name
    return kv


def _emit_revision_annotation(
    ctx: ParseContext,
    *,
    change_type: str,
    metadata: dict[str, str],
) -> None:
    """Emit a TRACKED_CHANGE annotation for a revision.

    Targets are intentionally empty at parse time; ``node_ref`` resolution
    via ``NodeIndex`` is a post-build concern. Downstream code can match
    annotations to Span/Div nodes by ``rev:id`` in ``Attr.kv``.
    """
    body: dict[str, object] = {"change_type": change_type}
    if "rev:id" in metadata:
        body["revision_id"] = metadata["rev:id"]
    if "rev:author" in metadata:
        body["author"] = metadata["rev:author"]
    if "rev:date" in metadata:
        body["date"] = metadata["rev:date"]
    if "rev:move-name" in metadata:
        body["move_name"] = metadata["rev:move-name"]
    ctx.builder.annotate(
        AnnotationType.TRACKED_CHANGE,
        targets=[],
        body=body,
    )


# --------------------------------------------------------------------------
# Body-level dispatch
# --------------------------------------------------------------------------


def _process_body_child(el: etree._Element, ctx: ParseContext) -> None:
    """Dispatch a top-level body child element."""
    tag = el.tag

    if tag == W_P:
        _handle_paragraph(el, ctx)
        # Phase 4C: a paragraph may carry <w:pPr><w:sectPr/></w:pPr>,
        # which closes a section at this block's position. Capture it
        # *after* the paragraph is emitted so the block count reflects
        # whatever _handle_paragraph produced (or didn't, for empty
        # paragraphs). Final body-direct sectPr is handled in the
        # W_SECTPR branch below.
        ppr = el.find(W_PPR)
        if ppr is not None:
            inner_sect = ppr.find(W_SECTPR)
            if inner_sect is not None:
                ctx.pending_sections.append((len(ctx.builder._blocks), inner_sect))
    elif tag == W_TBL:
        _flush_open_lists(ctx)
        _handle_table(el, ctx)
    elif tag == W_SECTPR:
        _flush_open_lists(ctx)
        # Body-direct sectPr is the final section's properties. Record
        # it at the current block count so _build_sections can close
        # the final section with end_block_index = len(body).
        ctx.pending_sections.append((len(ctx.builder._blocks), el))
    elif tag == W_SDT:
        # Phase 6.2: preserve the SDT wrapper as a Div(classes=("sdt",))
        # carrying the control's metadata, so the writer can re-emit the
        # <w:sdt> around the same content. Pre-6.2 we simply unwrapped
        # and lost tag / alias / lock / control-type on every round-trip.
        _flush_open_lists(ctx)
        sdt_attr = _parse_sdt_pr(el)
        sdt_content = el.find(W_SDTCONTENT)
        if sdt_content is None:
            return
        if sdt_attr:
            # Wrap inner content in a Div. The builder's begin_div/end
            # stack machinery handles nesting (including SDT-in-SDT).
            ctx.builder.begin_div(classes=("sdt",), kv=sdt_attr)
            for child in sdt_content:
                _process_body_child(child, ctx)
            ctx.builder.end()
        else:
            # No preservable sdtPr — fall back to legacy unwrap so the
            # content still lands in the AST.
            for child in sdt_content:
                _process_body_child(child, ctx)
    elif tag == W_BOOKMARK_START:
        bm_id = el.get(W_ID)
        bm_name = el.get(qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "name"))
        if bm_id and bm_name:
            ctx.bookmarks[bm_id] = bm_name
    elif tag in _REV_TAG_MAP:
        _handle_body_revision(el, ctx, tag)
    # Unknown elements silently skipped


def _handle_body_revision(el: etree._Element, ctx: ParseContext, tag: str) -> None:
    """Handle a block-level revision element (w:ins, w:del, w:moveFrom, w:moveTo).

    When ``ctx.track_changes`` is True, wrap the inner blocks in a ``Div``
    with ``rev-*`` classes and metadata, and emit a TRACKED_CHANGE
    annotation. Otherwise apply the legacy flatten behavior: include
    insertions and moveTo content, skip deletions and moveFrom content.
    """
    rev_class, change_type = _REV_TAG_MAP[tag]

    if not ctx.track_changes:
        # Legacy behavior: include ins/moveTo, skip del/moveFrom
        if tag in (W_INS, W_MOVE_TO):
            for child in el:
                _process_body_child(child, ctx)
        return

    metadata = _revision_metadata(el)
    ctx.builder.begin_div(classes=(rev_class,), kv=metadata)
    for child in el:
        _process_body_child(child, ctx)
    ctx.builder.end()
    _emit_revision_annotation(ctx, change_type=change_type, metadata=metadata)


def _handle_inline_revision(
    el: etree._Element,
    ctx: ParseContext,
    tag: str,
    out: list[Inline],
) -> None:
    """Handle an inline-level revision element (w:ins, w:del, w:moveFrom, w:moveTo).

    When ``ctx.track_changes`` is True, wrap the inner runs in a ``Span``
    with ``rev-*`` classes and metadata, and emit a TRACKED_CHANGE
    annotation. Otherwise apply the legacy flatten behavior.
    """
    rev_class, change_type = _REV_TAG_MAP[tag]

    if not ctx.track_changes:
        # Legacy behavior: include ins/moveTo runs, skip del/moveFrom
        if tag in (W_INS, W_MOVE_TO):
            for sub in el:
                if sub.tag == W_R:
                    _process_run(sub, ctx, out)
                elif sub.tag == W_HYPERLINK:
                    _process_hyperlink(sub, ctx, out)
        return

    # Collect inner inlines into a local list so we can wrap them
    inner: list[Inline] = []
    for sub in el:
        if sub.tag == W_R:
            _process_run(sub, ctx, inner)
        elif sub.tag == W_HYPERLINK:
            _process_hyperlink(sub, ctx, inner)
    if not inner:
        return

    metadata = _revision_metadata(el)
    out.append(
        Span(
            children=tuple(inner),
            attr=Attr(classes=(rev_class,), kv=metadata),
        )
    )
    _emit_revision_annotation(ctx, change_type=change_type, metadata=metadata)


# --------------------------------------------------------------------------
# Paragraph handling
# --------------------------------------------------------------------------


def _handle_paragraph(el: etree._Element, ctx: ParseContext) -> None:
    """Process a w:p element into heading, list item, or paragraph."""
    ppr = el.find(W_PPR)

    # Get style ID
    style_id = _get_style_id(ppr)

    # 1. Check for heading
    heading_level = ctx.styles.heading_level(style_id)

    # Also check for direct outlineLvl on the paragraph (overrides style)
    if heading_level is None and ppr is not None:
        outline_el = ppr.find(W_OUTLINE_LVL)
        if outline_el is not None:
            val = outline_el.get(W_VAL)
            if val is not None:
                with contextlib.suppress(ValueError):
                    heading_level = min(int(val) + 1, 6)

    if heading_level is not None:
        _flush_open_lists(ctx)
        inlines = _collect_inlines(el, ctx)
        text = _inlines_to_text(inlines)
        if text.strip():
            ctx.builder.heading(heading_level, text.strip())
            ctx.builder.with_provenance(extractor=_EXTRACTOR)
        return

    # 2. Check for code style
    if ctx.styles.is_code_style(style_id):
        _flush_open_lists(ctx)
        text = _collect_plain_text(el)
        if text:
            ctx.builder.code_block(text)
            ctx.builder.with_provenance(extractor=_EXTRACTOR)
        return

    # 3. Check for list item
    num_pr = ppr.find(W_NUMPR) if ppr is not None else None
    if num_pr is not None:
        num_id_el = num_pr.find(W_NUMID)
        if num_id_el is not None:
            num_id = num_id_el.get(W_VAL) or "0"
            # numId="0" means "no list" (explicit removal)
            if num_id != "0":
                _handle_list_paragraph(el, num_pr, ctx)
                return

    # 4. Regular paragraph
    _flush_open_lists(ctx)
    inlines = _collect_inlines(el, ctx)
    if inlines:
        ctx.builder.paragraph(*inlines)
        ctx.builder.with_provenance(extractor=_EXTRACTOR)


def _handle_list_paragraph(
    el: etree._Element,
    num_pr: etree._Element,
    ctx: ParseContext,
) -> None:
    """Process a paragraph that is a list item."""
    num_id_el = num_pr.find(W_NUMID)
    ilvl_el = num_pr.find(W_ILVL)
    num_id = num_id_el.get(W_VAL) if num_id_el is not None else "0"
    ilvl = int(ilvl_el.get(W_VAL) or "0") if ilvl_el is not None else 0
    ordered = ctx.numbering.is_ordered(num_id, str(ilvl))

    # Close lists that are deeper than current level or different list ID at same level
    while ctx.list_stack and (
        ctx.list_stack[-1].ilvl > ilvl
        or (ctx.list_stack[-1].ilvl == ilvl and ctx.list_stack[-1].num_id != num_id)
    ):
        top = ctx.list_stack[-1]
        if top.item_open:
            ctx.builder.end()  # Close list item
            top.item_open = False
        ctx.builder.end()  # Close list
        ctx.list_stack.pop()

    # Close any open item at the same level (sibling item)
    if ctx.list_stack and ctx.list_stack[-1].ilvl == ilvl and ctx.list_stack[-1].item_open:
        ctx.builder.end()  # Close previous list item
        ctx.list_stack[-1].item_open = False

    # Open new list if needed
    if not ctx.list_stack or ctx.list_stack[-1].ilvl < ilvl:
        ctx.builder.begin_list(ordered=ordered)
        ctx.list_stack.append(ListState(num_id=num_id, ilvl=ilvl, ordered=ordered))

    # Emit list item (leave it open for potential nested lists)
    inlines = _collect_inlines(el, ctx)
    ctx.builder.begin_list_item()
    ctx.list_stack[-1].item_open = True
    if inlines:
        ctx.builder.paragraph(*inlines)


def _flush_open_lists(ctx: ParseContext) -> None:
    """Close all open lists (call at end of document or on non-list paragraph)."""
    while ctx.list_stack:
        top = ctx.list_stack[-1]
        if top.item_open:
            ctx.builder.end()  # Close list item
            top.item_open = False
        ctx.builder.end()  # Close list
        ctx.list_stack.pop()


# --------------------------------------------------------------------------
# Run (inline) processing
# --------------------------------------------------------------------------


def _collect_inlines(para_el: etree._Element, ctx: ParseContext) -> list[Inline]:
    """Collect all inline content from a paragraph."""
    inlines: list[Inline] = []

    for child in para_el:
        tag = child.tag

        if tag == W_R:
            _process_run(child, ctx, inlines)
        elif tag == W_HYPERLINK:
            _process_hyperlink(child, ctx, inlines)
        elif tag in _REV_TAG_MAP:
            _handle_inline_revision(child, ctx, tag, inlines)
        elif tag == W_BOOKMARK_START:
            bm_id = child.get(W_ID)
            bm_name = child.get(
                qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "name")
            )
            if bm_id and bm_name:
                ctx.bookmarks[bm_id] = bm_name
        elif tag == W_SDT:
            # Phase 6.2: inline SDT — wrap runs in Span(classes=("sdt",))
            # carrying the sdtPr metadata so the writer can re-emit the
            # control around the same runs. Pre-6.2 this simply unwrapped,
            # losing the control definition on every round-trip.
            sdt_content = child.find(W_SDTCONTENT)
            if sdt_content is None:
                continue
            sdt_attr = _parse_sdt_pr(child)
            if not sdt_attr:
                # Legacy unwrap fallback when there's nothing worth
                # preserving (e.g. anonymous SDT with no tag / alias).
                for sub in sdt_content:
                    if sub.tag == W_R:
                        _process_run(sub, ctx, inlines)
                    elif sub.tag == W_HYPERLINK:
                        _process_hyperlink(sub, ctx, inlines)
                continue
            inner: list[Inline] = []
            for sub in sdt_content:
                if sub.tag == W_R:
                    _process_run(sub, ctx, inner)
                elif sub.tag == W_HYPERLINK:
                    _process_hyperlink(sub, ctx, inner)
            if inner:
                inlines.append(
                    Span(
                        children=tuple(inner),
                        attr=Attr(classes=("sdt",), kv=sdt_attr),
                    )
                )
        # Comment range markers, field chars, etc. silently skipped

    return _merge_adjacent_text(inlines)


def _process_run(run_el: etree._Element, ctx: ParseContext, out: list[Inline]) -> None:
    """Extract inline content from a w:r element."""
    rpr = run_el.find(W_RPR)
    is_bold = _has_toggle(rpr, W_B)
    is_italic = _has_toggle(rpr, W_I)
    is_underline = rpr is not None and rpr.find(W_U) is not None
    is_strike = _has_toggle(rpr, W_STRIKE)
    vert_align = _get_vert_align(rpr)

    for child in run_el:
        tag = child.tag

        if tag in (W_T, W_DEL_TEXT):
            text = child.text or ""
            if not text:
                continue
            node: Inline = Text(value=text)
            node = _apply_formatting(node, is_bold, is_italic, is_underline, is_strike, vert_align)
            out.append(node)

        elif tag == W_TAB:
            out.append(Text(value="\t"))

        elif tag == W_BR:
            br_type = child.get(W_TYPE)
            if br_type == "page":
                pass  # Page break handled at block level
            else:
                out.append(LineBreak())

        elif tag == W_DRAWING:
            img = _extract_image(child, ctx)
            if img is not None:
                out.append(img)

        elif tag == W_FOOTNOTE_REFERENCE:
            fn_id = child.get(W_ID)
            if fn_id and fn_id not in ("0", "-1"):
                from kaos_content.model.inlines import FootnoteRef

                out.append(FootnoteRef(identifier=fn_id))

        elif tag == W_ENDNOTE_REFERENCE:
            en_id = child.get(W_ID)
            if en_id and en_id not in ("0", "-1"):
                from kaos_content.model.inlines import FootnoteRef

                out.append(FootnoteRef(identifier=f"en-{en_id}"))


def _process_hyperlink(hl_el: etree._Element, ctx: ParseContext, out: list[Inline]) -> None:
    """Process a w:hyperlink element."""
    # Resolve URL
    rel_id = hl_el.get(R_ID)
    anchor = hl_el.get(W_ANCHOR)

    url = ""
    if rel_id:
        url = ctx.rels_external.get(rel_id, ctx.rels.get(rel_id, ""))
    elif anchor:
        url = f"#{anchor}"

    # Collect inline content from runs within the hyperlink
    link_inlines: list[Inline] = []
    for child in hl_el:
        if child.tag == W_R:
            _process_run(child, ctx, link_inlines)
        elif child.tag == W_INS:
            for sub in child:
                if sub.tag == W_R:
                    _process_run(sub, ctx, link_inlines)

    if link_inlines and url:
        out.append(Link(url=url, children=tuple(link_inlines)))
    elif link_inlines:
        # No URL — just include the text content
        out.extend(link_inlines)


# --------------------------------------------------------------------------
# Formatting helpers
# --------------------------------------------------------------------------


def _has_toggle(rpr: etree._Element | None, tag: str) -> bool:
    """Check if a run property toggle is set (handles val="0" / val="false")."""
    if rpr is None:
        return False
    el = rpr.find(tag)
    if el is None:
        return False
    val = el.get(W_VAL)
    if val is None:
        return True  # Presence without val means "on"
    return val.lower() not in ("0", "false", "off")


def _get_vert_align(rpr: etree._Element | None) -> str | None:
    """Get vertical alignment (superscript/subscript) from run properties."""
    if rpr is None:
        return None
    el = rpr.find(W_VERTALING)
    if el is None:
        return None
    return el.get(W_VAL)


def _apply_formatting(
    node: Inline,
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    vert_align: str | None,
) -> Inline:
    """Wrap a text node in formatting inlines."""
    if bold:
        node = Strong(children=(node,))
    if italic:
        node = Emphasis(children=(node,))
    if underline:
        node = Underline(children=(node,))
    if strikethrough:
        node = Strikethrough(children=(node,))
    if vert_align == "superscript":
        node = Superscript(children=(node,))
    elif vert_align == "subscript":
        node = Subscript(children=(node,))
    return node


def _get_style_id(ppr: etree._Element | None) -> str | None:
    """Get the paragraph style ID from pPr."""
    if ppr is None:
        return None
    style_el = ppr.find(W_PSTYLE)
    if style_el is None:
        return None
    return style_el.get(W_VAL)


W_OUTLINE_LVL = qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "outlineLvl")


# --------------------------------------------------------------------------
# Table handling
# --------------------------------------------------------------------------


def _handle_table(el: etree._Element, ctx: ParseContext) -> None:
    """Process a w:tbl element into a Table block."""
    # Determine grid columns
    grid_cols: list[float | None] = []
    tbl_grid = el.find(W_TBLGRID)
    if tbl_grid is not None:
        for gc in tbl_grid.iter(W_GRIDCOL):
            w_val = gc.get(qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w"))
            grid_cols.append(float(w_val) if w_val else None)

    rows: list[Row] = []
    for tr_el in el.iter(W_TR):
        cells: list[Cell] = []
        for tc_el in tr_el:
            if tc_el.tag != W_TC:
                continue

            # Parse cell properties
            row_span = 1
            col_span = 1
            tcpr = tc_el.find(W_TCPR)
            if tcpr is not None:
                gs = tcpr.find(W_GRIDSPAN)
                if gs is not None:
                    gs_val = gs.get(W_VAL)
                    if gs_val:
                        with contextlib.suppress(ValueError):
                            col_span = int(gs_val)

                vm = tcpr.find(W_VMERGE)
                if vm is not None:
                    val = vm.get(W_VAL)
                    if val is None or val == "":
                        # Continue of vertical merge — empty cell
                        cells.append(Cell(content=(), row_span=0, col_span=col_span))
                        continue
                    # val="restart" means start of vertical merge — row_span computed later

            # Parse cell content (paragraphs within the cell)
            cell_blocks = _parse_cell_content(tc_el, ctx)
            cells.append(Cell(content=tuple(cell_blocks), row_span=row_span, col_span=col_span))

        if cells:
            rows.append(Row(cells=tuple(cells)))

    if not rows:
        return

    # Build table — first row as header if it looks like one
    # For simplicity, treat all rows as body
    table_section = TableSection(rows=tuple(rows))
    table = Table(bodies=(table_section,))
    ctx.builder.add_block(table)
    ctx.builder.with_provenance(extractor=_EXTRACTOR)


def _parse_cell_content(tc_el: etree._Element, ctx: ParseContext) -> list:
    """Parse paragraphs within a table cell, returning Block nodes."""
    from kaos_content.model.blocks import Block, Paragraph

    blocks: list[Block] = []
    for child in tc_el:
        if child.tag == W_P:
            inlines = _collect_inlines(child, ctx)
            if inlines:
                blocks.append(Paragraph(children=tuple(inlines)))
        elif child.tag == W_TBL:
            # Nested table — build recursively
            # For now, flatten nested tables by extracting their text
            nested_inlines = _collect_nested_table_text(child, ctx)
            if nested_inlines:
                blocks.append(Paragraph(children=tuple(nested_inlines)))
    return blocks


def _collect_nested_table_text(tbl_el: etree._Element, ctx: ParseContext) -> list[Inline]:
    """Flatten a nested table into inline text."""
    inlines: list[Inline] = []
    for p_el in tbl_el.iter(W_P):
        para_inlines = _collect_inlines(p_el, ctx)
        if para_inlines:
            if inlines:
                inlines.append(Text(value=" | "))
            inlines.extend(para_inlines)
    return inlines


# --------------------------------------------------------------------------
# Image extraction
# --------------------------------------------------------------------------


_SDT_CONTROL_TAGS: tuple[str, ...] = (
    # Order matches ECMA-376 §17.5.2 precedence — text first because
    # Word emits it most commonly, richText next, then typed controls.
    "text",
    "richText",
    "picture",
    "date",
    "checkbox",
    "comboBox",
    "dropDownList",
    "docPartObj",
    "docPartList",
    "group",
)


def _parse_sdt_pr(sdt_el: etree._Element) -> dict[str, str]:
    """Extract preserving-metadata from a ``<w:sdt>`` element.

    Returns a dict of ``sdt.*`` kv pairs suitable for an :class:`Attr`.
    Keys follow a dotted namespace so REDACTION / rev:* / sdt.* all
    coexist on ``Attr.kv`` without collision.

    What we capture (enough for lossless Word round-trip):
    - ``sdt.tag`` — the ``<w:tag w:val="..."/>`` identifier callers
      use to locate the control.
    - ``sdt.alias`` — human-readable ``<w:alias w:val="..."/>`` name.
    - ``sdt.id`` — numeric ``<w:id w:val="..."/>`` (regenerated by Word
      when absent, but preserved if present so cross-refs survive).
    - ``sdt.lock`` — ``<w:lock w:val="..."/>`` one of sdtLocked /
      contentLocked / sdtContentLocked / unlocked.
    - ``sdt.control_type`` — the control-kind child tag (``text``,
      ``richText``, ``picture``, ``date``, ``checkbox``, ``comboBox``,
      ``dropDownList``, ``docPartObj``, ``docPartList``, ``group``).
    - ``sdt.showing_placeholder`` — ``"1"`` iff ``<w:showingPlcHdr/>``
      is present. Without it Word may clear user-typed content
      thinking it's a stale placeholder.

    What we deliberately DON'T capture on round-trip:
    - ``<w:dataBinding w:xpath ... w:storeItemID ...>``: references a
      ``/customXml/item*.xml`` part that lives in the OPC package but
      isn't (yet) threaded through the AST. Emitting a dangling
      binding reference produces a red-boxed error in Word, so we
      strip it — the control degrades gracefully to a static value.
    - ``<w:placeholder><w:docPart w:val="..."/>``: references
      ``/word/glossary/document.xml`` for the same reason.
    - Per-control-kind properties (dateFormat, list items, etc.):
      deferred to a future "full fidelity" iteration; for now the
      control shows but its validation rules are lost.

    Returns ``{}`` when the SDT has no ``<w:sdtPr>`` or no preservable
    metadata, signalling the caller to fall back to legacy unwrap.
    """
    pr = sdt_el.find(qn(W, "sdtPr"))
    if pr is None:
        return {}

    def _val(tag_name: str) -> str | None:
        el = pr.find(qn(W, tag_name))
        if el is None:
            return None
        val = el.get(qn(W, "val"))
        return val if val else None

    out: dict[str, str] = {}
    tag_val = _val("tag")
    if tag_val is not None:
        out["sdt.tag"] = tag_val
    alias = _val("alias")
    if alias is not None:
        out["sdt.alias"] = alias
    sdt_id = _val("id")
    if sdt_id is not None:
        out["sdt.id"] = sdt_id
    lock = _val("lock")
    if lock is not None:
        out["sdt.lock"] = lock

    # Control-type detection: the first matching direct child wins.
    # Mutually exclusive per the schema.
    for ctrl in _SDT_CONTROL_TAGS:
        if pr.find(qn(W, ctrl)) is not None:
            out["sdt.control_type"] = ctrl
            break

    if pr.find(qn(W, "showingPlcHdr")) is not None:
        out["sdt.showing_placeholder"] = "1"

    return out


def _inline_data_uri(data: bytes, fmt: str, index: int) -> str:
    """Default :func:`parse_docx` ``image_src_builder``.

    Emits a self-contained ``data:image/<fmt>;base64,<payload>`` URI so
    the resulting ContentDocument round-trips through the DOCX writer
    without a side-channel byte store. Mirror of kaos-pdf's
    ``_image_data_uri``. ``index`` is accepted for signature parity with
    caller-supplied builders but is unused here — base64 URIs don't
    need an index to be unique.
    """
    del index  # unused — default builder is stateless
    import base64

    encoded = base64.b64encode(data).decode("ascii")
    return f"data:image/{fmt};base64,{encoded}"


def _media_ext_to_fmt(target: str) -> str | None:
    """Normalize a media part path's extension to the MIME subtype the
    writer + its content-type table understand. Returns ``None`` for
    formats we don't support (Word's own renderer rejects them too).
    """
    ext = target.rsplit(".", 1)[-1].lower() if "." in target else ""
    # Mirror of _IMAGE_CONTENT_TYPES keys in the writer.
    mapping = {
        "png": "png",
        "jpg": "jpeg",
        "jpeg": "jpeg",
        "gif": "gif",
        "bmp": "bmp",
        "tiff": "tiff",
    }
    return mapping.get(ext)


def _extract_image(drawing_el: etree._Element, ctx: ParseContext) -> Image | None:
    """Extract an image from a ``<w:drawing>`` element.

    Phase 6.1: reads the embedded bytes from the OPC package and hands
    them to ``ctx.image_src_builder`` so ``Image.src`` is a URI the
    writer can re-embed. Before 6.1 we returned a bare ``docx://...``
    URI and the writer couldn't round-trip it — images silently dropped
    to alt text on write.
    """
    # Find the blip (image reference)
    blip = drawing_el.find(f".//{qn(A, 'blip')}")
    if blip is None:
        return None

    embed_id = blip.get(R_EMBED)
    if embed_id is None:
        return None

    # Resolve relationship ID to media path
    target = ctx.rels.get(embed_id)
    if target is None:
        return None

    # Get alt text from docPr
    doc_pr = drawing_el.find(f".//{qn(WP, 'docPr')}")
    alt = doc_pr.get("descr") if doc_pr is not None else None
    title = doc_pr.get("name") if doc_pr is not None else None

    # Extract dimensions from wp:extent (EMUs in OOXML → points for AST)
    width_pt: float | None = None
    height_pt: float | None = None
    extent = drawing_el.find(f".//{qn(WP, 'extent')}")
    if extent is not None:
        cx = extent.get("cx")
        cy = extent.get("cy")
        try:
            if cx is not None:
                width_pt = emu_to_pt(int(cx))
            if cy is not None:
                height_pt = emu_to_pt(int(cy))
        except (TypeError, ValueError):
            pass

    # Resolve ``Image.src`` by reading media bytes and handing them to
    # the configured builder. If the package lookup fails (missing
    # part, unsupported extension, old ``parse_docx`` call with
    # ``ctx.pkg=None``) fall back to the pre-6.1 bare ``docx://`` URI
    # so at least the AST shape survives.
    fmt = _media_ext_to_fmt(target)
    src: str
    if ctx.pkg is not None and ctx.image_src_builder is not None and fmt is not None:
        try:
            data = ctx.pkg.read_part(target)
        except (KeyError, FileNotFoundError, OSError):
            data = None
        if data:
            ctx.image_counter += 1
            src = ctx.image_src_builder(data, fmt, ctx.image_counter)
        else:
            media_name = target.rsplit("/", 1)[-1] if "/" in target else target
            src = f"docx://{media_name}"
    else:
        media_name = target.rsplit("/", 1)[-1] if "/" in target else target
        src = f"docx://{media_name}"

    return Image(
        src=src,
        alt=alt or None,
        title=title or None,
        width=width_pt,
        height=height_pt,
    )


# --------------------------------------------------------------------------
# Headers, footers, and page setup (Phase 4)
# --------------------------------------------------------------------------


def _parse_header_footer_part(
    root: etree._Element,
    parent_ctx: ParseContext,
) -> tuple:
    """Walk a ``<w:hdr>`` or ``<w:ftr>`` root and return a tuple of Blocks.

    Uses a fresh :class:`DocumentBuilder` so the header/footer content
    doesn't pollute the main body. Shares styles / numbering / rels with
    the parent context — for an MVP this means header images resolve via
    the document's rels_map, which covers the common case (logos in
    headers) but not per-part rels files.
    """
    sub_builder = DocumentBuilder()
    sub_ctx = ParseContext(
        builder=sub_builder,
        styles=parent_ctx.styles,
        numbering=parent_ctx.numbering,
        rels=parent_ctx.rels,
        rels_external=parent_ctx.rels_external,
        source_uri=parent_ctx.source_uri,
    )
    for child in root:
        _process_body_child(child, sub_ctx)
    _flush_open_lists(sub_ctx)
    return sub_builder.build().body


def _parse_headers_and_footers(
    doc_rels: object,
    pkg: OPCPackage,
    doc_dir: str,
    body_el: etree._Element,
    parent_ctx: ParseContext,
) -> tuple[dict[str, tuple], dict[str, tuple]]:
    """Parse all referenced header and footer parts.

    Walks every ``<w:sectPr>`` in the body, collects
    ``<w:headerReference r:id=... w:type=.../>`` and
    ``<w:footerReference ...>`` entries, and parses each referenced part
    once. The ``w:type`` attribute (``default`` / ``first`` / ``even``)
    becomes the dict key.
    """
    from kaos_office.opc.relationships import RelationshipManager

    assert isinstance(doc_rels, RelationshipManager)

    headers: dict[str, tuple] = {}
    footers: dict[str, tuple] = {}

    # Track which rIds we've already processed so the same header.xml
    # referenced by multiple sections doesn't parse twice.
    seen: set[str] = set()

    for ref_tag, bucket, _rel_type in (
        (W_HEADER_REFERENCE, headers, RT_HEADER),
        (W_FOOTER_REFERENCE, footers, RT_FOOTER),
    ):
        for ref in body_el.iter(ref_tag):
            rid = ref.get(R_ID_ATTR)
            if rid is None or rid in seen:
                continue
            seen.add(rid)
            ref_type = ref.get(W_TYPE) or "default"
            target = parent_ctx.rels.get(rid)
            if target is None:
                continue
            # target is already resolved relative to doc_dir in rels_map
            if not pkg.has_part(target):
                continue
            try:
                root = pkg.read_xml(target)
            except Exception:  # malformed XML part
                logger.debug("Failed to read header/footer part %s", target, exc_info=True)
                continue
            blocks = _parse_header_footer_part(root, parent_ctx)
            if blocks:
                bucket[ref_type] = blocks
    return headers, footers


def _twips_attr(attr: etree._Element | None, name: str) -> float | None:
    if attr is None:
        return None
    val = attr.get(qn(W, name))
    if val is None:
        return None
    try:
        return twips_to_pt(int(val))
    except ValueError:
        return None


def _parse_page_setup_from_sect(sect: etree._Element) -> PageSetup | None:
    """Extract page size / margins from a single ``<w:sectPr>`` element.

    Returns ``None`` if every geometry field is absent, matching the
    "no meaningful geometry to preserve" signal the caller uses.
    """
    pg_sz = sect.find(W_PGSZ)
    pg_mar = sect.find(W_PGMAR)

    setup = PageSetup(
        page_width_pt=_twips_attr(pg_sz, "w"),
        page_height_pt=_twips_attr(pg_sz, "h"),
        margin_top_pt=_twips_attr(pg_mar, "top"),
        margin_bottom_pt=_twips_attr(pg_mar, "bottom"),
        margin_left_pt=_twips_attr(pg_mar, "left"),
        margin_right_pt=_twips_attr(pg_mar, "right"),
        header_distance_pt=_twips_attr(pg_mar, "header"),
        footer_distance_pt=_twips_attr(pg_mar, "footer"),
    )
    if all(
        v is None
        for v in (
            setup.page_width_pt,
            setup.page_height_pt,
            setup.margin_top_pt,
            setup.margin_bottom_pt,
            setup.margin_left_pt,
            setup.margin_right_pt,
            setup.header_distance_pt,
            setup.footer_distance_pt,
        )
    ):
        return None
    return setup


def _parse_break_type(sect: etree._Element) -> SectionBreakType:
    """Extract ``<w:type w:val="..."/>`` — defaults to ``"nextPage"``.

    OOXML omits ``w:type`` when the default applies; we emit it verbatim
    for round-trip fidelity. Unknown values collapse to the default
    rather than raising — new OOXML versions may extend the vocabulary
    and we shouldn't fail-parse on that.
    """
    type_el = sect.find(qn(W, "type"))
    if type_el is None:
        return "nextPage"
    val = type_el.get(qn(W, "val"))
    if val in ("continuous", "nextPage", "nextColumn", "evenPage", "oddPage"):
        return val
    return "nextPage"


def _parse_page_setup(body_el: etree._Element) -> PageSetup | None:
    """Extract the document's representative page setup.

    OOXML may carry multiple ``<w:sectPr>`` — one per section. For the
    backward-compatible single-value ``DocumentMetadata.page_setup`` we
    surface the **final** section's geometry since that's the value
    Word treats as the document default. Per-section page setup lives
    on :attr:`ContentDocument.sections` (populated by
    :func:`_build_sections`).
    """
    sect_prs = list(body_el.iter(W_SECTPR))
    if not sect_prs:
        return None
    return _parse_page_setup_from_sect(sect_prs[-1])


def _build_sections(pending: list[tuple[int, etree._Element]]) -> tuple[Section, ...]:
    """Convert the walk-time boundary list into a tuple of ``Section``s.

    Each ``(end_block_index, sectPr_element)`` tuple becomes one
    ``Section`` in the same order, preserving OOXML document order.
    Returns an empty tuple when the document has no sectPr at all, so
    callers get the "implicit single section" shape for free.
    """
    sections: list[Section] = []
    for end_idx, sect_el in pending:
        sections.append(
            Section(
                end_block_index=end_idx,
                page_setup=_parse_page_setup_from_sect(sect_el),
                break_type=_parse_break_type(sect_el),
            )
        )
    return tuple(sections)


# --------------------------------------------------------------------------
# Footnotes and endnotes
# --------------------------------------------------------------------------


def _parse_footnotes(
    doc_rels: object,
    pkg: OPCPackage,
    doc_dir: str,
    rel_type: str,
    container_tag: str,
    item_tag: str,
) -> dict[str, list[etree._Element]]:
    """Parse footnotes or endnotes from their XML part."""
    from kaos_office.opc.relationships import RelationshipManager

    assert isinstance(doc_rels, RelationshipManager)
    target = doc_rels.first_target(rel_type)
    if target is None:
        return {}

    part_path = f"{doc_dir}/{target}" if doc_dir else target
    if not pkg.has_part(part_path):
        return {}

    root = pkg.read_xml(part_path)
    notes: dict[str, list[etree._Element]] = {}
    for note in root.iter(item_tag):
        note_id = note.get(W_ID)
        if note_id is None or note_id in ("0", "-1"):
            continue  # Skip separator and continuation notices
        # Collect the paragraph elements within the note
        paras = [p for p in note if p.tag == W_P]
        if paras:
            notes[note_id] = paras

    return notes


def _add_footnotes(ctx: ParseContext) -> None:
    """Add parsed footnote and endnote content to the builder."""
    for fn_id, para_els in ctx.footnotes.items():
        inlines = []
        for p_el in para_els:
            p_inlines = _collect_inlines(p_el, ctx)
            if p_inlines:
                if inlines:
                    inlines.append(Text(value=" "))
                inlines.extend(p_inlines)
        if inlines:
            from kaos_content.model.blocks import Paragraph

            ctx.builder.add_footnote(fn_id, Paragraph(children=tuple(inlines)))

    for en_id, para_els in ctx.endnotes.items():
        inlines = []
        for p_el in para_els:
            p_inlines = _collect_inlines(p_el, ctx)
            if p_inlines:
                if inlines:
                    inlines.append(Text(value=" "))
                inlines.extend(p_inlines)
        if inlines:
            from kaos_content.model.blocks import Paragraph

            ctx.builder.add_footnote(f"en-{en_id}", Paragraph(children=tuple(inlines)))


# --------------------------------------------------------------------------
# Comments
# --------------------------------------------------------------------------


def _parse_comments(
    doc_rels: object,
    pkg: OPCPackage,
    doc_dir: str,
) -> dict[str, dict[str, str]]:
    """Parse comments from comments.xml."""
    from kaos_office.opc.relationships import RelationshipManager

    assert isinstance(doc_rels, RelationshipManager)
    target = doc_rels.first_target(RT_COMMENTS)
    if target is None:
        return {}

    part_path = f"{doc_dir}/{target}" if doc_dir else target
    if not pkg.has_part(part_path):
        return {}

    root = pkg.read_xml(part_path)
    comments: dict[str, dict[str, str]] = {}
    for comment_el in root.iter(W_COMMENT):
        comment_id = comment_el.get(W_ID)
        if comment_id is None:
            continue
        author = (
            comment_el.get(
                qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "author")
            )
            or ""
        )
        date = (
            comment_el.get(
                qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "date")
            )
            or ""
        )
        # Collect text from all paragraphs in the comment
        text_parts = []
        for p in comment_el.iter(W_P):
            for t in p.iter(W_T):
                if t.text:
                    text_parts.append(t.text)
        text = " ".join(text_parts)
        comments[comment_id] = {"author": author, "date": date, "text": text}

    return comments


def _add_comment_annotations(ctx: ParseContext) -> None:
    """Add comment annotations to the builder."""
    for comment_id, comment_data in ctx.comments.items():
        if comment_data.get("text"):
            ctx.builder.annotate(
                AnnotationType.COMMENT,
                targets=[],
                body={
                    "comment_id": comment_id,
                    "author": comment_data.get("author", ""),
                    "date": comment_data.get("date", ""),
                    "text": comment_data["text"],
                },
            )


# --------------------------------------------------------------------------
# Text utilities
# --------------------------------------------------------------------------


def _collect_plain_text(para_el: etree._Element) -> str:
    """Collect plain text from a paragraph (ignoring formatting)."""
    parts: list[str] = []
    for t in para_el.iter(W_T):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def _inlines_to_text(inlines: list[Inline]) -> str:
    """Extract plain text from a list of inline nodes."""
    parts: list[str] = []
    for inline in inlines:
        if isinstance(inline, Text):
            parts.append(inline.value)
        elif hasattr(inline, "children"):
            parts.append(_inlines_to_text(list(inline.children)))  # type: ignore[attr-defined]  # ty: ignore[invalid-argument-type]
        elif hasattr(inline, "value"):
            parts.append(inline.value)  # type: ignore[attr-defined]  # ty: ignore[invalid-argument-type]
    return "".join(parts)


def _merge_adjacent_text(inlines: list[Inline]) -> list[Inline]:
    """Merge adjacent plain Text nodes."""
    if not inlines:
        return inlines

    merged: list[Inline] = []
    for inline in inlines:
        if isinstance(inline, Text) and merged and isinstance(merged[-1], Text):
            merged[-1] = Text(value=merged[-1].value + inline.value)
        else:
            merged.append(inline)
    return merged


# --------------------------------------------------------------------------
# Internal helpers
# --------------------------------------------------------------------------


def _iter_rel_ids(rels_mgr: object) -> list[str]:
    """Get all relationship IDs from a RelationshipManager."""
    from kaos_office.opc.relationships import RelationshipManager

    assert isinstance(rels_mgr, RelationshipManager)
    return list(rels_mgr._by_id.keys())


def _resolve_part(
    doc_rels: object,
    pkg: OPCPackage,
    doc_dir: str,
    part_name: str,
) -> bytes | None:
    """Try to read a supporting part (styles, numbering) by relationship or direct path."""
    from kaos_office.opc.relationships import RelationshipManager

    assert isinstance(doc_rels, RelationshipManager)

    # Common relationship type URIs
    rel_types = {
        "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "numbering": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    }

    rel_type = rel_types.get(part_name)
    if rel_type:
        target = doc_rels.first_target(rel_type)
        if target:
            full_path = f"{doc_dir}/{target}" if doc_dir else target
            if pkg.has_part(full_path):
                return pkg.read_part(full_path)

    # Fall back to well-known paths
    well_known = {
        "styles": "word/styles.xml",
        "numbering": "word/numbering.xml",
    }
    wk_path = well_known.get(part_name)
    if wk_path and pkg.has_part(wk_path):
        return pkg.read_part(wk_path)

    return None
