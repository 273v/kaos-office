"""DOCX writer — ContentDocument to WordprocessingML.

Serializes a kaos-content ContentDocument AST to a standards-compliant
DOCX file using pure lxml and the OPC write layer. No python-docx
dependency — consistent with kaos-office's native lxml philosophy.

Usage::

    from kaos_office.docx.writer import write_docx, write_docx_bytes

    write_docx(content_doc, "output.docx")
    docx_bytes = write_docx_bytes(content_doc)
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from kaos_core.logging import get_logger
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import (
    CT_FOOTER,
    CT_HEADER,
    CT_SETTINGS,
    R_ID_ATTR,
    RT_FOOTER,
    RT_HEADER,
    RT_SETTINGS,
    W_BODY,
    W_EVEN_AND_ODD_HEADERS,
    W_FOOTER_REFERENCE,
    W_FTR,
    W_HDR,
    W_HEADER_REFERENCE,
    W_P,
    W_PGMAR,
    W_PGSZ,
    W_PPR,
    W_PSTYLE,
    W_R,
    W_RPR,
    W_SETTINGS,
    W_T,
    W_TBL,
    W_TC,
    W_TITLEPG,
    W_TR,
    W_TYPE,
    R,
    W,
    pt_to_twips,
    qn,
)
from kaos_office.opc.package import OPCPackageWriter

logger = get_logger(__name__)

# Standard XML namespace for xml:space="preserve"
_XML_NS = "http://www.w3.org/XML/1998/namespace"
_XML_SPACE = f"{{{_XML_NS}}}space"

# XSI namespace for xsi:type on dcterms elements
_XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

# Content types
_CT_DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_CT_NUMBERING = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_CT_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml"
_CT_FOOTNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_CT_ENDNOTES = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
_CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

# Relationship types
_RT_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
_RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_RT_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
_RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
_RT_FOOTNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_RT_ENDNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
_RT_COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_RT_CORE_PROPS = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)

# Namespace maps
_W_NSMAP = {"w": W, "r": R}
_CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
_DC_NS = "http://purl.org/dc/elements/1.1/"
_DCTERMS_NS = "http://purl.org/dc/terms/"


def write_docx(doc: Any, path: str | Path) -> Path:
    """Write a ContentDocument to a DOCX file.

    Args:
        doc: A ``ContentDocument`` from kaos-content.
        path: Output file path.

    Returns:
        The output path.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    data = write_docx_bytes(doc)
    path.write_bytes(data)

    logger.info(
        "docx.writer: wrote %s, blocks=%d, size=%d, path=%s",
        doc.metadata.title or "untitled",
        len(doc.body),
        len(data),
        path,
    )
    return path


def write_docx_bytes(doc: Any) -> bytes:
    """Write a ContentDocument to DOCX bytes (in-memory)."""
    writer = OPCPackageWriter()

    # Content types
    writer.content_types.add_default(
        "rels", "application/vnd.openxmlformats-package.relationships+xml"
    )
    writer.content_types.add_default("xml", "application/xml")
    writer.content_types.add_override("/word/document.xml", _CT_DOCUMENT)
    writer.content_types.add_override("/word/styles.xml", _CT_STYLES)
    writer.content_types.add_override("/word/numbering.xml", _CT_NUMBERING)
    writer.content_types.add_override("/docProps/core.xml", _CT_CORE_PROPS)

    # Root rels
    writer.root_rels.add(_RT_OFFICE, "word/document.xml")
    writer.root_rels.add(_RT_CORE_PROPS, "docProps/core.xml")

    # Document rels
    doc_rels = writer.get_rels("word/document.xml")
    doc_rels.add(_RT_STYLES, "styles.xml")
    doc_rels.add(_RT_NUMBERING, "numbering.xml")

    # Build document.xml
    ctx = _WriteContext(doc_rels=doc_rels)
    _prepare_notes(doc, ctx)

    # Headers & footers (Phase 4) — parts are written before document.xml so
    # relationship IDs resolve at body-end serialization. The context's
    # header/footer_refs are consulted when building <w:sectPr>.
    _write_header_footer_parts(doc, ctx, writer)

    document_xml = _build_document(doc, ctx)
    writer.add_xml_part("word/document.xml", document_xml)

    # Build styles.xml
    writer.add_xml_part("word/styles.xml", _build_styles())

    # Build numbering.xml
    if ctx.has_lists:
        writer.add_xml_part("word/numbering.xml", _build_numbering())
    else:
        writer.add_xml_part("word/numbering.xml", _build_empty_numbering())

    # Build footnotes / endnotes / comments if any were collected
    if ctx.has_footnotes:
        writer.content_types.add_override("/word/footnotes.xml", _CT_FOOTNOTES)
        doc_rels.add(_RT_FOOTNOTES_REL, "footnotes.xml")
        writer.add_xml_part("word/footnotes.xml", _build_footnotes(ctx))

    if ctx.has_endnotes:
        writer.content_types.add_override("/word/endnotes.xml", _CT_ENDNOTES)
        doc_rels.add(_RT_ENDNOTES_REL, "endnotes.xml")
        writer.add_xml_part("word/endnotes.xml", _build_endnotes(ctx))

    if ctx.comments:
        writer.content_types.add_override("/word/comments.xml", _CT_COMMENTS)
        doc_rels.add(_RT_COMMENTS_REL, "comments.xml")
        writer.add_xml_part("word/comments.xml", _build_comments(ctx))

    # <w:evenAndOddHeaders/> in word/settings.xml is the document-wide gate
    # for w:type="even" header/footer references. Without it Word ignores
    # every even reference, regardless of which sectPr it lives in.
    has_even = any(kind == "even" for kind, _ in ctx.header_refs) or any(
        kind == "even" for kind, _ in ctx.footer_refs
    )
    if has_even:
        writer.content_types.add_override("/word/settings.xml", CT_SETTINGS)
        doc_rels.add(RT_SETTINGS, "settings.xml")
        writer.add_xml_part("word/settings.xml", _build_settings(even_and_odd=True))

    # Build core properties
    writer.add_xml_part("docProps/core.xml", _build_core_properties(doc))

    return writer.save_bytes()


def _build_settings(*, even_and_odd: bool) -> etree._Element:
    """Build a minimal word/settings.xml.

    Only the features the writer actually needs to toggle land here; the
    rest of the settings schema is left to Word's defaults. Grows as we
    add gated features (mirrorMargins, defaultTabStop, ...).
    """
    root = etree.Element(W_SETTINGS, nsmap=_W_NSMAP)
    if even_and_odd:
        etree.SubElement(root, W_EVEN_AND_ODD_HEADERS)
    return root


class _WriteContext:
    """Accumulates state during serialization of the document body.

    ``doc_rels`` is required when any Link / Footnote / Comment is emitted,
    because those need entries in ``word/_rels/document.xml.rels``.
    """

    __slots__ = (
        "comments",
        "doc_rels",
        "endnotes",
        "footer_refs",
        "footnotes",
        "has_endnotes",
        "has_footnotes",
        "has_lists",
        "header_refs",
        "hyperlink_urls",
        "list_counter",
    )

    def __init__(self, doc_rels: Any = None) -> None:
        self.has_lists = False
        self.has_footnotes = False
        self.has_endnotes = False
        self.list_counter = 0
        # Phase A: hyperlink URL -> rel_id (reuses the same rel for duplicate URLs)
        self.hyperlink_urls: dict[str, str] = {}
        self.doc_rels = doc_rels
        # Phase B: footnotes / endnotes indexed by identifier (ID -> list of Paragraph blocks)
        self.footnotes: dict[str, tuple[Any, ...]] = {}
        self.endnotes: dict[str, tuple[Any, ...]] = {}
        # Phase C: comments — list of (comment_id, metadata_dict)
        self.comments: list[tuple[int, dict[str, Any]]] = []
        # Phase 4: list of (kind, rel_id) per ref type for sectPr emission.
        self.header_refs: list[tuple[str, str]] = []
        self.footer_refs: list[tuple[str, str]] = []


def _write_header_footer_parts(doc: Any, ctx: _WriteContext, writer: OPCPackageWriter) -> None:
    """Write word/header*.xml and word/footer*.xml parts for each entry in
    ``doc.headers`` / ``doc.footers``. Records ``(kind, rel_id)`` in the
    write context so ``_build_document`` can emit the corresponding
    ``<w:headerReference>`` / ``<w:footerReference>`` in ``<w:sectPr>``.
    """
    headers = getattr(doc, "headers", {}) or {}
    footers = getattr(doc, "footers", {}) or {}
    if not headers and not footers:
        return

    for idx, (kind, blocks) in enumerate(headers.items(), start=1):
        part_name = f"header{idx}.xml"
        full_path = f"word/{part_name}"
        writer.content_types.add_override(f"/{full_path}", CT_HEADER)
        rel_id = ctx.doc_rels.add(RT_HEADER, part_name).id
        ctx.header_refs.append((kind, rel_id))
        writer.add_xml_part(full_path, _build_header_footer_part("hdr", blocks, ctx))

    for idx, (kind, blocks) in enumerate(footers.items(), start=1):
        part_name = f"footer{idx}.xml"
        full_path = f"word/{part_name}"
        writer.content_types.add_override(f"/{full_path}", CT_FOOTER)
        rel_id = ctx.doc_rels.add(RT_FOOTER, part_name).id
        ctx.footer_refs.append((kind, rel_id))
        writer.add_xml_part(full_path, _build_header_footer_part("ftr", blocks, ctx))


def _build_header_footer_part(kind: str, blocks: tuple, ctx: _WriteContext) -> etree._Element:
    """Build a ``<w:hdr>`` or ``<w:ftr>`` root containing the supplied blocks."""
    tag = W_HDR if kind == "hdr" else W_FTR
    root = etree.Element(tag, nsmap=_W_NSMAP)
    if not blocks:
        # w:hdr / w:ftr must contain at least one paragraph to open cleanly.
        etree.SubElement(root, W_P)
        return root
    for block in blocks:
        _serialize_block(root, block, ctx)
    return root


def _build_document(doc: Any, ctx: _WriteContext) -> etree._Element:
    """Build word/document.xml from ContentDocument."""
    root = etree.Element(qn(W, "document"), nsmap=_W_NSMAP)
    body = etree.SubElement(root, W_BODY)

    for block in doc.body:
        _serialize_block(body, block, ctx)

    sect_pr = etree.SubElement(body, qn(W, "sectPr"))

    # Header / footer references must precede pgSz / pgMar in sectPr.
    for kind, rid in ctx.header_refs:
        etree.SubElement(sect_pr, W_HEADER_REFERENCE, **{W_TYPE: kind, R_ID_ATTR: rid})
    for kind, rid in ctx.footer_refs:
        etree.SubElement(sect_pr, W_FOOTER_REFERENCE, **{W_TYPE: kind, R_ID_ATTR: rid})

    # <w:titlePg/> gates the "first" header/footer — without it Word
    # silently ignores any headerReference/footerReference with w:type="first".
    # The element must come after references and before pgSz per ECMA-376
    # §17.6.17 child-order rules.
    has_first = any(kind == "first" for kind, _ in ctx.header_refs) or any(
        kind == "first" for kind, _ in ctx.footer_refs
    )
    if has_first:
        etree.SubElement(sect_pr, W_TITLEPG)

    # Page setup: prefer doc.metadata.page_setup values; fall back to
    # US Letter with 1 inch margins so the output still opens in Word
    # when the source didn't carry any geometry.
    ps = getattr(getattr(doc, "metadata", None), "page_setup", None)

    def _twips(value: float | None, default_twips: int) -> str:
        if value is None:
            return str(default_twips)
        return str(pt_to_twips(value))

    pg_width = _twips(getattr(ps, "page_width_pt", None) if ps else None, 12240)
    pg_height = _twips(getattr(ps, "page_height_pt", None) if ps else None, 15840)
    margin_top = _twips(getattr(ps, "margin_top_pt", None) if ps else None, 1440)
    margin_right = _twips(getattr(ps, "margin_right_pt", None) if ps else None, 1440)
    margin_bottom = _twips(getattr(ps, "margin_bottom_pt", None) if ps else None, 1440)
    margin_left = _twips(getattr(ps, "margin_left_pt", None) if ps else None, 1440)
    margin_header = _twips(getattr(ps, "header_distance_pt", None) if ps else None, 720)
    margin_footer = _twips(getattr(ps, "footer_distance_pt", None) if ps else None, 720)

    etree.SubElement(sect_pr, W_PGSZ, **{qn(W, "w"): pg_width, qn(W, "h"): pg_height})
    etree.SubElement(
        sect_pr,
        W_PGMAR,
        **{
            qn(W, "top"): margin_top,
            qn(W, "right"): margin_right,
            qn(W, "bottom"): margin_bottom,
            qn(W, "left"): margin_left,
            qn(W, "header"): margin_header,
            qn(W, "footer"): margin_footer,
        },
    )

    return root


def _serialize_block(parent: etree._Element, block: Any, ctx: _WriteContext) -> None:
    """Serialize a single block to OOXML elements."""
    from kaos_content.model.blocks import (
        BlockQuote,
        BulletList,
        CodeBlock,
        Div,
        Heading,
        OrderedList,
        PageBreak,
        Paragraph,
        Table,
        ThematicBreak,
    )

    if isinstance(block, Div) and _is_revision(block):
        _serialize_revision_div(parent, block, ctx)
        return
    if isinstance(block, Paragraph):
        _serialize_paragraph(parent, block, ctx)
    elif isinstance(block, Heading):
        _serialize_heading(parent, block, ctx)
    elif isinstance(block, (BulletList, OrderedList)):
        _serialize_list(parent, block, ctx)
    elif isinstance(block, Table):
        _serialize_table(parent, block, ctx)
    elif isinstance(block, CodeBlock):
        _serialize_code_block(parent, block)
    elif isinstance(block, BlockQuote):
        for child in block.children:
            _serialize_block(parent, child, ctx)
    elif isinstance(block, ThematicBreak):
        p = etree.SubElement(parent, W_P)
        ppr = etree.SubElement(p, W_PPR)
        etree.SubElement(ppr, qn(W, "pBdr"))
    elif isinstance(block, PageBreak):
        p = etree.SubElement(parent, W_P)
        r = etree.SubElement(p, W_R)
        etree.SubElement(r, qn(W, "br"), **{qn(W, "type"): "page"})
    else:
        # Fallback: serialize as paragraph with text content
        from kaos_content.traversal.visitor import extract_text

        text = extract_text(block)
        if text.strip():
            p = etree.SubElement(parent, W_P)
            r = etree.SubElement(p, W_R)
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            t.text = text


def _serialize_paragraph(parent: etree._Element, para: Any, ctx: _WriteContext) -> None:
    """Serialize a Paragraph block."""
    p = etree.SubElement(parent, W_P)
    for inline in para.children:
        _serialize_inline(p, inline, ctx)


def _serialize_heading(parent: etree._Element, heading: Any, ctx: _WriteContext) -> None:
    """Serialize a Heading block with style + outline level."""
    p = etree.SubElement(parent, W_P)
    ppr = etree.SubElement(p, W_PPR)
    depth = heading.depth if hasattr(heading, "depth") else 1
    style_name = f"Heading{min(depth, 6)}"
    etree.SubElement(ppr, W_PSTYLE, **{qn(W, "val"): style_name})
    etree.SubElement(ppr, qn(W, "outlineLvl"), **{qn(W, "val"): str(depth - 1)})

    for inline in heading.children:
        _serialize_inline(p, inline, ctx)


def _serialize_list(parent: etree._Element, lst: Any, ctx: _WriteContext, level: int = 0) -> None:
    """Serialize an ordered or unordered list."""
    from kaos_content.model.blocks import BulletList, ListItem, OrderedList

    ctx.has_lists = True
    num_id = 1 if isinstance(lst, BulletList) else 2

    for item in lst.children:
        if isinstance(item, ListItem):
            for child in item.children:
                if isinstance(child, (BulletList, OrderedList)):
                    _serialize_list(parent, child, ctx, level=level + 1)
                else:
                    p = etree.SubElement(parent, W_P)
                    ppr = etree.SubElement(p, W_PPR)
                    num_pr = etree.SubElement(ppr, qn(W, "numPr"))
                    etree.SubElement(num_pr, qn(W, "ilvl"), **{qn(W, "val"): str(level)})
                    etree.SubElement(num_pr, qn(W, "numId"), **{qn(W, "val"): str(num_id)})
                    if hasattr(child, "children"):
                        for inline in child.children:
                            _serialize_inline(p, inline, ctx)
                    else:
                        from kaos_content.traversal.visitor import extract_text

                        text = extract_text(child)
                        if text:
                            r = etree.SubElement(p, W_R)
                            t = etree.SubElement(r, W_T)
                            t.text = text


def _serialize_table(parent: etree._Element, table: Any, ctx: _WriteContext) -> None:
    """Serialize a Table block."""
    tbl = etree.SubElement(parent, W_TBL)

    # Table properties
    tbl_pr = etree.SubElement(tbl, qn(W, "tblPr"))
    etree.SubElement(tbl_pr, qn(W, "tblStyle"), **{qn(W, "val"): "TableGrid"})
    etree.SubElement(tbl_pr, qn(W, "tblW"), **{qn(W, "w"): "0", qn(W, "type"): "auto"})

    # Table grid
    if hasattr(table, "head") and table.head and table.head.rows:
        n_cols = len(table.head.rows[0].cells) if table.head.rows else 0
    elif hasattr(table, "bodies") and table.bodies:
        n_cols = len(table.bodies[0].rows[0].cells) if table.bodies[0].rows else 0
    else:
        n_cols = 0

    if n_cols > 0:
        grid = etree.SubElement(tbl, qn(W, "tblGrid"))
        for _ in range(n_cols):
            etree.SubElement(grid, qn(W, "gridCol"), **{qn(W, "w"): "2000"})

    # Header rows
    if hasattr(table, "head") and table.head:
        for row in table.head.rows:
            _serialize_table_row(tbl, row, ctx, is_header=True)

    # Body rows
    if hasattr(table, "bodies"):
        for body in table.bodies:
            for row in body.rows:
                _serialize_table_row(tbl, row, ctx)


def _serialize_table_row(
    tbl: etree._Element, row: Any, ctx: _WriteContext, *, is_header: bool = False
) -> None:
    """Serialize a table row."""
    tr = etree.SubElement(tbl, W_TR)
    if is_header:
        tr_pr = etree.SubElement(tr, qn(W, "trPr"))
        etree.SubElement(tr_pr, qn(W, "tblHeader"))

    for cell in row.cells:
        tc = etree.SubElement(tr, W_TC)

        # Cell properties (col_span, row_span)
        tc_pr = etree.SubElement(tc, qn(W, "tcPr"))
        if hasattr(cell, "col_span") and cell.col_span and cell.col_span > 1:
            etree.SubElement(tc_pr, qn(W, "gridSpan"), **{qn(W, "val"): str(cell.col_span)})

        # Cell content — Cell uses .content, not .children
        cell_blocks = getattr(cell, "content", None) or getattr(cell, "children", None) or ()
        if cell_blocks:
            for child in cell_blocks:
                _serialize_block(tc, child, ctx)
        else:
            from kaos_content.traversal.visitor import extract_text

            text = extract_text(cell)
            p = etree.SubElement(tc, W_P)
            if text:
                r = etree.SubElement(p, W_R)
                t = etree.SubElement(r, W_T)
                t.text = text


def _serialize_code_block(parent: etree._Element, block: Any) -> None:
    """Serialize a code block as a monospaced paragraph."""
    text = block.value if hasattr(block, "value") else ""
    for line in text.split("\n"):
        p = etree.SubElement(parent, W_P)
        ppr = etree.SubElement(p, W_PPR)
        etree.SubElement(ppr, W_PSTYLE, **{qn(W, "val"): "Code"})
        r = etree.SubElement(p, W_R)
        rpr = etree.SubElement(r, W_RPR)
        etree.SubElement(
            rpr, qn(W, "rFonts"), **{qn(W, "ascii"): "Consolas", qn(W, "hAnsi"): "Consolas"}
        )
        t = etree.SubElement(r, W_T)
        t.set(_XML_SPACE, "preserve")
        t.text = line


def _serialize_inline(parent: etree._Element, inline: Any, ctx: _WriteContext) -> None:
    """Serialize an inline element to runs within a paragraph."""
    from kaos_content.model.inlines import (
        Code,
        Emphasis,
        FootnoteRef,
        LineBreak,
        Link,
        SoftBreak,
        Span,
        Strikethrough,
        Strong,
        Text,
    )

    if isinstance(inline, Span) and _is_revision(inline):
        _serialize_revision_span(parent, inline, ctx)
        return
    if isinstance(inline, Text):
        r = etree.SubElement(parent, W_R)
        t = etree.SubElement(r, W_T)
        t.set(_XML_SPACE, "preserve")
        t.text = inline.value
    elif isinstance(inline, Strong):
        for child in inline.children:
            r = etree.SubElement(parent, W_R)
            rpr = etree.SubElement(r, W_RPR)
            etree.SubElement(rpr, qn(W, "b"))
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            from kaos_content.traversal.visitor import extract_text

            t.text = extract_text(child)
    elif isinstance(inline, Emphasis):
        for child in inline.children:
            r = etree.SubElement(parent, W_R)
            rpr = etree.SubElement(r, W_RPR)
            etree.SubElement(rpr, qn(W, "i"))
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            from kaos_content.traversal.visitor import extract_text

            t.text = extract_text(child)
    elif isinstance(inline, Strikethrough):
        for child in inline.children:
            r = etree.SubElement(parent, W_R)
            rpr = etree.SubElement(r, W_RPR)
            etree.SubElement(rpr, qn(W, "strike"))
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            from kaos_content.traversal.visitor import extract_text

            t.text = extract_text(child)
    elif isinstance(inline, Code):
        r = etree.SubElement(parent, W_R)
        rpr = etree.SubElement(r, W_RPR)
        etree.SubElement(
            rpr, qn(W, "rFonts"), **{qn(W, "ascii"): "Consolas", qn(W, "hAnsi"): "Consolas"}
        )
        t = etree.SubElement(r, W_T)
        t.set(_XML_SPACE, "preserve")
        t.text = inline.value
    elif isinstance(inline, Link):
        _serialize_link(parent, inline, ctx)
    elif isinstance(inline, FootnoteRef):
        _serialize_footnote_ref(parent, inline, ctx)
    elif isinstance(inline, (LineBreak, SoftBreak)):
        r = etree.SubElement(parent, W_R)
        etree.SubElement(r, qn(W, "br"))
    else:
        # Fallback: extract text
        from kaos_content.traversal.visitor import extract_text

        text = extract_text(inline)
        if text:
            r = etree.SubElement(parent, W_R)
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            t.text = text


def _build_styles() -> etree._Element:
    """Build word/styles.xml with heading styles, Normal, and code."""
    root = etree.Element(qn(W, "styles"), nsmap=_W_NSMAP)

    # Default run properties
    doc_defaults = etree.SubElement(root, qn(W, "docDefaults"))
    rpr_default = etree.SubElement(doc_defaults, qn(W, "rPrDefault"))
    rpr = etree.SubElement(rpr_default, qn(W, "rPr"))
    etree.SubElement(rpr, qn(W, "rFonts"), **{qn(W, "ascii"): "Calibri", qn(W, "hAnsi"): "Calibri"})
    etree.SubElement(rpr, qn(W, "sz"), **{qn(W, "val"): "22"})

    # Normal style
    _add_style(root, "Normal", "paragraph", font_size=22)

    # Heading styles
    heading_sizes = {1: 32, 2: 28, 3: 24, 4: 22, 5: 20, 6: 20}
    for level, size in heading_sizes.items():
        _add_style(root, f"Heading{level}", "paragraph", font_size=size, bold=True)

    # Code style
    _add_style(root, "Code", "paragraph", font_name="Consolas", font_size=20)

    # Table Grid style
    tbl_style = etree.SubElement(root, qn(W, "style"))
    tbl_style.set(qn(W, "type"), "table")
    tbl_style.set(qn(W, "styleId"), "TableGrid")
    name_el = etree.SubElement(tbl_style, qn(W, "name"))
    name_el.set(qn(W, "val"), "Table Grid")
    tbl_pr = etree.SubElement(tbl_style, qn(W, "tblPr"))
    tbl_borders = etree.SubElement(tbl_pr, qn(W, "tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        etree.SubElement(
            tbl_borders,
            qn(W, side),
            **{
                qn(W, "val"): "single",
                qn(W, "sz"): "4",
                qn(W, "space"): "0",
                qn(W, "color"): "auto",
            },
        )

    return root


def _add_style(
    root: etree._Element,
    style_id: str,
    style_type: str,
    *,
    font_size: int = 22,
    bold: bool = False,
    font_name: str | None = None,
) -> None:
    """Add a style definition to styles.xml."""
    style = etree.SubElement(root, qn(W, "style"))
    style.set(qn(W, "type"), style_type)
    style.set(qn(W, "styleId"), style_id)
    name_el = etree.SubElement(style, qn(W, "name"))
    name_el.set(qn(W, "val"), style_id)

    rpr = etree.SubElement(style, W_RPR)
    etree.SubElement(rpr, qn(W, "sz"), **{qn(W, "val"): str(font_size)})
    if bold:
        etree.SubElement(rpr, qn(W, "b"))
    if font_name:
        etree.SubElement(
            rpr, qn(W, "rFonts"), **{qn(W, "ascii"): font_name, qn(W, "hAnsi"): font_name}
        )


def _build_numbering() -> etree._Element:
    """Build word/numbering.xml with bullet and number list definitions."""
    root = etree.Element(qn(W, "numbering"), nsmap=_W_NSMAP)

    # Abstract numbering 1: bullets
    abs1 = etree.SubElement(root, qn(W, "abstractNum"))
    abs1.set(qn(W, "abstractNumId"), "1")
    for lvl_idx in range(9):
        lvl = etree.SubElement(abs1, qn(W, "lvl"))
        lvl.set(qn(W, "ilvl"), str(lvl_idx))
        etree.SubElement(lvl, qn(W, "numFmt"), **{qn(W, "val"): "bullet"})
        etree.SubElement(lvl, qn(W, "lvlText"), **{qn(W, "val"): "\u2022"})

    # Abstract numbering 2: decimal
    abs2 = etree.SubElement(root, qn(W, "abstractNum"))
    abs2.set(qn(W, "abstractNumId"), "2")
    for lvl_idx in range(9):
        lvl = etree.SubElement(abs2, qn(W, "lvl"))
        lvl.set(qn(W, "ilvl"), str(lvl_idx))
        etree.SubElement(lvl, qn(W, "numFmt"), **{qn(W, "val"): "decimal"})
        etree.SubElement(lvl, qn(W, "lvlText"), **{qn(W, "val"): f"%{lvl_idx + 1}."})
        etree.SubElement(lvl, qn(W, "start"), **{qn(W, "val"): "1"})

    # Numbering instances
    num1 = etree.SubElement(root, qn(W, "num"))
    num1.set(qn(W, "numId"), "1")
    etree.SubElement(num1, qn(W, "abstractNumId"), **{qn(W, "val"): "1"})

    num2 = etree.SubElement(root, qn(W, "num"))
    num2.set(qn(W, "numId"), "2")
    etree.SubElement(num2, qn(W, "abstractNumId"), **{qn(W, "val"): "2"})

    return root


def _build_empty_numbering() -> etree._Element:
    """Build an empty word/numbering.xml."""
    return etree.Element(qn(W, "numbering"), nsmap=_W_NSMAP)


# ---------------------------------------------------------------------------
# Hyperlinks (Phase A)
# ---------------------------------------------------------------------------


def _serialize_link(parent: etree._Element, link: Any, ctx: _WriteContext) -> None:
    """Emit a w:hyperlink element backed by a document rels entry.

    External URLs are added to ``word/_rels/document.xml.rels`` as
    TargetMode="External" relationships. Internal references (``"#bookmark"``)
    use w:anchor instead of r:id.
    """
    url = getattr(link, "url", "") or ""
    link_text = _extract_link_text(link)

    if not url:
        _emit_styled_text(parent, link_text)
        return

    # Internal anchor reference
    if url.startswith("#"):
        hl = etree.SubElement(parent, qn(W, "hyperlink"), **{qn(W, "anchor"): url[1:]})
        _emit_link_runs(hl, link)
        return

    # External URL — reuse rel if seen before
    if ctx.doc_rels is None:
        # Fallback: no rel manager available, emit styled text
        _emit_styled_text(parent, link_text)
        return

    rel_id = ctx.hyperlink_urls.get(url)
    if rel_id is None:
        rel = ctx.doc_rels.add(_RT_HYPERLINK, url, external=True)
        rel_id = rel.id
        ctx.hyperlink_urls[url] = rel_id

    hl = etree.SubElement(parent, qn(W, "hyperlink"), **{qn(R, "id"): rel_id})
    _emit_link_runs(hl, link)


def _extract_link_text(link: Any) -> str:
    """Plain text from a Link's children."""
    from kaos_content.traversal.visitor import extract_text

    return extract_text(link)


def _emit_link_runs(parent: etree._Element, link: Any) -> None:
    """Emit the run(s) for a link's visible text with hyperlink styling.

    Hyperlinks in OOXML contain one or more runs; the runs typically carry
    the "Hyperlink" rStyle for blue+underlined appearance.
    """
    text = _extract_link_text(link)
    r = etree.SubElement(parent, W_R)
    rpr = etree.SubElement(r, W_RPR)
    etree.SubElement(rpr, qn(W, "rStyle"), **{qn(W, "val"): "Hyperlink"})
    etree.SubElement(rpr, qn(W, "color"), **{qn(W, "val"): "0563C1"})
    etree.SubElement(rpr, qn(W, "u"), **{qn(W, "val"): "single"})
    t = etree.SubElement(r, W_T)
    t.set(_XML_SPACE, "preserve")
    t.text = text


def _emit_styled_text(parent: etree._Element, text: str) -> None:
    """Fallback: styled text when no URL is available."""
    r = etree.SubElement(parent, W_R)
    rpr = etree.SubElement(r, W_RPR)
    etree.SubElement(rpr, qn(W, "color"), **{qn(W, "val"): "0563C1"})
    etree.SubElement(rpr, qn(W, "u"), **{qn(W, "val"): "single"})
    t = etree.SubElement(r, W_T)
    t.set(_XML_SPACE, "preserve")
    t.text = text


def _serialize_footnote_ref(parent: etree._Element, ref: Any, ctx: _WriteContext) -> None:
    """Emit a w:footnoteReference or w:endnoteReference run.

    The reader prefixes endnote identifiers with ``"en-"``. We strip that
    prefix here and emit the appropriate reference element.
    """
    identifier = getattr(ref, "identifier", "") or ""
    if identifier.startswith("en-"):
        note_id = identifier[3:]
        reference_tag = "endnoteReference"
        style = "EndnoteReference"
        known = ctx.endnotes
        has_flag = "has_endnotes"
    else:
        note_id = identifier
        reference_tag = "footnoteReference"
        style = "FootnoteReference"
        known = ctx.footnotes
        has_flag = "has_footnotes"

    # Ensure the note collection flag is set even if the reference appears
    # in the body without a corresponding entry in doc.footnotes (the reader
    # can produce this shape; we still emit a valid reference).
    if note_id and note_id not in known:
        known[note_id] = ()
    setattr(ctx, has_flag, True)

    r = etree.SubElement(parent, W_R)
    rpr = etree.SubElement(r, W_RPR)
    etree.SubElement(rpr, qn(W, "rStyle"), **{qn(W, "val"): style})
    etree.SubElement(r, qn(W, reference_tag), **{qn(W, "id"): note_id})


# ---------------------------------------------------------------------------
# Tracked-change revision serialization (Phase D)
# ---------------------------------------------------------------------------

# Map rev-* class → (OOXML element tag name for inline/block revisions)
_REV_CLASS_TO_TAG: dict[str, str] = {
    "rev-ins": "ins",
    "rev-del": "del",
    "rev-move-from": "moveFrom",
    "rev-move-to": "moveTo",
}


def _is_revision(node: Any) -> bool:
    """Whether a node has one of the rev-* classes on its Attr.classes."""
    attr = getattr(node, "attr", None)
    if attr is None:
        return False
    return any(cls in _REV_CLASS_TO_TAG for cls in getattr(attr, "classes", ()) or ())


def _revision_tag(node: Any) -> tuple[str, str] | None:
    """Return (OOXML tag name, rev class) for a revision-marked node, or None."""
    attr = getattr(node, "attr", None)
    if attr is None:
        return None
    for cls in getattr(attr, "classes", ()) or ():
        if cls in _REV_CLASS_TO_TAG:
            return _REV_CLASS_TO_TAG[cls], cls
    return None


def _revision_wrapper_attrs(node: Any) -> dict[str, str]:
    """Build w:id / w:author / w:date / w:name attributes from Attr.kv."""
    kv = (getattr(node.attr, "kv", None) or {}) if getattr(node, "attr", None) else {}
    attrs: dict[str, str] = {}
    if "rev:id" in kv:
        attrs[qn(W, "id")] = str(kv["rev:id"])
    if "rev:author" in kv:
        attrs[qn(W, "author")] = str(kv["rev:author"])
    if "rev:date" in kv:
        attrs[qn(W, "date")] = str(kv["rev:date"])
    if "rev:move-name" in kv:
        attrs[qn(W, "name")] = str(kv["rev:move-name"])
    return attrs


def _serialize_revision_span(parent: etree._Element, span: Any, ctx: _WriteContext) -> None:
    """Emit a <w:ins> / <w:del> / <w:moveFrom> / <w:moveTo> wrapper around runs.

    The wrapper carries revision metadata (id, author, date, move name).
    For ``rev-del`` / ``rev-move-from`` wrappers, the contained text nodes
    are post-processed so their ``<w:t>`` becomes ``<w:delText>`` per
    OOXML §17.16.2 (though most Word readers accept either).
    """
    tag_info = _revision_tag(span)
    if tag_info is None:
        # Fallback: just emit children as if the Span weren't there
        for child in getattr(span, "children", ()):
            _serialize_inline(parent, child, ctx)
        return

    tag, rev_class = tag_info
    wrapper = etree.SubElement(parent, qn(W, tag), **_revision_wrapper_attrs(span))
    for child in getattr(span, "children", ()):
        _serialize_inline(wrapper, child, ctx)

    # OOXML spec: deleted text runs use <w:delText> instead of <w:t>.
    # Rename for strict compliance.
    if rev_class in ("rev-del", "rev-move-from"):
        for t_el in wrapper.findall(f".//{W_T}"):
            t_el.tag = qn(W, "delText")


def _serialize_revision_div(parent: etree._Element, div: Any, ctx: _WriteContext) -> None:
    """Emit a block-level revision wrapper at the body level.

    The Div's children are block nodes that must be serialized inside the
    wrapper. The wrapper appears as a direct body child.
    """
    tag_info = _revision_tag(div)
    if tag_info is None:
        for child in getattr(div, "children", ()):
            _serialize_block(parent, child, ctx)
        return

    tag, rev_class = tag_info
    wrapper = etree.SubElement(parent, qn(W, tag), **_revision_wrapper_attrs(div))
    for child in getattr(div, "children", ()):
        _serialize_block(wrapper, child, ctx)

    if rev_class in ("rev-del", "rev-move-from"):
        for t_el in wrapper.findall(f".//{W_T}"):
            t_el.tag = qn(W, "delText")


# ---------------------------------------------------------------------------
# Footnotes / Endnotes (Phase B)
# ---------------------------------------------------------------------------

# OOXML reserves IDs 0 and -1 for the separator and continuation separator.
_FN_SEPARATOR_ID = "0"
_FN_CONTINUATION_ID = "-1"


def _prepare_notes(doc: Any, ctx: _WriteContext) -> None:
    """Scan document.footnotes and document.annotations for write-back content.

    Footnotes keyed with ``"en-"`` prefix (added by the reader) are endnotes.
    COMMENT annotations are collected with assigned sequential IDs for
    emission into word/comments.xml.
    """
    fns = getattr(doc, "footnotes", None) or {}
    if fns:
        for identifier, blocks in fns.items():
            if identifier.startswith("en-"):
                en_id = identifier[3:]
                ctx.endnotes[en_id] = tuple(blocks)
            else:
                ctx.footnotes[identifier] = tuple(blocks)

        if ctx.footnotes:
            ctx.has_footnotes = True
        if ctx.endnotes:
            ctx.has_endnotes = True

    # Collect COMMENT annotations. The reader stores (comment_id, metadata)
    # in annotation.body; we renumber sequentially so IDs match what Word
    # expects (0-indexed and gap-free).
    from kaos_content.model.annotation import AnnotationType

    annotations = getattr(doc, "annotations", None) or ()
    next_id = 0
    for ann in annotations:
        if getattr(ann, "type", None) != AnnotationType.COMMENT:
            continue
        ctx.comments.append((next_id, dict(ann.body or {})))
        next_id += 1


def _build_notes_xml(ctx: _WriteContext, *, kind: str) -> etree._Element:
    """Build word/footnotes.xml or word/endnotes.xml.

    Both parts always include the special separator (id=0) and
    continuation separator (id=-1) elements required by the OOXML spec —
    Word misbehaves without them. Real notes are emitted with IDs starting
    at 1, matching the reader's identifier strings.
    """
    if kind == "footnotes":
        root_tag = qn(W, "footnotes")
        item_tag = qn(W, "footnote")
        source = ctx.footnotes
        separator_tag = qn(W, "separator")
        continuation_tag = qn(W, "continuationSeparator")
        ref_style = "FootnoteReference"
    else:
        root_tag = qn(W, "endnotes")
        item_tag = qn(W, "endnote")
        source = ctx.endnotes
        separator_tag = qn(W, "separator")
        continuation_tag = qn(W, "continuationSeparator")
        ref_style = "EndnoteReference"

    root = etree.Element(root_tag, nsmap=_W_NSMAP)

    # Required separator (id=0)
    sep = etree.SubElement(
        root, item_tag, **{qn(W, "id"): _FN_SEPARATOR_ID, qn(W, "type"): "separator"}
    )
    _add_separator_paragraph(sep, separator_tag)

    # Required continuation separator (id=-1)
    cont = etree.SubElement(
        root, item_tag, **{qn(W, "id"): _FN_CONTINUATION_ID, qn(W, "type"): "continuationSeparator"}
    )
    _add_separator_paragraph(cont, continuation_tag)

    # Actual notes
    for identifier, blocks in source.items():
        note = etree.SubElement(root, item_tag, **{qn(W, "id"): identifier})
        if not blocks:
            # Always emit at least one paragraph
            etree.SubElement(note, W_P)
            continue
        # Emit the reference marker on the first paragraph
        for i, block in enumerate(blocks):
            _serialize_note_block(note, block, ref_style=ref_style if i == 0 else None)

    return root


def _add_separator_paragraph(parent: etree._Element, marker_tag: str) -> None:
    """Add the w:p containing the required <w:separator/> or <w:continuationSeparator/>."""
    p = etree.SubElement(parent, W_P)
    r = etree.SubElement(p, W_R)
    etree.SubElement(r, marker_tag)


def _serialize_note_block(parent: etree._Element, block: Any, *, ref_style: str | None) -> None:
    """Serialize a block (Paragraph) inside a footnote/endnote.

    The first paragraph gets a leading run with ``w:footnoteRef`` (or
    ``w:endnoteRef``) as the reference marker.
    """
    nt = getattr(block, "node_type", None)
    if nt != "paragraph":
        # Fall back to a plain paragraph with the block's text
        from kaos_content.traversal.visitor import extract_text

        p = etree.SubElement(parent, W_P)
        text = extract_text(block)
        if text:
            r = etree.SubElement(p, W_R)
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            t.text = text
        return

    p = etree.SubElement(parent, W_P)
    ppr = etree.SubElement(p, W_PPR)
    etree.SubElement(
        ppr,
        W_PSTYLE,
        **{qn(W, "val"): "FootnoteText" if ref_style == "FootnoteReference" else "EndnoteText"},
    )

    if ref_style is not None:
        ref_run = etree.SubElement(p, W_R)
        ref_rpr = etree.SubElement(ref_run, W_RPR)
        etree.SubElement(ref_rpr, qn(W, "rStyle"), **{qn(W, "val"): ref_style})
        ref_tag = "footnoteRef" if ref_style == "FootnoteReference" else "endnoteRef"
        etree.SubElement(ref_run, qn(W, ref_tag))
        # Tab separator
        tab_run = etree.SubElement(p, W_R)
        etree.SubElement(tab_run, qn(W, "tab"))

    # The paragraph children are inlines — serialize them onto the same w:p
    ctx_stub = _WriteContext()
    for inline in getattr(block, "children", ()):
        _serialize_inline(p, inline, ctx_stub)


def _build_footnotes(ctx: _WriteContext) -> etree._Element:
    """Build word/footnotes.xml."""
    return _build_notes_xml(ctx, kind="footnotes")


def _build_endnotes(ctx: _WriteContext) -> etree._Element:
    """Build word/endnotes.xml."""
    return _build_notes_xml(ctx, kind="endnotes")


# ---------------------------------------------------------------------------
# Comments (Phase C)
# ---------------------------------------------------------------------------


def _build_comments(ctx: _WriteContext) -> etree._Element:
    """Build word/comments.xml from collected comment metadata.

    Each comment is emitted as ``<w:comment w:id w:author w:date w:initials>``
    with the body text wrapped in a ``<w:p>``.
    """
    root = etree.Element(qn(W, "comments"), nsmap=_W_NSMAP)
    for comment_id, meta in ctx.comments:
        c = etree.SubElement(
            root,
            qn(W, "comment"),
            **{
                qn(W, "id"): str(comment_id),
                qn(W, "author"): str(meta.get("author", "") or ""),
                qn(W, "date"): str(meta.get("date", "") or ""),
                qn(W, "initials"): str(meta.get("initials", "") or ""),
            },
        )
        text = str(meta.get("text", "") or "")
        p = etree.SubElement(c, W_P)
        if text:
            r = etree.SubElement(p, W_R)
            t = etree.SubElement(r, W_T)
            t.set(_XML_SPACE, "preserve")
            t.text = text
    return root


def _build_core_properties(doc: Any) -> etree._Element:
    """Build docProps/core.xml with Dublin Core metadata."""
    import datetime

    root = etree.Element(
        f"{{{_CP_NS}}}coreProperties",
        nsmap={
            "cp": _CP_NS,
            "dc": _DC_NS,
            "dcterms": _DCTERMS_NS,
        },
    )

    title = doc.metadata.title or ""
    if title:
        dc_title = etree.SubElement(root, f"{{{_DC_NS}}}title")
        dc_title.text = title

    creator = ""
    if hasattr(doc.metadata, "source") and doc.metadata.source:
        creator = getattr(doc.metadata.source, "creator", "") or ""
    if not creator:
        creator = "kaos-office"
    dc_creator = etree.SubElement(root, f"{{{_DC_NS}}}creator")
    dc_creator.text = creator

    now = datetime.datetime.now(datetime.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    created = etree.SubElement(root, f"{{{_DCTERMS_NS}}}created")
    created.set(f"{{{_XSI_NS}}}type", "dcterms:W3CDTF")
    created.text = now
    modified = etree.SubElement(root, f"{{{_DCTERMS_NS}}}modified")
    modified.set(f"{{{_XSI_NS}}}type", "dcterms:W3CDTF")
    modified.text = now

    return root
