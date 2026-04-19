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
    W_BODY,
    W_P,
    W_PPR,
    W_PSTYLE,
    W_R,
    W_RPR,
    W_T,
    W_TBL,
    W_TC,
    W_TR,
    R,
    W,
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

# Relationship types
_RT_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
_RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
_RT_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
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
    ctx = _WriteContext()
    document_xml = _build_document(doc, ctx)
    writer.add_xml_part("word/document.xml", document_xml)

    # Build styles.xml
    writer.add_xml_part("word/styles.xml", _build_styles())

    # Build numbering.xml
    if ctx.has_lists:
        writer.add_xml_part("word/numbering.xml", _build_numbering())
    else:
        writer.add_xml_part("word/numbering.xml", _build_empty_numbering())

    # Build core properties
    writer.add_xml_part("docProps/core.xml", _build_core_properties(doc))

    return writer.save_bytes()


class _WriteContext:
    """Accumulates state during serialization."""

    __slots__ = ("has_lists", "hyperlink_rels", "list_counter")

    def __init__(self) -> None:
        self.has_lists = False
        self.list_counter = 0
        self.hyperlink_rels: list[tuple[str, str]] = []  # (rel_id, url)


def _build_document(doc: Any, ctx: _WriteContext) -> etree._Element:
    """Build word/document.xml from ContentDocument."""
    root = etree.Element(qn(W, "document"), nsmap=_W_NSMAP)
    body = etree.SubElement(root, W_BODY)

    for block in doc.body:
        _serialize_block(body, block, ctx)

    # Default section properties (Letter, 1" margins)
    sect_pr = etree.SubElement(body, qn(W, "sectPr"))
    etree.SubElement(sect_pr, qn(W, "pgSz"), **{qn(W, "w"): "12240", qn(W, "h"): "15840"})
    etree.SubElement(
        sect_pr,
        qn(W, "pgMar"),
        **{
            qn(W, "top"): "1440",
            qn(W, "right"): "1440",
            qn(W, "bottom"): "1440",
            qn(W, "left"): "1440",
        },
    )

    return root


def _serialize_block(parent: etree._Element, block: Any, ctx: _WriteContext) -> None:
    """Serialize a single block to OOXML elements."""
    from kaos_content.model.blocks import (
        BlockQuote,
        BulletList,
        CodeBlock,
        Heading,
        OrderedList,
        PageBreak,
        Paragraph,
        Table,
        ThematicBreak,
    )

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
        LineBreak,
        Link,
        SoftBreak,
        Strikethrough,
        Strong,
        Text,
    )

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
        # Serialize as plain text with URL annotation for now
        # Full hyperlink support requires document-level relationship
        from kaos_content.traversal.visitor import extract_text

        link_text = extract_text(inline)
        r = etree.SubElement(parent, W_R)
        rpr = etree.SubElement(r, W_RPR)
        etree.SubElement(rpr, qn(W, "color"), **{qn(W, "val"): "0563C1"})
        etree.SubElement(rpr, qn(W, "u"), **{qn(W, "val"): "single"})
        t = etree.SubElement(r, W_T)
        t.set(_XML_SPACE, "preserve")
        t.text = link_text
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
