"""DOCX Reader — parse_docx() → ContentDocument.

Two-pass architecture:
  Pass 1: Load metadata (styles, numbering, relationships, document metadata)
  Pass 2: Walk document body with tag dispatch → DocumentBuilder

Supports: paragraphs, headings, lists, tables, images, hyperlinks,
footnotes, endnotes, comments, track changes (accept/skip), page breaks.
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass, field
from pathlib import Path

from kaos_content import ContentDocument
from kaos_content.builders import DocumentBuilder
from kaos_content.model.annotation import AnnotationType
from kaos_content.model.blocks import Table
from kaos_content.model.inlines import (
    Emphasis,
    Image,
    Inline,
    LineBreak,
    Link,
    Strikethrough,
    Strong,
    Subscript,
    Superscript,
    Text,
    Underline,
)
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
    RT_COMMENTS,
    RT_ENDNOTES,
    RT_FOOTNOTES,
    RT_OFFICE_DOCUMENT,
    W_ANCHOR,
    W_B,
    W_BODY,
    W_BOOKMARK_START,
    W_BR,
    W_COMMENT,
    W_DEL,
    W_DRAWING,
    W_ENDNOTE,
    W_ENDNOTE_REFERENCE,
    W_ENDNOTES,
    W_FOOTNOTE,
    W_FOOTNOTE_REFERENCE,
    W_FOOTNOTES,
    W_GRIDCOL,
    W_GRIDSPAN,
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
    qn,
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


def parse_docx(path: str | Path) -> ContentDocument:
    """Parse a DOCX file into a ContentDocument.

    Args:
        path: Path to the .docx file.

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
        builder.set_source(uri=source_uri, mime_type=DOCX_MIME_TYPE)
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
        )

        for child in body:
            _process_body_child(child, ctx)

        # Flush any trailing open lists
        _flush_open_lists(ctx)

        # Add footnote/endnote content
        _add_footnotes(ctx)

        # Add comment annotations
        _add_comment_annotations(ctx)

        return builder.build()


# --------------------------------------------------------------------------
# Body-level dispatch
# --------------------------------------------------------------------------


def _process_body_child(el: etree._Element, ctx: ParseContext) -> None:
    """Dispatch a top-level body child element."""
    tag = el.tag

    if tag == W_P:
        _handle_paragraph(el, ctx)
    elif tag == W_TBL:
        _flush_open_lists(ctx)
        _handle_table(el, ctx)
    elif tag == W_SECTPR:
        _flush_open_lists(ctx)
        # Section properties — emit page break for section boundaries
        # (Could extract page size, margins for provenance in future)
    elif tag == W_SDT:
        # Structured document tag — unwrap and process contents
        sdt_content = el.find(W_SDTCONTENT)
        if sdt_content is not None:
            for child in sdt_content:
                _process_body_child(child, ctx)
    elif tag == W_BOOKMARK_START:
        bm_id = el.get(W_ID)
        bm_name = el.get(qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "name"))
        if bm_id and bm_name:
            ctx.bookmarks[bm_id] = bm_name
    elif tag == W_INS:
        # Track change: insertion at body level — include content
        for child in el:
            _process_body_child(child, ctx)
    elif tag == W_DEL:
        pass  # Track change: deletion at body level — skip
    elif tag == W_MOVE_TO:
        # Accept move — include the moveTo content
        for child in el:
            _process_body_child(child, ctx)
    elif tag == W_MOVE_FROM:
        pass  # Skip moveFrom content (it's in moveTo now)
    # Unknown elements silently skipped


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
        elif tag == W_INS:
            # Track change: insertion — include the content
            for sub in child:
                if sub.tag == W_R:
                    _process_run(sub, ctx, inlines)
                elif sub.tag == W_HYPERLINK:
                    _process_hyperlink(sub, ctx, inlines)
        elif tag == W_DEL:
            pass  # Track change: deletion — skip content
        elif tag == W_MOVE_TO:
            for sub in child:
                if sub.tag == W_R:
                    _process_run(sub, ctx, inlines)
        elif tag == W_MOVE_FROM:
            pass  # Skip moveFrom
        elif tag == W_BOOKMARK_START:
            bm_id = child.get(W_ID)
            bm_name = child.get(
                qn("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "name")
            )
            if bm_id and bm_name:
                ctx.bookmarks[bm_id] = bm_name
        elif tag == W_SDT:
            # Inline SDT — unwrap
            sdt_content = child.find(W_SDTCONTENT)
            if sdt_content is not None:
                for sub in sdt_content:
                    if sub.tag == W_R:
                        _process_run(sub, ctx, inlines)
                    elif sub.tag == W_HYPERLINK:
                        _process_hyperlink(sub, ctx, inlines)
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

        if tag == W_T:
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


def _extract_image(drawing_el: etree._Element, ctx: ParseContext) -> Image | None:
    """Extract an image from a w:drawing element."""
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

    # Future: extract dimensions from wp:extent for provenance bounding box

    # Build image URI
    media_name = target.rsplit("/", 1)[-1] if "/" in target else target
    src = f"docx://{media_name}"

    return Image(src=src, alt=alt or None, title=title or None)


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
