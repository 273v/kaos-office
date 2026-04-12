"""PPTX reader — parse PowerPoint presentations into ContentDocument.

Dual approach: python-pptx for high-level shape traversal, OPC/lxml fallback
for SmartArt (no OSS Python tool handles SmartArt text extraction).

Each slide is linearized as a Div block. Shapes are sorted by reading position
(top, left). Titles become headings, body text becomes paragraphs with bullet
detection, tables and charts become Table blocks.

Entry point: parse_pptx(path) → ContentDocument
"""

from __future__ import annotations

import contextlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

from kaos_content.builders.builder import DocumentBuilder
from kaos_content.model.blocks import Table
from kaos_content.model.document import ContentDocument
from kaos_content.model.inlines import Emphasis, Inline, Link, Strong, Text
from kaos_content.model.table import Cell, Row, TableSection
from kaos_core.logging import get_logger
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import (
    A_GRAPHIC,
    A_GRAPHIC_DATA,
    GD_DIAGRAM,
    PPTX_MIME_TYPE,
)
from kaos_office.opc.package import OPCPackage
from kaos_office.pptx.smartart import extract_smartart_texts

if TYPE_CHECKING:
    from pptx.chart.chart import Chart
    from pptx.presentation import Presentation
    from pptx.shapes.base import BaseShape
    from pptx.slide import Slide
    from pptx.table import Table as PptxTable
    from pptx.text.text import _Paragraph

logger = get_logger(__name__)

# Placeholder types to skip (metadata, not content)
_SKIP_PLACEHOLDER_TYPES = frozenset(
    {
        "DATE",  # dt
        "FOOTER",  # ftr
        "HEADER",  # hdr
        "SLIDE_NUMBER",  # sldNum
        "SLIDE_IMAGE",  # sldImg (in notes slides)
    }
)

# Placeholder types that map to headings
_TITLE_PLACEHOLDER_TYPES = frozenset(
    {
        "TITLE",  # title
        "CENTER_TITLE",  # ctrTitle
        "VERTICAL_TITLE",  # vertTitle
    }
)

_SUBTITLE_PLACEHOLDER_TYPES = frozenset(
    {
        "SUBTITLE",  # subTitle
    }
)

# Placeholder types that typically have inherited bullets from layout
_BODY_PLACEHOLDER_TYPES = frozenset(
    {
        "BODY",  # body
        "OBJECT",  # obj
        "VERTICAL_BODY",  # vertTx
    }
)

# Auto-numbering types that indicate ordered lists
_ORDERED_AUTONUM_TYPES = frozenset(
    {
        "arabicPeriod",
        "arabicParenR",
        "arabicParenBoth",
        "arabicPlain",
        "alphaLcPeriod",
        "alphaUcPeriod",
        "alphaLcParenR",
        "alphaUcParenR",
        "alphaLcParenBoth",
        "alphaUcParenBoth",
        "romanLcPeriod",
        "romanUcPeriod",
        "romanLcParenR",
        "romanUcParenR",
        "romanLcParenBoth",
        "romanUcParenBoth",
    }
)


@dataclass
class ParseContext:
    """Holds state during PPTX parsing."""

    builder: DocumentBuilder
    pkg: OPCPackage
    source_uri: str
    # Track list nesting state per text frame
    list_stack: list[_ListState] = field(default_factory=list)


@dataclass
class _ListState:
    """Tracks an open list level during paragraph processing."""

    level: int
    ordered: bool
    item_open: bool = False


def parse_pptx(path: str | Path) -> ContentDocument:
    """Parse a PPTX file into a ContentDocument.

    Each slide becomes a Div block with slide_number attribute.
    Shapes are sorted by reading position (top, left).
    Titles/subtitles become headings. Body text is processed with
    bullet detection. Tables, charts, images, SmartArt, and speaker
    notes are all extracted.

    Args:
        path: Path to the PPTX file.

    Returns:
        ContentDocument with linearized slide content.
    """
    from pptx import Presentation

    path = Path(path).resolve()
    source_uri = path.as_uri()

    prs = Presentation(str(path))

    # Open OPC package for SmartArt fallback
    with OPCPackage.open(path) as pkg:
        # Extract metadata
        core_xml = pkg.read_part("docProps/core.xml") if pkg.has_part("docProps/core.xml") else None
        app_xml = pkg.read_part("docProps/app.xml") if pkg.has_part("docProps/app.xml") else None

        title = _extract_metadata_title(core_xml)
        builder = DocumentBuilder(title=title)
        builder.set_source(uri=source_uri, mime_type=PPTX_MIME_TYPE)
        _apply_metadata(builder, core_xml, app_xml, prs)

        ctx = ParseContext(builder=builder, pkg=pkg, source_uri=source_uri)

        for slide_num, slide in enumerate(prs.slides, 1):
            _process_slide(slide, slide_num, ctx)

    return builder.build()


def _extract_metadata_title(core_xml: bytes | None) -> str | None:
    """Extract title from core.xml for DocumentBuilder constructor."""
    if core_xml is None:
        return None
    from kaos_office.ooxml.namespace import DC

    try:
        root = etree.fromstring(core_xml)
        title_el = root.find(f"{{{DC}}}title")
        if title_el is not None and title_el.text:
            return title_el.text.strip()
    except Exception as exc:
        logger.debug("Failed to extract metadata title from core.xml: %s", exc)
    return None


def _apply_metadata(
    builder: DocumentBuilder,
    core_xml: bytes | None,
    app_xml: bytes | None,
    prs: Presentation,
) -> None:
    """Apply document metadata from core.xml and app.xml."""
    from kaos_office.docx.metadata import DocxMetadata

    try:
        meta = DocxMetadata.from_xml(core_xml, app_xml)
        kwargs: dict = {}
        if meta.creator:
            kwargs["authors"] = (meta.creator,)
        if meta.created:
            kwargs["date"] = meta.created
        if meta.description:
            kwargs["description"] = meta.description

        extra: dict = {}
        slide_count = len(prs.slides)
        extra["slide_count"] = str(slide_count)
        if prs.slide_width and prs.slide_height:
            extra["slide_width_emu"] = str(prs.slide_width)
            extra["slide_height_emu"] = str(prs.slide_height)
        if meta.application:
            extra["application"] = meta.application
        if meta.company:
            extra["company"] = meta.company
        if extra:
            kwargs["extra"] = extra

        if kwargs:
            builder.set_metadata(**kwargs)
    except Exception as exc:
        logger.debug("Failed to apply document metadata: %s", exc)


def _process_slide(slide: Slide, slide_num: int, ctx: ParseContext) -> None:
    """Process a single slide into builder blocks."""
    ctx.builder.begin_div(slide_number=str(slide_num), classes="slide")

    # Sort shapes by reading position (top, then left)
    shapes = _sort_shapes(list(slide.shapes))

    for shape in shapes:
        _process_shape(shape, slide, ctx)

    ctx.builder.end()  # Close slide div

    # Speaker notes
    _process_notes(slide, ctx)


def _sort_shapes(shapes: list[BaseShape]) -> list[BaseShape]:
    """Sort shapes by reading position: top first, then left."""
    return sorted(shapes, key=lambda s: (s.top or 0, s.left or 0))


def _get_placeholder_type_name(shape: BaseShape) -> str | None:
    """Get placeholder type name, or None if not a placeholder."""
    try:
        ph = shape.placeholder_format
        if ph is not None and ph.type is not None:
            return ph.type.name
    except (ValueError, AttributeError):
        pass
    return None


def _process_shape(shape: BaseShape, slide: Slide, ctx: ParseContext) -> None:
    """Dispatch shape to appropriate handler based on type."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    # Check placeholder type for skip/heading logic
    ph_type = _get_placeholder_type_name(shape)
    if ph_type in _SKIP_PLACEHOLDER_TYPES:
        return

    # Group shapes: recurse into children
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        _process_group(shape, ctx)
        return

    # Connector lines: no content
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.LINE or st == MSO_SHAPE_TYPE.FREEFORM:
        return

    # Table
    if shape.has_table:
        _process_table(shape.table, ctx)  # ty: ignore[unresolved-attribute]
        return

    # Chart
    if shape.has_chart:
        _process_chart(shape.chart, ctx)  # ty: ignore[unresolved-attribute]
        return

    # Check for SmartArt (graphicFrame with diagram URI)
    if _is_smartart(shape):
        _process_smartart(shape, slide, ctx)
        return

    # Picture
    if st == MSO_SHAPE_TYPE.PICTURE:
        _process_picture(shape, ctx)
        return

    # Text shapes (including title/subtitle/body placeholders)
    if shape.has_text_frame:
        if ph_type in _TITLE_PLACEHOLDER_TYPES:
            text = shape.text_frame.text.strip()  # ty: ignore[unresolved-attribute]
            if text:
                ctx.builder.heading(1, text)
        elif ph_type in _SUBTITLE_PLACEHOLDER_TYPES:
            text = shape.text_frame.text.strip()  # ty: ignore[unresolved-attribute]
            if text:
                ctx.builder.heading(2, text)
        else:
            # Body/object placeholders inherit bullets from the layout
            is_body = ph_type in _BODY_PLACEHOLDER_TYPES
            _process_text_frame(shape.text_frame, ctx, is_body_placeholder=is_body)  # ty: ignore[unresolved-attribute]


def _process_group(group_shape: BaseShape, ctx: ParseContext) -> None:
    """Process group shape by recursing into sorted children."""
    children = _sort_shapes(list(group_shape.shapes))  # ty: ignore[unresolved-attribute]
    for child in children:
        # Groups don't have slide context for SmartArt, pass None
        _process_shape_simple(child, ctx)


def _process_shape_simple(shape: BaseShape, ctx: ParseContext) -> None:
    """Process a shape without slide context (for group children)."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    ph_type = _get_placeholder_type_name(shape)
    if ph_type in _SKIP_PLACEHOLDER_TYPES:
        return

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        _process_group(shape, ctx)
        return

    if shape.has_table:
        _process_table(shape.table, ctx)  # ty: ignore[unresolved-attribute]
        return

    if shape.has_chart:
        _process_chart(shape.chart, ctx)  # ty: ignore[unresolved-attribute]
        return

    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.PICTURE:
        _process_picture(shape, ctx)
        return

    if shape.has_text_frame:
        _process_text_frame(shape.text_frame, ctx)  # ty: ignore[unresolved-attribute]


def _process_text_frame(
    text_frame: object, ctx: ParseContext, *, is_body_placeholder: bool = False
) -> None:
    """Process a text frame's paragraphs with bullet/list detection.

    Args:
        text_frame: The text frame to process.
        ctx: Parse context.
        is_body_placeholder: If True, paragraphs without explicit bullet markers
            are treated as bullet list items (bullets inherited from layout).
    """
    paragraphs = text_frame.paragraphs  # type: ignore[attr-defined]  # ty: ignore[unresolved-attribute]
    if not paragraphs:
        return

    ctx.list_stack.clear()

    for para in paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Detect bullet/list from paragraph XML
        bullet_info = _get_bullet_info(para, is_body_placeholder=is_body_placeholder)

        if bullet_info is not None:
            level, ordered = bullet_info
            _handle_list_paragraph(para, level, ordered, ctx)
        else:
            # Close any open lists
            _flush_lists(ctx)
            # Regular paragraph with formatting
            inlines = _runs_to_inlines(para)
            if inlines:
                ctx.builder.paragraph(*inlines)

    # Flush any remaining open lists
    _flush_lists(ctx)


def _get_bullet_info(
    para: _Paragraph, *, is_body_placeholder: bool = False
) -> tuple[int, bool] | None:
    """Detect bullet/list info from paragraph properties.

    Args:
        para: The paragraph to check.
        is_body_placeholder: If True, treat paragraphs with lvl attribute but
            no explicit bullet markers as bullet items (inherited from layout).

    Returns (level, ordered) or None if not a list item.
    """
    from kaos_office.ooxml.namespace import A_BU_AUTO_NUM, A_BU_CHAR, A_BU_NONE, A_PPR

    pPr = para._p.find(A_PPR)
    if pPr is None:
        return None

    # Explicit no-bullet
    if pPr.find(A_BU_NONE) is not None:
        return None

    level = int(pPr.get("lvl", "0"))

    # Check for explicit bullet character
    bu_char = pPr.find(A_BU_CHAR)
    if bu_char is not None:
        return (level, False)

    # Check for auto-numbering
    bu_auto = pPr.find(A_BU_AUTO_NUM)
    if bu_auto is not None:
        auto_type = bu_auto.get("type", "")
        ordered = auto_type in _ORDERED_AUTONUM_TYPES
        return (level, ordered)

    # Body placeholders inherit bullets from the slide layout.
    # Only treat as bullet if lvl is explicitly set in the XML.
    # Level 0 without explicit lvl is ambiguous — likely a plain text
    # paragraph that precedes a bulleted sub-list.
    if is_body_placeholder and pPr.get("lvl") is not None:
        return (level, False)

    return None


def _handle_list_paragraph(para: _Paragraph, level: int, ordered: bool, ctx: ParseContext) -> None:
    """Handle a paragraph that is a list item with bullet detection."""
    stack = ctx.list_stack

    # Close deeper levels
    while stack and stack[-1].level > level:
        if stack[-1].item_open:
            ctx.builder.end()  # Close list item
            stack[-1].item_open = False
        ctx.builder.end()  # Close list
        stack.pop()

    # If same level but different list type, close and reopen
    if stack and stack[-1].level == level and stack[-1].ordered != ordered:
        if stack[-1].item_open:
            ctx.builder.end()  # Close list item
            stack[-1].item_open = False
        ctx.builder.end()  # Close list
        stack.pop()

    # Open new list levels as needed
    if not stack or stack[-1].level < level:
        # If a list item is open at a parent level, nest inside it
        if stack and stack[-1].item_open:
            pass  # Leave item open — sub-list nests inside it

        # Open the list directly at the target level.
        # PPTX doesn't require phantom intermediate levels — just start
        # the list at whatever indent the content actually uses.
        ctx.builder.begin_list(ordered=ordered)
        stack.append(_ListState(level=level, ordered=ordered))

    # Close previous item at same level
    if stack and stack[-1].level == level and stack[-1].item_open:
        ctx.builder.end()  # Close previous list item
        stack[-1].item_open = False

    # Open new list item
    ctx.builder.begin_list_item()
    stack[-1].item_open = True

    # Add paragraph content
    inlines = _runs_to_inlines(para)
    if inlines:
        ctx.builder.paragraph(*inlines)


def _flush_lists(ctx: ParseContext) -> None:
    """Close all open lists in the stack."""
    while ctx.list_stack:
        state = ctx.list_stack.pop()
        if state.item_open:
            ctx.builder.end()  # Close list item
        ctx.builder.end()  # Close list


def _runs_to_inlines(para: _Paragraph) -> list[Inline]:
    """Convert paragraph runs to inline content with formatting."""
    inlines: list[Inline] = []

    for run in para.runs:
        text = run.text
        if not text:
            continue

        inline: Inline = Text(value=text)

        # Check for hyperlink via python-pptx API
        url = ""
        with contextlib.suppress(Exception):
            if run.hyperlink and run.hyperlink.address:
                url = run.hyperlink.address

        # Apply formatting from run properties
        bold = run.font.bold is True
        italic = run.font.italic is True

        if url:
            link_child: Inline = Text(value=text)
            if bold and italic:
                link_child = Strong(children=(Emphasis(children=(link_child,)),))
            elif bold:
                link_child = Strong(children=(link_child,))
            elif italic:
                link_child = Emphasis(children=(link_child,))
            inline = Link(url=url, children=(link_child,))
        elif bold and italic:
            inline = Strong(children=(Emphasis(children=(inline,)),))
        elif bold:
            inline = Strong(children=(inline,))
        elif italic:
            inline = Emphasis(children=(inline,))

        inlines.append(inline)

    return inlines


def _process_table(table: PptxTable, ctx: ParseContext) -> None:
    """Process a PPTX table into a Table block."""
    from kaos_content.model.blocks import Paragraph

    if not table.rows:
        return

    num_cols = len(table.columns)
    rows: list[Row] = []

    for _row_idx, row in enumerate(table.rows):
        cells: list[Cell] = []
        for col_idx in range(num_cols):
            try:
                cell = row.cells[col_idx]
            except IndexError:
                cells.append(Cell(content=()))
                continue

            # Check for merge markers on the raw XML
            tc_el = cell._tc
            h_merge = tc_el.get("hMerge")
            v_merge = tc_el.get("vMerge")

            # Skip continuation cells
            if h_merge == "1" or v_merge == "1":
                continue

            # Get span info
            grid_span = int(tc_el.get("gridSpan", "1"))
            row_span = int(tc_el.get("rowSpan", "1"))

            # Extract cell text
            text = cell.text.strip()
            content = (Paragraph(children=(Text(value=text),)),) if text else ()
            cells.append(Cell(content=content, col_span=grid_span, row_span=row_span))

        if cells:
            rows.append(Row(cells=tuple(cells)))

    if not rows:
        return

    # First row as header, rest as body
    head = TableSection(rows=(rows[0],))
    body = TableSection(rows=tuple(rows[1:])) if len(rows) > 1 else None
    bodies = (body,) if body else ()

    ctx.builder.add_block(Table(head=head, bodies=bodies))


def _process_chart(chart: Chart, ctx: ParseContext) -> None:
    """Extract chart data as a linearized Table."""
    # Title
    title_text = None
    if chart.has_title:
        with contextlib.suppress(Exception):
            title_text = chart.chart_title.text_frame.text.strip()

    # Extract categories and series data
    categories: list[str] = []
    series_names: list[str] = []
    series_values: list[list[str]] = []

    try:
        for plot in chart.plots:
            cats = [str(c) for c in plot.categories]
            if cats and not categories:
                categories = cats

            for ser in plot.series:
                name = ""
                with contextlib.suppress(Exception):
                    name = str(ser.name) if ser.name else ""
                series_names.append(name)

                vals: list[str] = []
                try:
                    for v in ser.values:
                        vals.append(str(v) if v is not None else "")
                except Exception as exc:
                    logger.debug("Failed to extract chart series values: %s", exc)
                series_values.append(vals)
    except Exception as exc:
        logger.debug("Failed to extract chart data: %s", exc)

    if not categories and not series_values:
        # No data extracted — add title as paragraph if present
        if title_text:
            ctx.builder.paragraph(Text(value=f"[Chart: {title_text}]"))
        return

    # Build table: Category column + series columns
    if title_text:
        ctx.builder.heading(3, title_text)

    headers = ["Category"] + [n or f"Series {i + 1}" for i, n in enumerate(series_names)]
    table_rows: list[list[str]] = []
    for i, cat in enumerate(categories):
        row = [cat]
        for sv in series_values:
            row.append(sv[i] if i < len(sv) else "")
        table_rows.append(row)

    if headers and table_rows:
        ctx.builder.table(headers, table_rows)


def _is_smartart(shape: BaseShape) -> bool:
    """Check if a shape is a SmartArt diagram by inspecting raw XML."""
    el = shape._element
    graphic = el.find(f".//{A_GRAPHIC}")
    if graphic is None:
        return False
    gd = graphic.find(A_GRAPHIC_DATA)
    if gd is None:
        return False
    return gd.get("uri") == GD_DIAGRAM


def _process_smartart(shape: BaseShape, slide: Slide, ctx: ParseContext) -> None:
    """Extract SmartArt text via OPC fallback."""
    el = shape._element
    graphic = el.find(f".//{A_GRAPHIC}")
    if graphic is None:
        return
    gd = graphic.find(A_GRAPHIC_DATA)
    if gd is None:
        return

    # Find the slide part path for resolving relationships
    slide_part = slide.part.partname
    # slide_part is like '/ppt/slides/slide1.xml'
    slide_part_str = str(slide_part).lstrip("/")
    slide_dir = "/".join(slide_part_str.split("/")[:-1])

    # Get slide relationships from OPC layer
    rels_path = f"{slide_dir}/_rels/{slide_part_str.split('/')[-1]}.rels"
    slide_rels = ctx.pkg.relationships(rels_path)

    texts = extract_smartart_texts(gd, ctx.pkg, slide_rels, slide_dir)
    if texts:
        ctx.builder.begin_div(classes="smartart")
        for text in texts:
            ctx.builder.paragraph(Text(value=text))
        ctx.builder.end()


def _process_picture(shape: BaseShape, ctx: ParseContext) -> None:
    """Extract image reference from a picture shape."""
    name = shape.name or "image"
    alt = None

    # Try to get alt text from cNvPr descr attribute
    with contextlib.suppress(Exception):
        el = shape._element
        from kaos_office.ooxml.namespace import P_CNV_PR, P_NV_PIC_PR, P_NV_SP_PR

        for nvpr_tag in (P_NV_PIC_PR, P_NV_SP_PR):
            nvpr = el.find(nvpr_tag)
            if nvpr is not None:
                cnvpr = nvpr.find(P_CNV_PR)
                if cnvpr is not None:
                    alt = cnvpr.get("descr")
                    if not alt:
                        alt = cnvpr.get("name")
                    break

    # Determine content type / extension
    ext = "png"
    with contextlib.suppress(Exception):
        content_type = shape.image.content_type  # ty: ignore[unresolved-attribute]
        if content_type:
            ext = content_type.split("/")[-1]
            if ext == "jpeg":
                ext = "jpg"

    src = f"pptx://{name}.{ext}"
    ctx.builder.image(src=src, alt=alt or name)


def _process_notes(slide: Slide, ctx: ParseContext) -> None:
    """Extract speaker notes if present."""
    if not slide.has_notes_slide:
        return

    try:
        notes_tf = slide.notes_slide.notes_text_frame
        if notes_tf is None:
            return
        text = notes_tf.text.strip()
        if text:
            ctx.builder.begin_div(classes="speaker-notes")
            ctx.builder.paragraph(Text(value=text))
            ctx.builder.end()
    except Exception as exc:
        logger.debug("Failed to extract speaker notes: %s", exc)


# --- Public helpers for MCP tools ---


def get_slide_count(path: str | Path) -> int:
    """Get the number of slides in a PPTX file."""
    from pptx import Presentation

    prs = Presentation(str(path))
    return len(prs.slides)


def get_slide_notes(path: str | Path, slide_number: int) -> str | None:
    """Extract speaker notes from a specific slide (1-based numbering).

    Returns the notes text, or None if the slide has no notes.

    Raises:
        ValueError: If slide_number is out of range.
    """
    from pptx import Presentation

    prs = Presentation(str(path))
    if slide_number < 1 or slide_number > len(prs.slides):
        msg = f"Slide {slide_number} out of range (1-{len(prs.slides)})"
        raise ValueError(msg)

    slide = prs.slides[slide_number - 1]
    if not slide.has_notes_slide:
        return None

    try:
        notes_tf = slide.notes_slide.notes_text_frame
        if notes_tf is None:
            return None
        text = notes_tf.text.strip()
        return text if text else None
    except Exception as exc:
        logger.debug("Failed to extract slide notes: %s", exc)
        return None


def get_slide_text(path: str | Path, slide_number: int) -> str:
    """Extract text from a specific slide (1-based numbering).

    Returns plain text of all shapes on the slide.
    """
    from pptx import Presentation

    prs = Presentation(str(path))
    if slide_number < 1 or slide_number > len(prs.slides):
        msg = f"Slide {slide_number} out of range (1-{len(prs.slides)})"
        raise ValueError(msg)

    slide = prs.slides[slide_number - 1]
    parts: list[str] = []

    for shape in _sort_shapes(list(slide.shapes)):
        ph_type = _get_placeholder_type_name(shape)
        if ph_type in _SKIP_PLACEHOLDER_TYPES:
            continue
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()  # ty: ignore[unresolved-attribute]
            if text:
                parts.append(text)
        elif shape.has_table:
            for row in shape.table.rows:  # ty: ignore[unresolved-attribute]
                row_texts = []
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        row_texts.append(ct)
                if row_texts:
                    parts.append(" | ".join(row_texts))

    return "\n\n".join(parts)


def list_slides(path: str | Path) -> list[dict[str, str | int]]:
    """List slides with their titles and shape counts.

    Returns list of dicts with slide_number, title, shape_count, has_notes.
    """
    from pptx import Presentation

    prs = Presentation(str(path))
    result: list[dict[str, str | int]] = []

    for i, slide in enumerate(prs.slides, 1):
        title = ""
        shape_count = len(slide.shapes)

        for shape in slide.shapes:
            ph_type = _get_placeholder_type_name(shape)
            if ph_type in _TITLE_PLACEHOLDER_TYPES and shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    title = text
                    break

        has_notes = False
        if slide.has_notes_slide:
            with contextlib.suppress(Exception):
                notes_text = slide.notes_slide.notes_text_frame.text.strip()
                has_notes = bool(notes_text)

        result.append(
            {
                "slide_number": i,
                "title": title,
                "shape_count": shape_count,
                "has_notes": has_notes,
            }
        )

    return result
