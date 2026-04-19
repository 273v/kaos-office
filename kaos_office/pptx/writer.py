"""PPTX writer — ContentDocument to PresentationML.

Serializes a kaos-content ContentDocument AST to a PPTX file using
python-pptx. Handles slide boundary detection, layout selection,
and block-to-shape mapping.

Usage::

    from kaos_office.pptx.writer import write_pptx, write_pptx_bytes

    write_pptx(content_doc, "output.pptx")
    pptx_bytes = write_pptx_bytes(content_doc)
"""

from __future__ import annotations

import contextlib
from io import BytesIO
from pathlib import Path
from typing import Any

from kaos_core.logging import get_logger
from pptx import Presentation
from pptx.presentation import Presentation as PptxPresentation
from pptx.util import Inches, Pt

logger = get_logger(__name__)


def write_pptx(
    doc: Any,
    path: str | Path,
    *,
    template: str | Path | None = None,
) -> Path:
    """Write a ContentDocument to a PPTX file.

    Args:
        doc: A ``ContentDocument`` from kaos-content.
        path: Output file path.
        template: Optional .pptx template for branded output.

    Returns:
        The output path.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    data = write_pptx_bytes(doc, template=template)
    path.write_bytes(data)

    logger.info(
        "pptx.writer: wrote %s, blocks=%d, size=%d, path=%s",
        doc.metadata.title or "untitled",
        len(doc.body),
        len(data),
        path,
    )
    return path


def write_pptx_bytes(
    doc: Any,
    *,
    template: str | Path | None = None,
) -> bytes:
    """Write a ContentDocument to PPTX bytes (in-memory).

    Args:
        doc: A ``ContentDocument`` from kaos-content.
        template: Optional .pptx template for branded output.

    Returns:
        PPTX file as bytes.
    """
    prs = Presentation(str(template)) if template else Presentation()

    # Set metadata
    if doc.metadata.title:
        prs.core_properties.title = doc.metadata.title

    # Detect slide boundaries and build slides
    slide_groups = _segment_into_slides(doc.body)

    for slide_blocks in slide_groups:
        _add_slide(prs, slide_blocks)

    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Slide segmentation
# ---------------------------------------------------------------------------


def _segment_into_slides(body: tuple) -> list[list]:
    """Split document body into groups of blocks, one group per slide.

    If the body contains Div(classes="slide") blocks, use those as boundaries.
    Otherwise, auto-segment: each Heading(depth=1) starts a new slide,
    and Table blocks get their own slide.
    """
    from kaos_content.model.blocks import Heading, Table

    # Check for explicit slide Divs
    try:
        from kaos_content.model.blocks import Div

        explicit_divs = [
            b
            for b in body
            if isinstance(b, Div)
            and "slide" in (getattr(getattr(b, "attr", None), "classes", ()) or ())
        ]
    except ImportError:
        explicit_divs = []

    if explicit_divs:
        # Use explicit slide boundaries
        slides = []
        for div in explicit_divs:
            children = getattr(div, "children", ())
            if children:
                slides.append(list(children))
        return slides if slides else [[]]

    # Auto-segment
    if not body:
        return [[]]

    slides: list[list] = []
    current: list = []

    for block in body:
        if isinstance(block, Heading) and block.depth == 1:
            # H1 starts a new slide
            if current:
                slides.append(current)
            current = [block]
        elif isinstance(block, Table) and current:
            # Tables get their own slide if there's already content
            slides.append(current)
            current = [block]
        else:
            current.append(block)

    if current:
        slides.append(current)

    return slides if slides else [[]]


# ---------------------------------------------------------------------------
# Slide building
# ---------------------------------------------------------------------------

_TITLE_CONTENT_LAYOUT = 1  # "Title and Content"
_TITLE_SLIDE_LAYOUT = 0  # "Title Slide"
_BLANK_LAYOUT = 6  # "Blank" (index 5 is "Title Only")


def _add_slide(prs: PptxPresentation, blocks: list[Any]) -> None:
    """Add a single slide to the presentation from a group of blocks."""
    from kaos_content.model.blocks import Heading, Table

    if not blocks:
        return

    # Determine layout
    first = blocks[0]
    has_title = isinstance(first, Heading) and first.depth == 1
    has_subtitle = len(blocks) >= 2 and isinstance(blocks[1], Heading) and blocks[1].depth == 2
    has_table = any(isinstance(b, Table) for b in blocks)

    if has_title and has_subtitle and len(blocks) == 2:
        layout_idx = _TITLE_SLIDE_LAYOUT
    elif has_title:
        layout_idx = _TITLE_CONTENT_LAYOUT
    elif has_table:
        layout_idx = _BLANK_LAYOUT
    else:
        layout_idx = _BLANK_LAYOUT

    # Clamp to available layouts
    if layout_idx >= len(prs.slide_layouts):
        layout_idx = min(len(prs.slide_layouts) - 1, _BLANK_LAYOUT)

    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

    # Process blocks
    body_blocks = []
    speaker_notes = None

    for block in blocks:
        # Check for speaker notes Div (classes live on block.attr, not on the block itself)
        if type(block).__name__ == "Div":
            attr = getattr(block, "attr", None)
            classes = getattr(attr, "classes", ()) if attr is not None else ()
            if "speaker-notes" in (classes or ()):
                speaker_notes = block
                continue

        if isinstance(block, Heading) and block.depth == 1 and has_title:
            # Set title placeholder
            if slide.shapes.title is not None:
                slide.shapes.title.text = _extract_text(block)
            has_title = False  # only first H1
        elif isinstance(block, Heading) and block.depth == 2 and has_subtitle:
            # Set subtitle placeholder (index 1)
            with contextlib.suppress(KeyError, IndexError):
                slide.placeholders[1].text = _extract_text(block)
            has_subtitle = False
        else:
            body_blocks.append(block)

    # Add body content. If any body block is non-textual (Figure, Table),
    # we must route through _add_body_textbox so figures become picture
    # shapes and tables become table shapes — text placeholders cannot
    # host these shape types.
    if body_blocks:
        from kaos_content.model.blocks import Figure as _Figure
        from kaos_content.model.blocks import Table as _Table

        has_shapes = any(isinstance(b, (_Figure, _Table)) for b in body_blocks)

        try:
            ph1 = slide.placeholders[1] if layout_idx == _TITLE_CONTENT_LAYOUT else None
        except (KeyError, IndexError):
            ph1 = None

        if ph1 is not None and not has_shapes:
            _fill_body_placeholder(ph1, body_blocks)
        else:
            _add_body_textbox(slide, body_blocks)

    # Add speaker notes — preserve inline formatting by recursing through the
    # Div's block children rather than just flattening to text.
    if speaker_notes is not None:
        notes_children = list(getattr(speaker_notes, "children", ()) or ())
        if notes_children:
            notes_slide = slide.notes_slide
            tf = notes_slide.notes_text_frame
            tf.clear()
            first = True
            for nb in notes_children:
                _add_block_to_textframe(tf, nb, first=first)
                first = False


def _fill_body_placeholder(placeholder: Any, blocks: list) -> None:
    """Fill a body placeholder with blocks."""
    tf = placeholder.text_frame
    tf.clear()
    first = True
    for block in blocks:
        _add_block_to_textframe(tf, block, first=first)
        first = False


def _add_body_textbox(slide: Any, blocks: list) -> None:
    """Add a textbox to the slide for body content."""
    from kaos_content.model.blocks import Figure, Table

    # Figures (containing Images) and Tables each get their own shape;
    # everything else flows into a single text frame.
    text_blocks: list[Any] = []
    table_blocks: list[Any] = []
    figure_blocks: list[Any] = []
    for b in blocks:
        if isinstance(b, Table):
            table_blocks.append(b)
        elif isinstance(b, Figure):
            figure_blocks.append(b)
        else:
            text_blocks.append(b)

    if text_blocks:
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9.0)
        height = Inches(5.0)
        txbox = slide.shapes.add_textbox(left, top, width, height)
        tf = txbox.text_frame
        tf.word_wrap = True
        first = True
        for block in text_blocks:
            _add_block_to_textframe(tf, block, first=first)
            first = False

    for table_block in table_blocks:
        _add_table_shape(slide, table_block)

    for figure_block in figure_blocks:
        _add_figure_shape(slide, figure_block)


# ---------------------------------------------------------------------------
# Block → text frame content
# ---------------------------------------------------------------------------


def _add_block_to_textframe(tf: Any, block: Any, *, first: bool = False) -> None:
    """Add a block to a text frame."""
    from kaos_content.model.blocks import (
        BlockQuote,
        BulletList,
        CodeBlock,
        Heading,
        OrderedList,
        Paragraph,
    )

    if isinstance(block, Paragraph):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        _fill_paragraph(p, block)
    elif isinstance(block, Heading):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        _fill_paragraph(p, block)
        p.font.bold = True
        # Scale heading size
        sizes = {1: 28, 2: 24, 3: 20, 4: 18, 5: 16, 6: 14}
        depth = getattr(block, "depth", 1)
        p.font.size = Pt(sizes.get(depth, 18))
    elif isinstance(block, (BulletList, OrderedList)):
        _add_list_to_textframe(
            tf, block, level=0, first=first, ordered=isinstance(block, OrderedList)
        )
    elif isinstance(block, CodeBlock):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        text = getattr(block, "value", "") or ""
        run = p.add_run()
        run.text = text
        run.font.name = "Consolas"
        run.font.size = Pt(10)
    elif isinstance(block, BlockQuote):
        for i, child in enumerate(getattr(block, "children", ())):
            _add_block_to_textframe(tf, child, first=first and i == 0)
    else:
        # Fallback: extract text
        text = _extract_text(block)
        if text.strip():
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.text = text


def _add_list_to_textframe(
    tf: Any,
    lst: Any,
    *,
    level: int = 0,
    first: bool = False,
    ordered: bool = False,
) -> None:
    """Add a list (bullet or ordered) to a text frame."""
    from kaos_content.model.blocks import BulletList, ListItem, OrderedList, Paragraph
    from pptx.oxml.ns import qn

    for item_idx, item in enumerate(getattr(lst, "children", ())):
        if not isinstance(item, ListItem):
            continue

        for child_idx, child in enumerate(item.children):
            if isinstance(child, (BulletList, OrderedList)):
                _add_list_to_textframe(
                    tf,
                    child,
                    level=level + 1,
                    first=first and item_idx == 0 and child_idx == 0,
                    ordered=isinstance(child, OrderedList),
                )
            else:
                p = (
                    tf.paragraphs[0]
                    if (first and item_idx == 0 and child_idx == 0)
                    else tf.add_paragraph()
                )
                p.level = level

                if isinstance(child, Paragraph):
                    _fill_paragraph(p, child)
                else:
                    p.text = _extract_text(child)

                # Set bullet/numbering via XML
                pPr = p._p.get_or_add_pPr()
                if ordered:
                    bu = pPr.makeelement(qn("a:buAutoNum"), {})
                    bu.set("type", "arabicPeriod")
                    pPr.append(bu)
                else:
                    bu = pPr.makeelement(qn("a:buChar"), {})
                    bu.set("char", "\u2022")
                    pPr.append(bu)


# ---------------------------------------------------------------------------
# Table → shape
# ---------------------------------------------------------------------------


def _add_table_shape(slide: Any, table_block: Any) -> None:
    """Add a Table block as a table shape on the slide."""
    head = getattr(table_block, "head", None)
    bodies = getattr(table_block, "bodies", ())

    # Count rows and cols
    rows_data = []
    if head:
        for row in head.rows:
            rows_data.append((row, True))
    for body in bodies:
        for row in body.rows:
            rows_data.append((row, False))

    if not rows_data:
        return

    n_cols = max(len(r.cells) for r, _ in rows_data) if rows_data else 0
    if n_cols == 0:
        return

    n_rows = len(rows_data)

    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9.0)
    height = Inches(0.4 * n_rows)

    tbl_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    tbl = tbl_shape.table

    # First pass: set text and record merge regions
    merge_regions: list[tuple[int, int, int, int]] = []  # (row, col, col_span, row_span)
    for row_idx, (row, _is_header) in enumerate(rows_data):
        for col_idx, cell in enumerate(row.cells):
            if col_idx >= n_cols:
                continue
            cell_text = _extract_text(cell)
            tbl.cell(row_idx, col_idx).text = cell_text

            col_span = int(getattr(cell, "col_span", 1) or 1)
            row_span = int(getattr(cell, "row_span", 1) or 1)
            if col_span > 1 or row_span > 1:
                merge_regions.append((row_idx, col_idx, col_span, row_span))

    # Second pass: apply merges via raw XML (python-pptx has no merge API)
    for row_idx, col_idx, col_span, row_span in merge_regions:
        _apply_cell_merge(tbl, row_idx, col_idx, col_span, row_span, n_rows, n_cols)


def _add_figure_shape(slide: Any, figure: Any) -> None:
    """Add a Figure block (containing Images) as a picture shape on the slide.

    The first Image inline in the figure is embedded via
    ``slide.shapes.add_picture``. Subsequent images or non-image children
    fall through to a text box (alt text fallback).

    Image data resolution:
    - ``src`` starting with ``pptx://`` is an artifact URI from the reader;
      without an artifact store hook, we can't resolve it, so we fall back
      to emitting the alt text only.
    - ``src`` that is a real local file path is embedded directly.
    - ``src`` bytes (passed via Attr.kv ``image-data`` as a data URL or path)
      could be added in future work.
    """
    from pathlib import Path as _Path

    images = _find_images(figure)
    if not images:
        # No images — just emit figure content as text fallback
        text = _extract_text(figure)
        if text.strip():
            _add_text_only_box(slide, text)
        return

    image = images[0]
    src = getattr(image, "src", "") or ""
    alt = getattr(image, "alt", None) or getattr(image, "title", None) or ""

    # Try to resolve the image source
    image_file: Any = None
    if src and not src.startswith(("pptx://", "http://", "https://")):
        # Assume local file path
        path = _Path(src)
        if path.exists():
            image_file = str(path)

    if image_file is None:
        # Can't resolve — emit alt text as a text box so the slide isn't empty
        if alt:
            _add_text_only_box(slide, alt)
        return

    # Default positioning: centered-ish on the slide
    left = Inches(1.0)
    top = Inches(1.5)

    # Respect explicit Image.width / Image.height (in points per kaos-content
    # convention). Fall back to a sensible max width when unset so the slide
    # doesn't render full-bleed.
    from pptx.util import Pt

    width_pt = getattr(image, "width", None)
    height_pt = getattr(image, "height", None)
    size_kwargs: dict[str, Any] = {}
    if width_pt is not None:
        size_kwargs["width"] = Pt(float(width_pt))
    if height_pt is not None:
        size_kwargs["height"] = Pt(float(height_pt))
    if not size_kwargs:
        size_kwargs["width"] = Inches(8.0)

    try:
        pic = slide.shapes.add_picture(image_file, left, top, **size_kwargs)
    except (OSError, ValueError):
        # Corrupted or unreadable image — fall back to alt text
        if alt:
            _add_text_only_box(slide, alt)
        return

    # Set alt text via the cNvPr descr attribute (matches reader extraction)
    if alt:
        with contextlib.suppress(Exception):
            el = pic._element
            nvpr = el.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr")
            if nvpr is not None:
                cnvpr = nvpr.find(
                    "{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr"
                )
                if cnvpr is not None:
                    cnvpr.set("descr", alt)


def _find_images(figure: Any) -> list[Any]:
    """Walk a Figure's subtree and return every Image inline."""
    from kaos_content.model.inlines import Image

    out: list[Any] = []

    def _walk(node: Any) -> None:
        if isinstance(node, Image):
            out.append(node)
            return
        for c in getattr(node, "children", ()) or ():
            _walk(c)

    _walk(figure)
    return out


def _add_text_only_box(slide: Any, text: str) -> None:
    """Helper: drop a simple text box on the slide with the given content."""
    left = Inches(1.0)
    top = Inches(1.5)
    width = Inches(8.0)
    height = Inches(1.0)
    txbox = slide.shapes.add_textbox(left, top, width, height)
    tf = txbox.text_frame
    tf.word_wrap = True
    tf.text = text


def _apply_cell_merge(
    tbl: Any,
    row: int,
    col: int,
    col_span: int,
    row_span: int,
    n_rows: int,
    n_cols: int,
) -> None:
    """Set gridSpan/rowSpan on the origin cell and hMerge/vMerge on continuations.

    Uses raw lxml attribute manipulation via the private ``_tc`` element
    since python-pptx has no high-level merge API.
    """
    origin_tc = tbl.cell(row, col)._tc
    if col_span > 1:
        origin_tc.set("gridSpan", str(col_span))
    if row_span > 1:
        origin_tc.set("rowSpan", str(row_span))

    # Mark continuation cells in the same row (horizontal merge)
    for c in range(col + 1, min(col + col_span, n_cols)):
        tc = tbl.cell(row, c)._tc
        tc.set("hMerge", "1")

    # Mark continuation cells in subsequent rows (vertical merge)
    for r in range(row + 1, min(row + row_span, n_rows)):
        for c in range(col, min(col + col_span, n_cols)):
            tc = tbl.cell(r, c)._tc
            tc.set("vMerge", "1")
            # Cells diagonally merged need hMerge too
            if c > col:
                tc.set("hMerge", "1")


# ---------------------------------------------------------------------------
# Inline text extraction
# ---------------------------------------------------------------------------


def _fill_paragraph(p: Any, block: Any) -> None:
    """Fill a python-pptx paragraph with inline content from a block."""
    children = getattr(block, "children", ())
    if not children:
        text = _extract_text(block)
        if text:
            p.text = text
        return

    for i, inline in enumerate(children):
        _add_inline_to_paragraph(p, inline, first_run=(i == 0))


def _add_inline_to_paragraph(p: Any, inline: Any, *, first_run: bool = False) -> None:
    """Add an inline element as a run in a paragraph."""
    from kaos_content.model.inlines import Code, Emphasis, LineBreak, Link, SoftBreak, Strong, Text

    if isinstance(inline, Text):
        run = p.runs[0] if (first_run and p.runs) else p.add_run()
        run.text = inline.value
    elif isinstance(inline, Strong):
        for child in inline.children:
            run = p.add_run()
            run.text = _extract_text(child)
            run.font.bold = True
    elif isinstance(inline, Emphasis):
        for child in inline.children:
            run = p.add_run()
            run.text = _extract_text(child)
            run.font.italic = True
    elif isinstance(inline, Code):
        run = p.add_run()
        run.text = inline.value
        run.font.name = "Consolas"
        run.font.size = Pt(10)
    elif isinstance(inline, Link):
        run = p.add_run()
        run.text = _extract_text(inline)
        run.hyperlink.address = inline.url
    elif isinstance(inline, (LineBreak, SoftBreak)):
        run = p.add_run()
        run.text = "\n"
    else:
        text = _extract_text(inline)
        if text:
            run = p.add_run()
            run.text = text


def _extract_text(node: Any) -> str:
    """Extract plain text from any AST node."""
    # Leaf with value
    value = getattr(node, "value", None)
    if isinstance(value, str):
        return value

    # Node with children
    children = getattr(node, "children", None)
    if children:
        return "".join(_extract_text(c) for c in children)

    # Cell content
    content = getattr(node, "content", None)
    if content:
        return "".join(_extract_text(c) for c in content)

    # Text attribute
    text = getattr(node, "text", None)
    if isinstance(text, str):
        return text

    return ""
