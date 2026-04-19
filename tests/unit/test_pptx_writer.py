"""Unit tests for PPTX writer.

Round-trip tests: ContentDocument → write_pptx → parse_pptx → verify.
Synthetic tests: build from scratch → write → verify structure.
Slide segmentation: verify auto-detection of slide boundaries.
"""

from __future__ import annotations

import time
import zipfile
from io import BytesIO
from pathlib import Path

from kaos_content.model.blocks import (
    BulletList,
    CodeBlock,
    Heading,
    ListItem,
    OrderedList,
    Paragraph,
    Table,
)
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Code, Emphasis, Strong, Text
from kaos_content.model.table import Cell, Row, TableSection
from kaos_content.serializers.text import serialize_text

from kaos_office.pptx.reader import parse_pptx
from kaos_office.pptx.writer import write_pptx, write_pptx_bytes

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "pptx"


def _zip_parts(pptx_bytes: bytes) -> list[str]:
    return sorted(zipfile.ZipFile(BytesIO(pptx_bytes)).namelist())


# ---------------------------------------------------------------------------
# OPC structure
# ---------------------------------------------------------------------------


class TestOPCStructure:
    def test_valid_pptx_zip(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Test"),
            body=(Paragraph(children=(Text(value="Hello"),)),),
        )
        parts = _zip_parts(write_pptx_bytes(doc))
        assert "[Content_Types].xml" in parts
        assert "_rels/.rels" in parts
        # Should have at least one slide
        slide_parts = [p for p in parts if "slide" in p.lower() and p.endswith(".xml")]
        assert len(slide_parts) >= 1

    def test_empty_body(self) -> None:
        doc = ContentDocument(metadata=DocumentMetadata(title="Empty"), body=())
        data = write_pptx_bytes(doc)
        assert len(data) > 0

    def test_metadata_title(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="My Presentation"),
            body=(Paragraph(children=(Text(value="content"),)),),
        )
        data = write_pptx_bytes(doc)
        zf = zipfile.ZipFile(BytesIO(data))
        # Core properties should contain the title
        if "docProps/core.xml" in zf.namelist():
            core = zf.read("docProps/core.xml").decode()
            assert "My Presentation" in core


# ---------------------------------------------------------------------------
# Slide segmentation
# ---------------------------------------------------------------------------


class TestSlideSegmentation:
    def test_h1_starts_new_slide(self, tmp_path: Path) -> None:
        """Each H1 should start a new slide."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Slide 1"),)),
                Paragraph(children=(Text(value="Content 1"),)),
                Heading(depth=1, children=(Text(value="Slide 2"),)),
                Paragraph(children=(Text(value="Content 2"),)),
                Heading(depth=1, children=(Text(value="Slide 3"),)),
            ),
        )
        out = tmp_path / "segments.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Slide 1" in rt
        assert "Slide 2" in rt
        assert "Slide 3" in rt

    def test_title_subtitle_slide(self, tmp_path: Path) -> None:
        """H1 followed by H2 with nothing else → title slide."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Big Title"),)),
                Heading(depth=2, children=(Text(value="Subtitle Here"),)),
            ),
        )
        out = tmp_path / "title.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Big Title" in rt
        assert "Subtitle Here" in rt

    def test_single_paragraph_one_slide(self, tmp_path: Path) -> None:
        """A single paragraph without H1 → one slide."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="Just text"),)),),
        )
        out = tmp_path / "single.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Just text" in rt


# ---------------------------------------------------------------------------
# Block type serialization
# ---------------------------------------------------------------------------


class TestBlockSerialization:
    def test_heading(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Heading(depth=1, children=(Text(value="Title Text"),)),),
        )
        out = tmp_path / "heading.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Title Text" in rt

    def test_bullet_list(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Lists"),)),
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Bullet A"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Bullet B"),)),)),
                    )
                ),
            ),
        )
        out = tmp_path / "bullets.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Bullet A" in rt
        assert "Bullet B" in rt

    def test_ordered_list(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Steps"),)),
                OrderedList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Step 1"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Step 2"),)),)),
                    )
                ),
            ),
        )
        out = tmp_path / "ordered.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Step 1" in rt
        assert "Step 2" in rt

    def test_table(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Table(
                    head=TableSection(
                        rows=(
                            Row(
                                cells=(
                                    Cell(content=(Paragraph(children=(Text(value="Col1"),)),)),
                                    Cell(content=(Paragraph(children=(Text(value="Col2"),)),)),
                                )
                            ),
                        )
                    ),
                    bodies=(
                        TableSection(
                            rows=(
                                Row(
                                    cells=(
                                        Cell(content=(Paragraph(children=(Text(value="A"),)),)),
                                        Cell(content=(Paragraph(children=(Text(value="B"),)),)),
                                    )
                                ),
                            )
                        ),
                    ),
                ),
            ),
        )
        out = tmp_path / "table.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "Col1" in rt
        assert "A" in rt

    def test_code_block(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Code"),)),
                CodeBlock(value="print('hello')", language="python"),
            ),
        )
        out = tmp_path / "code.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "hello" in rt


# ---------------------------------------------------------------------------
# Inline formatting
# ---------------------------------------------------------------------------


class TestInlineFormatting:
    def test_bold(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Fmt"),)),
                Paragraph(
                    children=(
                        Text(value="normal "),
                        Strong(children=(Text(value="bold"),)),
                    )
                ),
            ),
        )
        out = tmp_path / "bold.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "bold" in rt

    def test_italic(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Fmt"),)),
                Paragraph(
                    children=(
                        Text(value="normal "),
                        Emphasis(children=(Text(value="italic"),)),
                    )
                ),
            ),
        )
        out = tmp_path / "italic.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "italic" in rt

    def test_inline_code(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Fmt"),)),
                Paragraph(children=(Code(value="x = 1"),)),
            ),
        )
        out = tmp_path / "icode.pptx"
        write_pptx(doc, out)
        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)
        assert "x = 1" in rt


# ---------------------------------------------------------------------------
# Comprehensive creation test
# ---------------------------------------------------------------------------


class TestComprehensiveDocument:
    def test_all_block_types(self, tmp_path: Path) -> None:
        """Multi-slide doc with all supported block types."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Full Test"),
            body=(
                Heading(depth=1, children=(Text(value="Slide One"),)),
                Heading(depth=2, children=(Text(value="The subtitle"),)),
                Heading(depth=1, children=(Text(value="Slide Two"),)),
                Paragraph(
                    children=(
                        Text(value="Mixed: "),
                        Strong(children=(Text(value="bold"),)),
                        Text(value=", "),
                        Emphasis(children=(Text(value="italic"),)),
                        Text(value=", "),
                        Code(value="code"),
                    )
                ),
                BulletList(
                    children=(ListItem(children=(Paragraph(children=(Text(value="Bullet"),)),)),)
                ),
                OrderedList(
                    children=(ListItem(children=(Paragraph(children=(Text(value="Num"),)),)),)
                ),
                Heading(depth=1, children=(Text(value="Slide Three"),)),
                CodeBlock(value="fn main() {}", language="rust"),
                Table(
                    head=TableSection(
                        rows=(
                            Row(
                                cells=(
                                    Cell(content=(Paragraph(children=(Text(value="H1"),)),)),
                                    Cell(content=(Paragraph(children=(Text(value="H2"),)),)),
                                )
                            ),
                        )
                    ),
                    bodies=(
                        TableSection(
                            rows=(
                                Row(
                                    cells=(
                                        Cell(content=(Paragraph(children=(Text(value="r1"),)),)),
                                        Cell(content=(Paragraph(children=(Text(value="r2"),)),)),
                                    )
                                ),
                            )
                        ),
                    ),
                ),
            ),
        )
        out = tmp_path / "full.pptx"
        write_pptx(doc, out)
        assert out.exists()

        doc2 = parse_pptx(out)
        rt = serialize_text(doc2)

        assert "Slide One" in rt
        assert "The subtitle" in rt
        assert "bold" in rt
        assert "italic" in rt
        assert "code" in rt
        assert "Bullet" in rt
        assert "Num" in rt
        assert "fn main()" in rt or "main()" in rt
        assert "H1" in rt
        assert "r1" in rt


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------


class TestErrorHandling:
    def test_write_creates_parent_dirs(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="test"),)),),
        )
        out = tmp_path / "a" / "b" / "output.pptx"
        result = write_pptx(doc, out)
        assert result == out
        assert out.exists()


# ---------------------------------------------------------------------------
# Performance
# ---------------------------------------------------------------------------


class TestPerformance:
    def test_50_slides_under_5s(self, tmp_path: Path) -> None:
        blocks = []
        for i in range(50):
            blocks.append(Heading(depth=1, children=(Text(value=f"Slide {i}"),)))
            blocks.append(
                Paragraph(
                    children=(
                        Text(value=f"Content for slide {i}. "),
                        Strong(children=(Text(value="Important"),)),
                    )
                )
            )
        doc = ContentDocument(metadata=DocumentMetadata(title="Big"), body=tuple(blocks))

        start = time.monotonic()
        data = write_pptx_bytes(doc)
        elapsed = time.monotonic() - start

        assert elapsed < 5.0, f"50-slide write took {elapsed:.2f}s (budget 5s)"
        assert len(data) > 0
