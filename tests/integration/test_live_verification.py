"""Live verification: generate Office files and open them with LibreOffice.

This is the ``integration`` tier — not run by default. Invoke explicitly:

    uv run pytest tests/integration/test_live_verification.py -v

Each test:
1. Builds a synthetic ContentDocument / TabularDocument
2. Writes it via the appropriate writer (write_docx / write_xlsx / write_pptx)
3. Runs LibreOffice in headless mode to convert the output to PDF
4. Asserts that LibreOffice produced a non-empty PDF without errors

If LibreOffice itself opens the file and renders it without errors, the
file is compatible with Office/LibreOffice/Google Docs. This is the
strictest acceptance gate.

Skip condition: ``libreoffice`` command not in PATH.
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import pytest
from kaos_content.model.attr import Attr
from kaos_content.model.blocks import (
    BulletList,
    CodeBlock,
    Div,
    Figure,
    Heading,
    ListItem,
    OrderedList,
    Paragraph,
    Table,
)
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Code, Emphasis, Image, Link, Span, Strong, Text
from kaos_content.model.table import Cell, Row, TableSection
from kaos_content.model.tabular import Column, ColumnType, TabularDocument
from kaos_content.model.tabular import Table as TabTable

from kaos_office.docx.writer import write_docx
from kaos_office.pptx.writer import write_pptx
from kaos_office.xlsx.writer import write_xlsx

LIBREOFFICE = shutil.which("libreoffice") or shutil.which("soffice")
pytestmark = pytest.mark.skipif(LIBREOFFICE is None, reason="LibreOffice not available")


def _convert_to_pdf(input_path: Path, out_dir: Path) -> Path:
    """Run LibreOffice in headless mode to convert ``input_path`` to PDF.

    Returns the path of the produced PDF. Raises if LibreOffice fails
    or produces no output.
    """
    assert LIBREOFFICE is not None
    # --headless --convert-to pdf --outdir <dir> <file>
    result = subprocess.run(
        [
            LIBREOFFICE,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(input_path),
        ],
        capture_output=True,
        timeout=60,
        check=False,
    )
    # LibreOffice emits warnings to stderr — tolerate nonzero exit if a PDF
    # was produced. Fail only if no output appears.
    pdf_path = out_dir / (input_path.stem + ".pdf")
    if not pdf_path.exists() or pdf_path.stat().st_size == 0:
        msg = (
            f"LibreOffice failed to produce PDF for {input_path.name}: "
            f"exit={result.returncode}, stderr={result.stderr.decode(errors='replace')[:500]}"
        )
        raise AssertionError(msg)
    return pdf_path


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------


class TestDocxLiveVerification:
    def test_comprehensive_docx_opens_in_libreoffice(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="DOCX Live Test"),
            body=(
                Heading(depth=1, children=(Text(value="Main Heading"),)),
                Paragraph(
                    children=(
                        Text(value="Formatting test: "),
                        Strong(children=(Text(value="bold"),)),
                        Text(value=", "),
                        Emphasis(children=(Text(value="italic"),)),
                        Text(value=", "),
                        Code(value="inline_code"),
                        Text(value=", "),
                        Link(
                            url="https://example.com",
                            children=(Text(value="a link"),),
                        ),
                        Text(value="."),
                    )
                ),
                Heading(depth=2, children=(Text(value="Lists"),)),
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Bullet A"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Bullet B"),)),)),
                    )
                ),
                OrderedList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Step 1"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Step 2"),)),)),
                    )
                ),
                Heading(depth=2, children=(Text(value="Table"),)),
                Table(
                    head=TableSection(
                        rows=(
                            Row(
                                cells=(
                                    Cell(content=(Paragraph(children=(Text(value="Name"),)),)),
                                    Cell(content=(Paragraph(children=(Text(value="Value"),)),)),
                                )
                            ),
                        )
                    ),
                    bodies=(
                        TableSection(
                            rows=(
                                Row(
                                    cells=(
                                        Cell(content=(Paragraph(children=(Text(value="Alpha"),)),)),
                                        Cell(content=(Paragraph(children=(Text(value="100"),)),)),
                                    )
                                ),
                            )
                        ),
                    ),
                ),
                CodeBlock(value="def hello():\n    return 42", language="python"),
                Paragraph(children=(Text(value="End of document."),)),
            ),
        )
        out = tmp_path / "comprehensive.docx"
        write_docx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 1000  # non-trivial PDF

    def test_docx_with_redlines_opens_in_libreoffice(self, tmp_path: Path) -> None:
        """A document with tracked changes renders without crashing."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Redline Test"),
            body=(
                Paragraph(
                    children=(
                        Text(value="The price is "),
                        Span(
                            attr=Attr(
                                classes=("rev-del",),
                                kv={
                                    "rev:id": "0",
                                    "rev:author": "Alice",
                                    "rev:date": "2026-04-18T10:00:00Z",
                                },
                            ),
                            children=(Text(value="$100"),),
                        ),
                        Span(
                            attr=Attr(
                                classes=("rev-ins",),
                                kv={
                                    "rev:id": "1",
                                    "rev:author": "Alice",
                                    "rev:date": "2026-04-18T10:00:00Z",
                                },
                            ),
                            children=(Text(value="$150"),),
                        ),
                        Text(value="."),
                    )
                ),
            ),
        )
        out = tmp_path / "redline.docx"
        write_docx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 500

    def test_phase4_headers_footers_page_setup_live(self, tmp_path: Path) -> None:
        """DOCX Phase 4 — headers, footers and custom page geometry survive
        LibreOffice. The conversion to PDF + text extraction via kaos-pdf is a
        true cross-tool round-trip: if LibreOffice rejects our sectPr or can't
        resolve the header/footer rels, this test fails.
        """
        from kaos_content.model.metadata import PageSetup
        from kaos_pdf import extract_pdf

        doc = ContentDocument(
            metadata=DocumentMetadata(
                title="Phase 4 Live Test",
                page_setup=PageSetup(
                    page_width_pt=612.0,
                    page_height_pt=792.0,
                    margin_top_pt=72.0,
                    margin_bottom_pt=72.0,
                    margin_left_pt=72.0,
                    margin_right_pt=72.0,
                    header_distance_pt=36.0,
                    footer_distance_pt=36.0,
                ),
            ),
            body=(
                Heading(depth=1, children=(Text(value="Body Heading"),)),
                Paragraph(
                    children=(
                        Text(
                            value=("Body paragraph one. " * 40),
                        ),
                    )
                ),
                Paragraph(
                    children=(
                        Text(
                            value=(
                                "Body paragraph two with a marker BODYSENTINEL for confirmation."
                            )
                        ),
                    )
                ),
            ),
            headers={
                "default": (
                    Paragraph(children=(Text(value="HEADERSENTINEL: live-verified header"),)),
                )
            },
            footers={
                "default": (
                    Paragraph(children=(Text(value="FOOTERSENTINEL: live-verified footer"),)),
                )
            },
        )
        out = tmp_path / "phase4.docx"
        write_docx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 500

        # Cross-tool round-trip: kaos-pdf extracts text from the PDF that
        # LibreOffice rendered. The body + header + footer sentinels must all
        # show up across the extracted pages.
        extracted = extract_pdf(pdf)
        from kaos_content.serializers.text import serialize_text

        rendered_text = serialize_text(extracted)
        assert "BODYSENTINEL" in rendered_text, (
            "body text lost somewhere in write → LibreOffice → extract"
        )
        assert "HEADERSENTINEL" in rendered_text, (
            "header didn't survive LibreOffice rendering — sectPr or header rel may be malformed"
        )
        assert "FOOTERSENTINEL" in rendered_text, "footer didn't survive LibreOffice rendering"


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------


class TestXlsxLiveVerification:
    def test_comprehensive_xlsx_opens_in_libreoffice(self, tmp_path: Path) -> None:
        doc = TabularDocument(
            tables=(
                TabTable(
                    name="Sheet1",
                    columns=(
                        Column(name="Name", column_type=ColumnType.TEXT),
                        Column(name="Count", column_type=ColumnType.INTEGER),
                        Column(name="Price", column_type=ColumnType.FLOAT),
                        Column(name="Active", column_type=ColumnType.BOOLEAN),
                        Column(name="Joined", column_type=ColumnType.DATE),
                    ),
                    rows=(
                        ("Alice", 30, 19.99, True, "2024-01-15"),
                        ("Bob", 25, 9.50, False, "2024-06-01"),
                        ("Charlie", 42, 150.00, True, "2023-11-20"),
                    ),
                ),
            ),
        )
        out = tmp_path / "data.xlsx"
        write_xlsx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 500

    def test_multi_sheet_xlsx_opens_in_libreoffice(self, tmp_path: Path) -> None:
        doc = TabularDocument(
            tables=(
                TabTable(
                    name="First",
                    columns=(Column(name="A", column_type=ColumnType.TEXT),),
                    rows=(("hello",), ("world",)),
                ),
                TabTable(
                    name="Second",
                    columns=(Column(name="N", column_type=ColumnType.INTEGER),),
                    rows=((1,), (2,), (3,)),
                ),
            ),
        )
        out = tmp_path / "multi.xlsx"
        write_xlsx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 500


# ---------------------------------------------------------------------------
# PPTX
# ---------------------------------------------------------------------------


class TestPptxLiveVerification:
    def test_comprehensive_pptx_opens_in_libreoffice(self, tmp_path: Path) -> None:
        # Build a red PNG for image test
        from PIL import Image as PILImage

        img_path = tmp_path / "red.png"
        PILImage.new("RGB", (200, 100), color="red").save(img_path)

        doc = ContentDocument(
            metadata=DocumentMetadata(title="PPTX Live Test"),
            body=(
                # Slide 1: title + subtitle
                Heading(depth=1, children=(Text(value="Deck Title"),)),
                Heading(depth=2, children=(Text(value="A Subtitle"),)),
                # Slide 2: bullets
                Heading(depth=1, children=(Text(value="Key Points"),)),
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Point one"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Point two"),)),)),
                    )
                ),
                # Slide 3: image + speaker notes
                Heading(depth=1, children=(Text(value="Image Slide"),)),
                Figure(
                    children=(Paragraph(children=(Image(src=str(img_path), alt="Red rectangle"),)),)
                ),
                Div(
                    attr=Attr(classes=("speaker-notes",)),
                    children=(
                        Paragraph(children=(Text(value="Speaker note: highlight this image."),)),
                    ),
                ),
                # Slide 4: table with merge
                Heading(depth=1, children=(Text(value="Data Table"),)),
                Table(
                    head=TableSection(
                        rows=(
                            Row(
                                cells=(
                                    Cell(
                                        col_span=2,
                                        content=(
                                            Paragraph(children=(Text(value="Combined Header"),)),
                                        ),
                                    ),
                                    Cell(content=(Paragraph(children=(Text(value="Other"),)),)),
                                )
                            ),
                        )
                    ),
                    bodies=(
                        TableSection(
                            rows=(
                                Row(
                                    cells=(
                                        Cell(content=(Paragraph(children=(Text(value="r1c1"),)),)),
                                        Cell(content=(Paragraph(children=(Text(value="r1c2"),)),)),
                                        Cell(content=(Paragraph(children=(Text(value="r1c3"),)),)),
                                    )
                                ),
                            )
                        ),
                    ),
                ),
            ),
        )
        out = tmp_path / "deck.pptx"
        write_pptx(doc, out)
        pdf = _convert_to_pdf(out, tmp_path)
        assert pdf.stat().st_size > 1000
