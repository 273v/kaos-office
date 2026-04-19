"""Unit tests for DOCX writer.

Round-trip tests: parse fixture → ContentDocument → write_docx → re-parse → verify.
Synthetic tests: build ContentDocument programmatically → write → verify structure.
Modification round-trips: parse → modify → write → re-parse → verify edit persisted.
Performance tests: large documents under time budgets.

Test patterns ported from kelvin-office (kelvin_office/tests/test_round_trip.py,
test_docx_serializer.py, test_style_numbering_roundtrip.py, integration/test_end_to_end.py).
"""

from __future__ import annotations

import time
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from kaos_content.model.blocks import (
    BlockQuote,
    BulletList,
    CodeBlock,
    Heading,
    ListItem,
    OrderedList,
    PageBreak,
    Paragraph,
    Table,
    ThematicBreak,
)
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import (
    Code,
    Emphasis,
    LineBreak,
    Link,
    Strikethrough,
    Strong,
    Text,
)
from kaos_content.model.table import Cell, Row, TableSection
from kaos_content.serializers.text import serialize_text
from lxml import etree

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx, write_docx_bytes

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _zip_parts(docx_bytes: bytes) -> list[str]:
    """Extract sorted part names from DOCX bytes."""
    return sorted(zipfile.ZipFile(BytesIO(docx_bytes)).namelist())


def _doc_xml(docx_bytes: bytes) -> etree._Element:
    """Parse word/document.xml from DOCX bytes."""
    zf = zipfile.ZipFile(BytesIO(docx_bytes))
    return etree.fromstring(zf.read("word/document.xml"))


# ---------------------------------------------------------------------------
# OPC structure tests
# ---------------------------------------------------------------------------


class TestOPCStructure:
    """Verify that the output is a valid OPC package."""

    def test_required_parts(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Test"),
            body=(Paragraph(children=(Text(value="Hello"),)),),
        )
        parts = _zip_parts(write_docx_bytes(doc))
        assert "[Content_Types].xml" in parts
        assert "_rels/.rels" in parts
        assert "word/document.xml" in parts
        assert "word/styles.xml" in parts
        assert "word/_rels/document.xml.rels" in parts
        assert "docProps/core.xml" in parts

    def test_empty_document(self) -> None:
        doc = ContentDocument(metadata=DocumentMetadata(title="Empty"), body=())
        data = write_docx_bytes(doc)
        assert len(data) > 0
        parts = _zip_parts(data)
        assert "word/document.xml" in parts

    def test_numbering_present_when_lists(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Lists"),
            body=(
                BulletList(
                    children=(ListItem(children=(Paragraph(children=(Text(value="item"),)),)),)
                ),
            ),
        )
        parts = _zip_parts(write_docx_bytes(doc))
        assert "word/numbering.xml" in parts


# ---------------------------------------------------------------------------
# Block serialization tests
# ---------------------------------------------------------------------------


class TestBlockSerialization:
    """Verify each block type is correctly serialized."""

    def test_paragraph(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="Hello world"),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        root.findall(f".//{{{W}}}p")
        texts = [t.text for t in root.findall(f".//{{{W}}}t")]
        assert "Hello world" in texts

    def test_heading_depth(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="H1"),)),
                Heading(depth=3, children=(Text(value="H3"),)),
            ),
        )
        root = _doc_xml(write_docx_bytes(doc))
        styles = root.findall(f".//{{{W}}}pStyle")
        style_vals = [s.get(f"{{{W}}}val") for s in styles]
        assert "Heading1" in style_vals
        assert "Heading3" in style_vals

    def test_bullet_list(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="A"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="B"),)),)),
                    )
                ),
            ),
        )
        root = _doc_xml(write_docx_bytes(doc))
        num_ids = root.findall(f".//{{{W}}}numId")
        assert len(num_ids) == 2
        # Bullet list uses numId=1
        assert all(n.get(f"{{{W}}}val") == "1" for n in num_ids)

    def test_ordered_list(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                OrderedList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="One"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Two"),)),)),
                    )
                ),
            ),
        )
        root = _doc_xml(write_docx_bytes(doc))
        num_ids = root.findall(f".//{{{W}}}numId")
        assert len(num_ids) == 2
        # Ordered list uses numId=2
        assert all(n.get(f"{{{W}}}val") == "2" for n in num_ids)

    def test_table_structure(self) -> None:
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
        root = _doc_xml(write_docx_bytes(doc))
        tables = root.findall(f".//{{{W}}}tbl")
        assert len(tables) == 1
        rows = tables[0].findall(f".//{{{W}}}tr")
        assert len(rows) == 2
        # Header row should have tblHeader
        header_markers = rows[0].findall(f".//{{{W}}}tblHeader")
        assert len(header_markers) == 1

    def test_code_block(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(CodeBlock(value="line1\nline2", language="python"),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        code_styles = [
            s for s in root.findall(f".//{{{W}}}pStyle") if s.get(f"{{{W}}}val") == "Code"
        ]
        assert len(code_styles) == 2  # one per line

    def test_page_break(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(PageBreak(),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        breaks = root.findall(f".//{{{W}}}br")
        assert any(b.get(f"{{{W}}}type") == "page" for b in breaks)


# ---------------------------------------------------------------------------
# Inline serialization tests
# ---------------------------------------------------------------------------


class TestInlineSerialization:
    def test_bold(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Strong(children=(Text(value="bold"),)),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        bolds = root.findall(f".//{{{W}}}b")
        assert len(bolds) >= 1

    def test_italic(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Emphasis(children=(Text(value="italic"),)),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        italics = root.findall(f".//{{{W}}}i")
        assert len(italics) >= 1

    def test_strikethrough(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Strikethrough(children=(Text(value="struck"),)),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        strikes = root.findall(f".//{{{W}}}strike")
        assert len(strikes) >= 1

    def test_inline_code(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Code(value="x = 1"),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        fonts = root.findall(f".//{{{W}}}rFonts")
        consolas = [f for f in fonts if f.get(f"{{{W}}}ascii") == "Consolas"]
        assert len(consolas) >= 1

    def test_link_styled(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(Link(url="https://example.com", children=(Text(value="click"),)),)
                ),
            ),
        )
        root = _doc_xml(write_docx_bytes(doc))
        colors = root.findall(f".//{{{W}}}color")
        blue = [c for c in colors if c.get(f"{{{W}}}val") == "0563C1"]
        assert len(blue) >= 1

    def test_line_break(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="a"), LineBreak(), Text(value="b"))),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        breaks = root.findall(f".//{{{W}}}br")
        assert len(breaks) >= 1

    def test_xml_space_preserve(self) -> None:
        """Verify xml:space='preserve' uses the correct XML namespace."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="  spaces  "),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        xml_ns = "http://www.w3.org/XML/1998/namespace"
        t_els = root.findall(f".//{{{W}}}t")
        assert any(t.get(f"{{{xml_ns}}}space") == "preserve" for t in t_els)


# ---------------------------------------------------------------------------
# Round-trip tests (fixture-based)
# ---------------------------------------------------------------------------


class TestRoundTrip:
    """Parse real DOCX → write → re-parse → verify content."""

    @pytest.fixture()
    def _write_and_reparse(self, tmp_path: Path):
        """Helper: write DOCX bytes, re-parse, return both texts."""

        def _inner(fixture_name: str) -> tuple[str, str, int, int]:
            src = parse_docx(FIXTURES / fixture_name)
            orig = serialize_text(src)
            data = write_docx_bytes(src)
            out = tmp_path / "output.docx"
            out.write_bytes(data)
            dst = parse_docx(out)
            rt = serialize_text(dst)
            return orig, rt, len(src.body), len(dst.body)

        return _inner

    def test_multi_paragraph(self, _write_and_reparse, tmp_path: Path) -> None:
        orig, rt, _n_orig, _n_rt = _write_and_reparse("MultiParagraphSample.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.95

    def test_cfpb_document(self, _write_and_reparse, tmp_path: Path) -> None:
        orig, rt, _n_orig, _n_rt = _write_and_reparse("bcfp_consumer-rights-summary_2018-09.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.95

    def test_large_document(self, _write_and_reparse, tmp_path: Path) -> None:
        orig, rt, _n_orig, _n_rt = _write_and_reparse(
            "mutual-to-stock-application-for-conversion.docx"
        )
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.95

    def test_treasury_document(self, _write_and_reparse, tmp_path: Path) -> None:
        orig, rt, _n_orig, _n_rt = _write_and_reparse("p2021-203386.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.95

    def test_cheese_sample(self, _write_and_reparse, tmp_path: Path) -> None:
        """CheeseSample.docx — 1MB, rich formatting (from kelvin-office)."""
        orig, rt, _n_orig, _n_rt = _write_and_reparse("CheeseSample.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.90

    def test_toro_term_loan(self, _write_and_reparse, tmp_path: Path) -> None:
        """Toro 2022 Term Loan — complex legal document with numbering."""
        orig, rt, _n_orig, _n_rt = _write_and_reparse("Toro 2022 Term Loan.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.90

    def test_right_to_use_leases(self, _write_and_reparse, tmp_path: Path) -> None:
        """565KB legal document with complex structure."""
        orig, rt, _n_orig, _n_rt = _write_and_reparse(
            "right-to-use-leases-with-operating-budget-treatment-with-cancellation-clause.docx"
        )
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.90

    def test_policy_procedure_template(self, _write_and_reparse, tmp_path: Path) -> None:
        """PolicyProcedureTemplate — structured template document."""
        orig, rt, _n_orig, _n_rt = _write_and_reparse(
            "PolicyProcedureTemplate_PhysicalFacility_Final.docx"
        )
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.90

    def test_cms_model_notice(self, _write_and_reparse, tmp_path: Path) -> None:
        """CMS HRA model notice — government form."""
        orig, rt, _n_orig, _n_rt = _write_and_reparse("cms-10704-hra-model-notice.docx")
        orig_words = set(orig.lower().split())
        rt_words = set(rt.lower().split())
        overlap = len(orig_words & rt_words) / len(orig_words) if orig_words else 1.0
        assert overlap >= 0.90


# ---------------------------------------------------------------------------
# Style round-trip tests (from kelvin-office TestStyleRoundTrip)
# ---------------------------------------------------------------------------


class TestStyleRoundTrip:
    """Verify heading styles survive write → re-parse cycle."""

    def test_heading_style_preserved(self, tmp_path: Path) -> None:
        """Heading pStyle must survive round-trip."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Heading(depth=1, children=(Text(value="Title"),)),
                Heading(depth=2, children=(Text(value="Subtitle"),)),
                Paragraph(children=(Text(value="Body text"),)),
            ),
        )
        out = tmp_path / "styles_rt.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)

        # Verify heading blocks survived
        headings = [b for b in doc2.body if type(b).__name__ == "Heading"]
        assert len(headings) >= 2
        assert headings[0].depth == 1
        assert headings[1].depth == 2

    def test_heading_style_in_xml(self) -> None:
        """pStyle element must reference correct Heading ID."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Heading(depth=3, children=(Text(value="H3"),)),),
        )
        root = _doc_xml(write_docx_bytes(doc))
        styles = root.findall(f".//{{{W}}}pStyle")
        vals = [s.get(f"{{{W}}}val") for s in styles]
        assert "Heading3" in vals

        # outlineLvl must also be present
        outline_lvls = root.findall(f".//{{{W}}}outlineLvl")
        assert any(o.get(f"{{{W}}}val") == "2" for o in outline_lvls)  # 0-indexed

    def test_all_six_heading_levels(self, tmp_path: Path) -> None:
        """All 6 heading levels round-trip with correct depth."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=tuple(Heading(depth=d, children=(Text(value=f"Level {d}"),)) for d in range(1, 7)),
        )
        out = tmp_path / "h6_rt.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)
        headings = [b for b in doc2.body if type(b).__name__ == "Heading"]
        assert [h.depth for h in headings] == [1, 2, 3, 4, 5, 6]


# ---------------------------------------------------------------------------
# Numbering round-trip tests (from kelvin-office TestNumberingRoundTrip)
# ---------------------------------------------------------------------------


class TestNumberingRoundTrip:
    """Verify list numbering survives write → re-parse cycle."""

    def test_bullet_list_roundtrip(self, tmp_path: Path) -> None:
        """Bullet list items remain bullet items after round-trip."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Alpha"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Beta"),)),)),
                    )
                ),
            ),
        )
        out = tmp_path / "bullet_rt.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)

        rt_text = serialize_text(doc2)
        assert "Alpha" in rt_text
        assert "Beta" in rt_text
        # Reader should detect list items
        bullet_lists = [b for b in doc2.body if type(b).__name__ == "BulletList"]
        assert len(bullet_lists) >= 1

    def test_ordered_list_roundtrip(self, tmp_path: Path) -> None:
        """Ordered list items remain ordered after round-trip."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                OrderedList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Step 1"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Step 2"),)),)),
                    )
                ),
            ),
        )
        out = tmp_path / "ordered_rt.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)

        rt_text = serialize_text(doc2)
        assert "Step 1" in rt_text
        assert "Step 2" in rt_text

    def test_numbering_xml_has_abstract_and_num(self) -> None:
        """numbering.xml must contain abstractNum and num elements."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                BulletList(
                    children=(ListItem(children=(Paragraph(children=(Text(value="A"),)),)),)
                ),
            ),
        )
        data = write_docx_bytes(doc)
        zf = zipfile.ZipFile(BytesIO(data))
        numbering_xml = zf.read("word/numbering.xml")
        root = etree.fromstring(numbering_xml)

        abstract_nums = root.findall(f".//{{{W}}}abstractNum")
        assert len(abstract_nums) >= 1

        nums = root.findall(f".//{{{W}}}num")
        assert len(nums) >= 1


# ---------------------------------------------------------------------------
# Creation-from-scratch tests (from kelvin-office TestWordClient)
# ---------------------------------------------------------------------------


class TestDocumentCreation:
    """Build documents programmatically → write → reload → verify."""

    def test_comprehensive_document(self, tmp_path: Path) -> None:
        """Create a document with all supported block types and verify round-trip."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Comprehensive Test"),
            body=(
                Heading(depth=1, children=(Text(value="Main Title"),)),
                Paragraph(
                    children=(
                        Text(value="Normal, "),
                        Strong(children=(Text(value="bold"),)),
                        Text(value=", "),
                        Emphasis(children=(Text(value="italic"),)),
                        Text(value=", "),
                        Code(value="code"),
                        Text(value="."),
                    )
                ),
                Heading(depth=2, children=(Text(value="Lists Section"),)),
                BulletList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Bullet A"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Bullet B"),)),)),
                    )
                ),
                OrderedList(
                    children=(
                        ListItem(children=(Paragraph(children=(Text(value="Item 1"),)),)),
                        ListItem(children=(Paragraph(children=(Text(value="Item 2"),)),)),
                    )
                ),
                Heading(depth=2, children=(Text(value="Table Section"),)),
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
                BlockQuote(children=(Paragraph(children=(Text(value="A famous quote."),)),)),
                ThematicBreak(),
                Paragraph(children=(Text(value="Final paragraph."),)),
            ),
        )

        out = tmp_path / "comprehensive.docx"
        write_docx(doc, out)
        assert out.exists()
        assert out.stat().st_size > 0

        doc2 = parse_docx(out)
        rt_text = serialize_text(doc2)

        # Verify key content survived
        assert "Main Title" in rt_text
        assert "bold" in rt_text
        assert "italic" in rt_text
        assert "code" in rt_text
        assert "Bullet A" in rt_text
        assert "Item 1" in rt_text
        assert "Alpha" in rt_text
        assert "hello()" in rt_text
        assert "famous quote" in rt_text
        assert "Final paragraph" in rt_text

    def test_write_creates_parent_dirs(self, tmp_path: Path) -> None:
        """write_docx must create intermediate directories."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="test"),)),),
        )
        out = tmp_path / "a" / "b" / "c" / "output.docx"
        result = write_docx(doc, out)
        assert result == out
        assert out.exists()


# ---------------------------------------------------------------------------
# Error handling tests (from kelvin-office test_error_handling)
# ---------------------------------------------------------------------------


class TestErrorHandling:
    def test_write_empty_body(self) -> None:
        """Empty body should produce a valid DOCX, not crash."""
        doc = ContentDocument(metadata=DocumentMetadata(title="Empty"), body=())
        data = write_docx_bytes(doc)
        assert len(data) > 0
        # Must be a valid ZIP
        zf = zipfile.ZipFile(BytesIO(data))
        assert "word/document.xml" in zf.namelist()

    def test_write_preserves_title_metadata(self) -> None:
        """Title in metadata should appear in docProps/core.xml."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title="My Important Document"),
            body=(Paragraph(children=(Text(value="content"),)),),
        )
        data = write_docx_bytes(doc)
        zf = zipfile.ZipFile(BytesIO(data))
        core_xml = zf.read("docProps/core.xml").decode()
        assert "My Important Document" in core_xml


# ---------------------------------------------------------------------------
# Performance tests (from kelvin-office test_performance_large_document)
# ---------------------------------------------------------------------------


class TestPerformance:
    def test_large_document_write_under_3s(self) -> None:
        """Writing a 500-block document should complete under 3 seconds."""
        blocks = []
        for i in range(500):
            blocks.append(
                Paragraph(
                    children=(
                        Text(value=f"Paragraph {i} with some content to make it realistic. "),
                        Strong(children=(Text(value="Bold part"),)),
                        Text(value=" and more text."),
                    )
                )
            )
        doc = ContentDocument(metadata=DocumentMetadata(title="Large"), body=tuple(blocks))

        start = time.monotonic()
        data = write_docx_bytes(doc)
        elapsed = time.monotonic() - start

        assert elapsed < 3.0, f"write_docx_bytes took {elapsed:.2f}s (budget 3s)"
        assert len(data) > 0

    def test_fixture_roundtrip_under_5s(self, tmp_path: Path) -> None:
        """Parse → write → re-parse of largest fixture under 5s total."""
        fixture = (
            FIXTURES
            / "right-to-use-leases-with-operating-budget-treatment-with-cancellation-clause.docx"
        )
        if not fixture.exists():
            pytest.skip("Fixture not available")

        start = time.monotonic()
        doc = parse_docx(fixture)
        data = write_docx_bytes(doc)
        out = tmp_path / "perf.docx"
        out.write_bytes(data)
        parse_docx(out)
        elapsed = time.monotonic() - start

        assert elapsed < 5.0, f"Full round-trip took {elapsed:.2f}s (budget 5s)"
