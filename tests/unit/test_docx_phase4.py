"""Tests for DOCX Phase 4 — headers, footers, and page setup round-trip."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.model.metadata import PageSetup
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx, write_docx_bytes
from kaos_office.ooxml.namespace import (
    CT_FOOTER,
    CT_HEADER,
    W_BODY,
    W_FOOTER_REFERENCE,
    W_HDR,
    W_HEADER_REFERENCE,
    W_PGMAR,
    W_PGSZ,
    W_SECTPR,
    W,
    pt_to_twips,
    qn,
    twips_to_pt,
)

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"


def _fixture_with_footers() -> Path | None:
    """Pick a DOCX fixture that actually has a footer part inside."""
    for p in sorted(FIXTURES.glob("*.docx")):
        try:
            with zipfile.ZipFile(p) as zf:
                names = zf.namelist()
        except zipfile.BadZipFile:
            continue
        if any(n.startswith("word/footer") and n.endswith(".xml") for n in names):
            return p
    return None


class TestTwipsHelpers:
    def test_roundtrip_integer_twips(self) -> None:
        # 1440 twips = 72 pt = 1 inch
        assert twips_to_pt(1440) == pytest.approx(72.0)
        assert pt_to_twips(72.0) == 1440

    def test_zero(self) -> None:
        assert twips_to_pt(0) == 0.0
        assert pt_to_twips(0) == 0

    def test_fractional_points(self) -> None:
        # 1/20 pt resolution; values within the step round cleanly
        assert pt_to_twips(10.5) == 210


class TestReaderPageSetup:
    """Reader extracts page size / margins from the body's final <w:sectPr>."""

    def test_letter_page(self) -> None:
        fixture = _fixture_with_footers()
        if fixture is None:
            pytest.skip("no DOCX fixture with footer")
        doc = parse_docx(fixture)
        ps = doc.metadata.page_setup
        assert ps is not None
        # Letter is 612 x 792 pt. Fixtures in this repo all use it.
        assert ps.page_width_pt == pytest.approx(612.0, abs=1.0)
        assert ps.page_height_pt == pytest.approx(792.0, abs=1.0)
        # Margins should be positive and each fit within the page.
        assert ps.margin_top_pt is not None and 0 <= ps.margin_top_pt < 612
        assert ps.margin_left_pt is not None and 0 <= ps.margin_left_pt < 612


class TestReaderFooters:
    """Reader captures footer content from word/footer*.xml parts."""

    def test_footer_parsed(self) -> None:
        fixture = _fixture_with_footers()
        if fixture is None:
            pytest.skip("no DOCX fixture with footer")
        doc = parse_docx(fixture)
        assert "default" in doc.footers
        blocks = doc.footers["default"]
        assert len(blocks) > 0

    def test_no_footers_fixture_has_empty_dict(self) -> None:
        # make_minimal_docx has no footer part → footers dict stays empty
        import io

        from tests.conftest import make_minimal_docx

        data = make_minimal_docx()
        import tempfile

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            tf.write(data)
            p = Path(tf.name)
        try:
            doc = parse_docx(p)
            assert doc.footers == {}
            assert doc.headers == {}
        finally:
            p.unlink(missing_ok=True)
        # silence unused import warning
        _ = io


class TestWriterPageSetup:
    """Writer emits real page geometry when doc.metadata.page_setup is set."""

    def test_custom_page_setup_emitted(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(
                title="",
                page_setup=PageSetup(
                    page_width_pt=595.0,  # A4 width (~595 pt)
                    page_height_pt=842.0,  # A4 height (~842 pt)
                    margin_top_pt=50.0,
                    margin_bottom_pt=50.0,
                    margin_left_pt=40.0,
                    margin_right_pt=40.0,
                    header_distance_pt=30.0,
                    footer_distance_pt=30.0,
                ),
            ),
            body=(Paragraph(children=(Text(value="Body"),)),),
        )
        data = write_docx_bytes(doc)
        import io

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("word/document.xml")
        root = etree.fromstring(xml)
        body = root.find(W_BODY)
        assert body is not None
        sect_pr = body.find(W_SECTPR)
        assert sect_pr is not None
        pg_sz = sect_pr.find(W_PGSZ)
        assert pg_sz is not None
        assert pg_sz.get(qn(W, "w")) == str(pt_to_twips(595.0))
        assert pg_sz.get(qn(W, "h")) == str(pt_to_twips(842.0))
        pg_mar = sect_pr.find(W_PGMAR)
        assert pg_mar is not None
        assert pg_mar.get(qn(W, "top")) == str(pt_to_twips(50.0))
        assert pg_mar.get(qn(W, "left")) == str(pt_to_twips(40.0))
        assert pg_mar.get(qn(W, "footer")) == str(pt_to_twips(30.0))

    def test_no_page_setup_defaults_to_letter(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(),
            body=(Paragraph(children=(Text(value="Body"),)),),
        )
        data = write_docx_bytes(doc)
        import io

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("word/document.xml")
        root = etree.fromstring(xml)
        body = root.find(W_BODY)
        sect_pr = body.find(W_SECTPR) if body is not None else None
        assert sect_pr is not None
        pg_sz = sect_pr.find(W_PGSZ)
        assert pg_sz is not None
        # Letter — unchanged default
        assert pg_sz.get(qn(W, "w")) == "12240"
        assert pg_sz.get(qn(W, "h")) == "15840"


class TestWriterHeadersFooters:
    """Writer emits header/footer parts + sectPr references."""

    def test_header_and_footer_parts_written(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="Body"),)),),
            headers={"default": (Paragraph(children=(Text(value="HEADER MARKER"),)),)},
            footers={"default": (Paragraph(children=(Text(value="FOOTER MARKER"),)),)},
        )
        data = write_docx_bytes(doc)
        import io

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = zf.namelist()
            hdr = zf.read("word/header1.xml").decode("utf-8")
            ftr = zf.read("word/footer1.xml").decode("utf-8")
            ct = zf.read("[Content_Types].xml").decode("utf-8")
            rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")

        assert "word/header1.xml" in names
        assert "word/footer1.xml" in names
        assert "HEADER MARKER" in hdr
        assert "FOOTER MARKER" in ftr
        assert CT_HEADER in ct
        assert CT_FOOTER in ct
        assert "relationships/header" in rels
        assert "relationships/footer" in rels

    def test_sectpr_references_header_and_footer(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="x"),)),),
            headers={"default": (Paragraph(children=(Text(value="H"),)),)},
            footers={"first": (Paragraph(children=(Text(value="F"),)),)},
        )
        data = write_docx_bytes(doc)
        import io

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("word/document.xml")
        root = etree.fromstring(xml)
        body = root.find(W_BODY)
        sect_pr = body.find(W_SECTPR) if body is not None else None
        assert sect_pr is not None
        hdr_refs = sect_pr.findall(W_HEADER_REFERENCE)
        ftr_refs = sect_pr.findall(W_FOOTER_REFERENCE)
        assert len(hdr_refs) == 1
        assert hdr_refs[0].get(qn(W, "type")) == "default"
        assert len(ftr_refs) == 1
        assert ftr_refs[0].get(qn(W, "type")) == "first"

    def test_header_content_matches(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={
                "default": (
                    Paragraph(children=(Text(value="line 1"),)),
                    Paragraph(children=(Text(value="line 2"),)),
                )
            },
        )
        data = write_docx_bytes(doc)
        import io

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml = zf.read("word/header1.xml")
        root = etree.fromstring(xml)
        assert root.tag == W_HDR
        paragraphs = root.findall(qn(W, "p"))
        # header root should carry both paragraphs
        texts = [etree.tostring(p, encoding="unicode") for p in paragraphs]
        combined = "".join(texts)
        assert "line 1" in combined
        assert "line 2" in combined


class TestHeaderFooterRoundTrip:
    """parse → write → re-parse preserves header / footer content."""

    def test_construct_then_roundtrip(self, tmp_path: Path) -> None:
        orig = ContentDocument(
            body=(Paragraph(children=(Text(value="Body"),)),),
            headers={"default": (Paragraph(children=(Text(value="ROUND-TRIP HEADER"),)),)},
            footers={"default": (Paragraph(children=(Text(value="ROUND-TRIP FOOTER"),)),)},
            metadata=DocumentMetadata(
                page_setup=PageSetup(page_width_pt=612.0, page_height_pt=792.0)
            ),
        )
        out = tmp_path / "rt.docx"
        write_docx(orig, out)
        reloaded = parse_docx(out)
        assert "default" in reloaded.headers
        assert "default" in reloaded.footers
        header_text = "".join(
            getattr(c, "value", "")
            for block in reloaded.headers["default"]
            for c in getattr(block, "children", ())
        )
        footer_text = "".join(
            getattr(c, "value", "")
            for block in reloaded.footers["default"]
            for c in getattr(block, "children", ())
        )
        assert "ROUND-TRIP HEADER" in header_text
        assert "ROUND-TRIP FOOTER" in footer_text
        # Page geometry also preserved
        ps = reloaded.metadata.page_setup
        assert ps is not None
        assert ps.page_width_pt == pytest.approx(612.0, abs=1.0)

    def test_fixture_footer_roundtrip(self, tmp_path: Path) -> None:
        fixture = _fixture_with_footers()
        if fixture is None:
            pytest.skip("no DOCX fixture with footer")
        src = parse_docx(fixture)
        out = tmp_path / "rt.docx"
        write_docx(src, out)
        reloaded = parse_docx(out)
        # Footer count preserved (may not be textual identity due to style
        # normalization, but the structural presence must survive).
        assert set(reloaded.footers.keys()) == set(src.footers.keys())


class TestBuilderHeaderFooterAPI:
    """The DocumentBuilder shortcut methods set_header / set_footer populate
    the ContentDocument.headers / .footers dicts."""

    def test_builder_methods(self) -> None:
        from kaos_content.builders.builder import DocumentBuilder

        b = DocumentBuilder(title="Demo")
        b.paragraph("body")
        b.set_header(
            "default",
            Paragraph(children=(Text(value="H"),)),
        )
        b.set_footer(
            "default",
            Paragraph(children=(Text(value="F"),)),
        )
        doc = b.build()
        assert "default" in doc.headers
        assert "default" in doc.footers
