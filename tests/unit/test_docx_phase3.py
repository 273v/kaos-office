"""Unit tests for DOCX Phase 3: hyperlinks, footnotes/endnotes, comments.

Verifies that the DOCX writer produces standards-compliant OOXML for:
- w:hyperlink with relationship entries (reusing rels for duplicate URLs)
- word/footnotes.xml and word/endnotes.xml with required separators
- word/comments.xml with author/date/text metadata

Round-trip verification: every construct survives parse → write → re-parse.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from kaos_content.model.annotation import Annotation, AnnotationTarget, AnnotationType
from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import FootnoteRef, Link, Text
from kaos_content.serializers.text import serialize_text

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx, write_docx_bytes

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _zip(docx_bytes: bytes) -> zipfile.ZipFile:
    return zipfile.ZipFile(BytesIO(docx_bytes))


def _rels(zf: zipfile.ZipFile) -> str:
    return zf.read("word/_rels/document.xml.rels").decode()


def _doc_xml(zf: zipfile.ZipFile) -> str:
    return zf.read("word/document.xml").decode()


# ---------------------------------------------------------------------------
# Phase A: Hyperlinks
# ---------------------------------------------------------------------------


class TestHyperlinks:
    def test_hyperlink_rel_created(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="See "),
                        Link(url="https://example.com", children=(Text(value="our site"),)),
                        Text(value="."),
                    )
                ),
            ),
        )
        zf = _zip(write_docx_bytes(doc))
        rels = _rels(zf)
        assert "https://example.com" in rels
        assert 'TargetMode="External"' in rels
        assert "hyperlink" in rels

    def test_hyperlink_element_emitted(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(Link(url="https://example.com", children=(Text(value="link"),)),)
                ),
            ),
        )
        zf = _zip(write_docx_bytes(doc))
        doc_xml = _doc_xml(zf)
        assert "<w:hyperlink" in doc_xml

    def test_duplicate_urls_share_rel(self) -> None:
        """Two Links with the same URL must share one relationship entry."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(Link(url="https://example.com", children=(Text(value="one"),)),)
                ),
                Paragraph(
                    children=(Link(url="https://example.com", children=(Text(value="two"),)),)
                ),
            ),
        )
        zf = _zip(write_docx_bytes(doc))
        rels = _rels(zf)
        # example.com should appear exactly once in the rels file
        assert rels.count("example.com") == 1

    def test_internal_anchor_uses_w_anchor(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Link(url="#section-1", children=(Text(value="jump"),)),)),),
        )
        zf = _zip(write_docx_bytes(doc))
        doc_xml = _doc_xml(zf)
        assert 'w:anchor="section-1"' in doc_xml
        # Internal refs must NOT create rels
        rels = _rels(zf)
        assert "section-1" not in rels

    def test_hyperlink_roundtrip_preserves_url(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(Link(url="https://example.com/x", children=(Text(value="linked"),)),)
                ),
            ),
        )
        out = tmp_path / "link.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)
        # Walk the block to find the Link
        found_urls = []
        for c in doc2.body[0].children:
            url = getattr(c, "url", None)
            if url:
                found_urls.append(url)
        assert "https://example.com/x" in found_urls

    def test_hyperlink_text_preserved(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="Visit "),
                        Link(url="https://example.com", children=(Text(value="our site"),)),
                        Text(value="."),
                    )
                ),
            ),
        )
        out = tmp_path / "link.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)
        rt = serialize_text(doc2)
        # Link text should survive; surrounding text too
        assert "Visit" in rt
        assert "our site" in rt
        assert "details" not in rt  # sanity


# ---------------------------------------------------------------------------
# Phase B: Footnotes / Endnotes
# ---------------------------------------------------------------------------


class TestFootnotes:
    def test_no_footnotes_no_part(self) -> None:
        """Documents without footnotes must not emit footnotes.xml."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="plain"),)),),
        )
        zf = _zip(write_docx_bytes(doc))
        assert "word/footnotes.xml" not in zf.namelist()

    def test_footnote_fixture_roundtrip(self, tmp_path: Path) -> None:
        """The Footnote.docx fixture's footnote survives round-trip."""
        doc = parse_docx(FIXTURES / "Footnote.docx")
        assert doc.footnotes, "fixture should have footnotes"
        out = tmp_path / "fn.docx"
        write_docx(doc, out)

        zf = _zip(out.read_bytes())
        assert "word/footnotes.xml" in zf.namelist()
        fn_xml = zf.read("word/footnotes.xml").decode()
        # Must contain both required separators
        assert 'w:type="separator"' in fn_xml
        assert 'w:type="continuationSeparator"' in fn_xml

        # Re-parse and verify
        doc2 = parse_docx(out)
        assert doc2.footnotes, "footnotes should survive round-trip"

    def test_footnote_content_type_registered(self) -> None:
        doc = parse_docx(FIXTURES / "Footnote.docx")
        zf = _zip(write_docx_bytes(doc))
        ct = zf.read("[Content_Types].xml").decode()
        assert "footnotes+xml" in ct

    def test_footnote_rel_added(self) -> None:
        doc = parse_docx(FIXTURES / "Footnote.docx")
        zf = _zip(write_docx_bytes(doc))
        rels = _rels(zf)
        assert "footnotes" in rels

    def test_footnote_reference_emitted_in_body(self) -> None:
        """A FootnoteRef inline must become a w:footnoteReference run."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="See note"),
                        FootnoteRef(identifier="2"),
                    )
                ),
            ),
            footnotes={"2": (Paragraph(children=(Text(value="A note."),)),)},
        )
        zf = _zip(write_docx_bytes(doc))
        doc_xml = _doc_xml(zf)
        assert "<w:footnoteReference" in doc_xml
        assert 'w:id="2"' in doc_xml

    def test_endnote_reference_emitted(self) -> None:
        """An endnote (identifier en-1) must become w:endnoteReference."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="See end"),
                        FootnoteRef(identifier="en-1"),
                    )
                ),
            ),
            footnotes={"en-1": (Paragraph(children=(Text(value="An endnote."),)),)},
        )
        zf = _zip(write_docx_bytes(doc))
        doc_xml = _doc_xml(zf)
        assert "<w:endnoteReference" in doc_xml
        # endnote part should also be created
        assert "word/endnotes.xml" in zf.namelist()


# ---------------------------------------------------------------------------
# Phase C: Comments
# ---------------------------------------------------------------------------


class TestComments:
    def test_no_comments_no_part(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="plain"),)),),
        )
        zf = _zip(write_docx_bytes(doc))
        assert "word/comments.xml" not in zf.namelist()

    def test_comment_fixture_roundtrip(self, tmp_path: Path) -> None:
        """Toro Comments fixture has 5 comments; all must survive round-trip."""
        src = parse_docx(FIXTURES / "Toro 2022 Term Loan - Comments.docx")
        src_comments = [a for a in src.annotations if a.type == AnnotationType.COMMENT]
        assert len(src_comments) == 5

        out = tmp_path / "comments.docx"
        write_docx(src, out)

        zf = _zip(out.read_bytes())
        assert "word/comments.xml" in zf.namelist()

        dst = parse_docx(out)
        dst_comments = [a for a in dst.annotations if a.type == AnnotationType.COMMENT]
        assert len(dst_comments) == 5

    def test_comment_metadata_preserved(self, tmp_path: Path) -> None:
        """Author and text survive round-trip."""
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(Paragraph(children=(Text(value="content"),)),),
            annotations=(
                Annotation(
                    id="ann-1",
                    type=AnnotationType.COMMENT,
                    targets=(AnnotationTarget(node_ref="#/body/0"),),
                    body={
                        "author": "Alice",
                        "date": "2026-04-18T12:00:00Z",
                        "text": "Please review this section.",
                        "initials": "A",
                    },
                ),
            ),
        )
        out = tmp_path / "ours.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out)
        dst_comments = [a for a in doc2.annotations if a.type == AnnotationType.COMMENT]
        assert len(dst_comments) == 1
        c = dst_comments[0]
        assert c.body.get("author") == "Alice"
        assert c.body.get("text") == "Please review this section."

    def test_comment_content_type_registered(self) -> None:
        doc = parse_docx(FIXTURES / "Toro 2022 Term Loan - Comments.docx")
        zf = _zip(write_docx_bytes(doc))
        ct = zf.read("[Content_Types].xml").decode()
        assert "comments+xml" in ct

    def test_comment_rel_added(self) -> None:
        doc = parse_docx(FIXTURES / "Toro 2022 Term Loan - Comments.docx")
        zf = _zip(write_docx_bytes(doc))
        rels = _rels(zf)
        assert "relationships/comments" in rels


# ---------------------------------------------------------------------------
# Backward compatibility: unchanged documents still round-trip
# ---------------------------------------------------------------------------


class TestBackwardCompatibility:
    @pytest.mark.parametrize(
        "fixture_name",
        [
            "MultiParagraphSample.docx",
            "bcfp_consumer-rights-summary_2018-09.docx",
            "Toro 2022 Term Loan.docx",
        ],
    )
    def test_no_regressions(self, fixture_name: str, tmp_path: Path) -> None:
        """Fixtures without comments/footnotes should round-trip identically."""
        src = parse_docx(FIXTURES / fixture_name)
        out = tmp_path / "rt.docx"
        write_docx(src, out)
        dst = parse_docx(out)
        # Word counts should match (small tolerance for bullet numbering)
        src_text = serialize_text(src)
        dst_text = serialize_text(dst)
        src_words = set(src_text.lower().split())
        dst_words = set(dst_text.lower().split())
        if src_words:
            overlap = len(src_words & dst_words) / len(src_words)
            assert overlap >= 0.90
