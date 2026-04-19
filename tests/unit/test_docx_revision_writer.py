"""Unit tests for DOCX revision writing (w:ins / w:del / w:moveFrom / w:moveTo).

Complements test_docx_revisions.py which covers the reader side. Together
they verify that the writer is the exact inverse of the parser:
track-changes documents round-trip end-to-end with full metadata.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

from kaos_content.model.attr import Attr
from kaos_content.model.blocks import Div, Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Span, Text
from kaos_content.revision import Revisions
from kaos_content.serializers.text import serialize_text

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx, write_docx_bytes

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _doc_xml(docx_bytes: bytes) -> str:
    zf = zipfile.ZipFile(BytesIO(docx_bytes))
    return zf.read("word/document.xml").decode()


# ---------------------------------------------------------------------------
# Inline revisions
# ---------------------------------------------------------------------------


class TestInlineRevisions:
    def test_rev_ins_span_emits_w_ins(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="hello "),
                        Span(
                            attr=Attr(
                                classes=("rev-ins",),
                                kv={
                                    "rev:id": "1",
                                    "rev:author": "Alice",
                                    "rev:date": "2026-04-18T10:00:00Z",
                                },
                            ),
                            children=(Text(value="world"),),
                        ),
                    )
                ),
            ),
        )
        xml = _doc_xml(write_docx_bytes(doc))
        assert "<w:ins " in xml
        assert 'w:id="1"' in xml
        assert 'w:author="Alice"' in xml
        assert 'w:date="2026-04-18T10:00:00Z"' in xml

    def test_rev_del_span_emits_w_del_with_delText(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="keep "),
                        Span(
                            attr=Attr(
                                classes=("rev-del",),
                                kv={"rev:id": "2", "rev:author": "Bob"},
                            ),
                            children=(Text(value="drop me"),),
                        ),
                    )
                ),
            ),
        )
        xml = _doc_xml(write_docx_bytes(doc))
        assert "<w:del " in xml
        # OOXML §17.16.2: deleted text uses w:delText, not w:t
        assert "<w:delText" in xml
        assert ">drop me</w:delText>" in xml
        # The w:t of unrelated runs must remain
        assert "<w:t " in xml and "keep " in xml

    def test_rev_move_to_and_move_from_emit_correct_tags(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Span(
                            attr=Attr(
                                classes=("rev-move-to",),
                                kv={"rev:id": "3", "rev:author": "Alice", "rev:move-name": "m1"},
                            ),
                            children=(Text(value="moved"),),
                        ),
                    )
                ),
                Paragraph(
                    children=(
                        Span(
                            attr=Attr(
                                classes=("rev-move-from",),
                                kv={"rev:id": "4", "rev:author": "Alice", "rev:move-name": "m1"},
                            ),
                            children=(Text(value="moved"),),
                        ),
                    )
                ),
            ),
        )
        xml = _doc_xml(write_docx_bytes(doc))
        assert "<w:moveTo " in xml
        assert "<w:moveFrom " in xml
        assert 'w:name="m1"' in xml


# ---------------------------------------------------------------------------
# Block-level revisions (Div wrappers)
# ---------------------------------------------------------------------------


class TestBlockRevisions:
    def test_rev_ins_div_emits_block_level_w_ins(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(children=(Text(value="Existing."),)),
                Div(
                    attr=Attr(
                        classes=("rev-ins",),
                        kv={"rev:id": "5", "rev:author": "Bob", "rev:date": "2026-04-18T11:00:00Z"},
                    ),
                    children=(Paragraph(children=(Text(value="Inserted."),)),),
                ),
            ),
        )
        xml = _doc_xml(write_docx_bytes(doc))
        assert "<w:ins " in xml
        assert 'w:id="5"' in xml
        # The inserted paragraph should appear inside w:ins
        assert "Inserted." in xml

    def test_rev_del_div_uses_delText(self) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Div(
                    attr=Attr(
                        classes=("rev-del",),
                        kv={"rev:id": "6", "rev:author": "Bob"},
                    ),
                    children=(Paragraph(children=(Text(value="Deleted paragraph."),)),),
                ),
            ),
        )
        xml = _doc_xml(write_docx_bytes(doc))
        assert "<w:del " in xml
        assert "<w:delText" in xml
        assert ">Deleted paragraph.</w:delText>" in xml


# ---------------------------------------------------------------------------
# Round-trip through reader
# ---------------------------------------------------------------------------


class TestRevisionRoundTrip:
    def test_synthetic_roundtrip_preserves_metadata(self, tmp_path: Path) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title=""),
            body=(
                Paragraph(
                    children=(
                        Text(value="The deadline is "),
                        Span(
                            attr=Attr(
                                classes=("rev-del",),
                                kv={
                                    "rev:id": "0",
                                    "rev:author": "Alice",
                                    "rev:date": "2026-04-15T10:30:00Z",
                                },
                            ),
                            children=(Text(value="Monday"),),
                        ),
                        Span(
                            attr=Attr(
                                classes=("rev-ins",),
                                kv={
                                    "rev:id": "1",
                                    "rev:author": "Alice",
                                    "rev:date": "2026-04-15T10:30:00Z",
                                },
                            ),
                            children=(Text(value="Friday"),),
                        ),
                        Text(value="."),
                    )
                ),
            ),
        )
        out = tmp_path / "rev.docx"
        write_docx(doc, out)
        doc2 = parse_docx(out, track_changes=True)
        revs = Revisions.from_document(doc2)
        assert len(revs) == 2
        by_id = {r.id: r for r in revs}
        assert by_id["0"].author == "Alice"
        assert by_id["0"].change_type.value == "deletion"
        assert by_id["1"].change_type.value == "insertion"

    def test_toro_redline_full_roundtrip(self, tmp_path: Path) -> None:
        """End-to-end: parse redline → write → re-parse with same revision count + metadata."""
        src = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        src_revs = Revisions.from_document(src)
        assert len(src_revs) > 0

        out = tmp_path / "toro_rt.docx"
        write_docx(src, out)
        dst = parse_docx(out, track_changes=True)
        dst_revs = Revisions.from_document(dst)

        # Count match
        assert len(dst_revs) == len(src_revs)

        # Per-ID metadata match
        src_by_id = {r.id: r for r in src_revs}
        dst_by_id = {r.id: r for r in dst_revs}
        assert src_by_id.keys() == dst_by_id.keys()
        for rid in src_by_id:
            s, d = src_by_id[rid], dst_by_id[rid]
            assert s.author == d.author, f"author mismatch for rev {rid}"
            assert s.change_type == d.change_type, f"type mismatch for rev {rid}"

    def test_markup_view_roundtrip(self, tmp_path: Path) -> None:
        """The markup serialization should have high word overlap across round-trip."""
        src = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        out = tmp_path / "markup_rt.docx"
        write_docx(src, out)
        dst = parse_docx(out, track_changes=True)

        src_markup = serialize_text(src, view="markup")
        dst_markup = serialize_text(dst, view="markup")
        src_words = set(src_markup.split())
        dst_words = set(dst_markup.split())
        overlap = len(src_words & dst_words) / len(src_words) if src_words else 0
        assert overlap >= 0.95
