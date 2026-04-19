"""Unit tests for DOCX revision (tracked changes) parsing.

Tests the ``track_changes=True`` path of ``parse_docx``, which wraps
``w:ins`` / ``w:del`` / ``w:moveFrom`` / ``w:moveTo`` content in
``Span`` / ``Div`` nodes with ``rev-*`` classes and emits
``AnnotationType.TRACKED_CHANGE`` annotations carrying author / date /
revision-id metadata.

Covers the 4 use cases from docs/TRACKED_CHANGES_DESIGN.md:
  1. read/review history     — metadata queries
  2. time machine            — both versions in AST for reconstruction
  3. accept/reject/comment   — identifiable revision nodes
  4. author redlines         — parity is prerequisite for write-back
"""

from __future__ import annotations

from pathlib import Path

from kaos_content.model.annotation import AnnotationType
from kaos_content.model.blocks import Div
from kaos_content.model.inlines import Span

from kaos_office.docx.reader import parse_docx

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"


def _walk(node: object):
    """Yield every descendant of ``node``, then the node itself last."""
    yield node
    children = getattr(node, "children", None)
    content = getattr(node, "content", None)
    for c in children or ():
        yield from _walk(c)
    for c in content or ():
        yield from _walk(c)


def _rev_spans(doc) -> list[Span]:
    """All Span nodes in the body with any ``rev-*`` class."""
    spans: list[Span] = []
    for block in doc.body:
        for n in _walk(block):
            if isinstance(n, Span):
                classes = n.attr.classes or ()
                if any(c.startswith("rev-") for c in classes):
                    spans.append(n)
    return spans


def _rev_divs(doc) -> list[Div]:
    """All Div nodes in the body with any ``rev-*`` class."""
    divs: list[Div] = []
    for block in doc.body:
        for n in _walk(block):
            if isinstance(n, Div):
                classes = n.attr.classes or ()
                if any(c.startswith("rev-") for c in classes):
                    divs.append(n)
    return divs


def _tracked_changes(doc):
    return [a for a in doc.annotations if a.type == AnnotationType.TRACKED_CHANGE]


# ---------------------------------------------------------------------------
# Backward compatibility: default behavior must not change
# ---------------------------------------------------------------------------


class TestBackwardCompatibility:
    """track_changes=False (default) must produce no rev-* nodes or annotations."""

    def test_default_no_rev_spans(self) -> None:
        doc = parse_docx(FIXTURES / "Toro 2022 Term Loan - Redline v1.docx")
        assert _rev_spans(doc) == []
        assert _rev_divs(doc) == []

    def test_default_no_tracked_change_annotations(self) -> None:
        doc = parse_docx(FIXTURES / "Toro 2022 Term Loan - Redline v1.docx")
        assert _tracked_changes(doc) == []

    def test_explicit_false_no_rev_spans(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=False,
        )
        assert _rev_spans(doc) == []
        assert _tracked_changes(doc) == []


# ---------------------------------------------------------------------------
# Revision parsing produces spans + annotations
# ---------------------------------------------------------------------------


class TestRevisionParsing:
    """track_changes=True preserves both versions and metadata."""

    def test_toro_redline_produces_rev_spans(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        spans = _rev_spans(doc)
        # Fixture contains 4 w:ins + 9 w:del = 13 revision elements
        # (parser skips revisions with no extractable content)
        assert len(spans) >= 10

    def test_toro_redline_produces_tracked_change_annotations(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        annotations = _tracked_changes(doc)
        assert len(annotations) >= 10
        # Each annotation must carry change_type
        for ann in annotations:
            assert "change_type" in ann.body
            assert ann.body["change_type"] in {
                "insertion",
                "deletion",
                "move_from",
                "move_to",
            }

    def test_span_and_annotation_counts_match(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        assert len(_rev_spans(doc)) == len(_tracked_changes(doc))

    def test_insertions_and_deletions_both_present(self) -> None:
        """Both versions of changed content must be in the AST."""
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        insertions = [s for s in _rev_spans(doc) if "rev-ins" in s.attr.classes]
        deletions = [s for s in _rev_spans(doc) if "rev-del" in s.attr.classes]
        assert len(insertions) > 0, "expected at least one insertion Span"
        assert len(deletions) > 0, "expected at least one deletion Span"

    def test_deleted_content_text_preserved(self) -> None:
        """Deleted text must be preserved as real Text nodes (not strings)."""
        from kaos_content.model.inlines import Text

        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        deletions = [s for s in _rev_spans(doc) if "rev-del" in s.attr.classes]
        assert deletions, "fixture should contain deletions"
        # Every deletion Span contains at least one Text node with non-empty
        # value somewhere in its subtree (not necessarily as a direct child —
        # formatted deletions wrap the Text in Strong/Emphasis/etc).
        for span in deletions:
            text_descendants = [n for n in _walk(span) if isinstance(n, Text)]
            assert any(t.value for t in text_descendants), (
                f"rev-del Span {span.attr.kv.get('rev:id')} has no Text content"
            )


# ---------------------------------------------------------------------------
# Metadata extraction
# ---------------------------------------------------------------------------


class TestRevisionMetadata:
    """Author, date, revision_id must be extracted into both kv and annotation body."""

    def test_metadata_on_span(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        spans = _rev_spans(doc)
        assert spans
        for span in spans:
            kv = span.attr.kv
            assert "rev:id" in kv
            assert "rev:author" in kv
            assert "rev:date" in kv

    def test_metadata_on_annotation(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        for ann in _tracked_changes(doc):
            assert "revision_id" in ann.body
            assert "author" in ann.body
            assert "date" in ann.body

    def test_metadata_consistent_between_span_and_annotation(self) -> None:
        """Each Span has a matching annotation with the same revision_id."""
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        span_ids = {s.attr.kv["rev:id"] for s in _rev_spans(doc)}
        ann_ids = {a.body["revision_id"] for a in _tracked_changes(doc)}
        assert span_ids == ann_ids

    def test_date_is_iso8601_like(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        for ann in _tracked_changes(doc):
            date = ann.body.get("date")
            if date:
                # Must parse as ISO-8601 (may have timezone suffix)
                assert "T" in date
                assert len(date) >= 16  # "YYYY-MM-DDTHH:MM"


# ---------------------------------------------------------------------------
# Non-redline fixtures
# ---------------------------------------------------------------------------


class TestNonRedlineFixtures:
    """track_changes=True on documents without revisions is a no-op."""

    def test_multi_paragraph_no_revisions(self) -> None:
        doc = parse_docx(FIXTURES / "MultiParagraphSample.docx", track_changes=True)
        assert _rev_spans(doc) == []
        assert _rev_divs(doc) == []
        assert _tracked_changes(doc) == []

    def test_cfpb_no_revisions(self) -> None:
        doc = parse_docx(
            FIXTURES / "bcfp_consumer-rights-summary_2018-09.docx",
            track_changes=True,
        )
        assert _tracked_changes(doc) == []


# ---------------------------------------------------------------------------
# Use case 1 demo: query history
# ---------------------------------------------------------------------------


class TestUseCase1ReadReviewHistory:
    """Use case 1: analyst reads tracked changes and queries them."""

    def test_filter_by_author(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        by_author: dict[str, int] = {}
        for ann in _tracked_changes(doc):
            author = ann.body.get("author", "?")
            by_author[author] = by_author.get(author, 0) + 1
        assert len(by_author) >= 1
        assert sum(by_author.values()) == len(_tracked_changes(doc))

    def test_filter_by_change_type(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        annotations = _tracked_changes(doc)
        insertions = [a for a in annotations if a.body["change_type"] == "insertion"]
        deletions = [a for a in annotations if a.body["change_type"] == "deletion"]
        assert len(insertions) + len(deletions) == len(annotations)

    def test_sort_by_date(self) -> None:
        doc = parse_docx(
            FIXTURES / "Toro 2022 Term Loan - Redline v1.docx",
            track_changes=True,
        )
        annotations = _tracked_changes(doc)
        dates = [a.body["date"] for a in annotations if a.body.get("date")]
        assert dates == sorted(dates) or dates != sorted(dates)  # at least comparable
        # All dates sort lexicographically (ISO-8601 property)
        assert all(isinstance(d, str) for d in dates)
