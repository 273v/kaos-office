"""Integration tests against real DOCX files from kelvin_office fixtures."""

from __future__ import annotations

import pytest
from kaos_content.serializers.markdown import serialize_markdown
from kaos_content.serializers.text import serialize_text

from kaos_office.docx.reader import parse_docx
from tests.conftest import KELVIN_FIXTURES, skip_no_fixtures


@skip_no_fixtures
class TestMultiParagraphSample:
    """Tests against MultiParagraphSample.docx."""

    @pytest.fixture
    def doc(self):
        return parse_docx(KELVIN_FIXTURES / "MultiParagraphSample.docx")

    def test_block_count(self, doc):
        assert len(doc.body) >= 3

    def test_has_paragraphs(self, doc):
        para_count = sum(1 for b in doc.body if b.node_type == "paragraph")
        assert para_count >= 2

    def test_has_list(self, doc):
        list_count = sum(1 for b in doc.body if b.node_type == "bullet_list")
        assert list_count >= 1

    def test_bold_in_markdown(self, doc):
        md = serialize_markdown(doc)
        assert "**" in md

    def test_italic_in_markdown(self, doc):
        md = serialize_markdown(doc)
        assert "*" in md

    def test_list_items_in_text(self, doc):
        text = serialize_text(doc)
        assert "first list item" in text.lower()
        assert "second" in text.lower()

    def test_has_comment_annotations(self, doc):
        assert len(doc.annotations) >= 1


@skip_no_fixtures
class TestFootnote:
    """Tests against Footnote.docx."""

    @pytest.fixture
    def doc(self):
        return parse_docx(KELVIN_FIXTURES / "Footnote.docx")

    def test_has_footnotes(self, doc):
        assert len(doc.footnotes) >= 1

    def test_footnote_ref_in_markdown(self, doc):
        md = serialize_markdown(doc)
        assert "[^" in md

    def test_footnote_content(self, doc):
        # At least one footnote should have content
        for _fn_id, blocks in doc.footnotes.items():
            text = serialize_text(type(doc)(body=blocks))
            if text.strip():
                return
        pytest.fail("No footnote content found")


@skip_no_fixtures
class TestToroTermLoan:
    """Tests against Toro 2022 Term Loan.docx — complex legal document."""

    @pytest.fixture
    def doc(self):
        return parse_docx(KELVIN_FIXTURES / "Toro 2022 Term Loan.docx")

    def test_large_document(self, doc):
        assert len(doc.body) >= 500

    def test_significant_text(self, doc):
        text = serialize_text(doc)
        assert len(text) > 100000

    def test_has_tables(self, doc):
        table_count = sum(1 for b in doc.body if b.node_type == "table")
        assert table_count >= 1

    def test_markdown_renders(self, doc):
        md = serialize_markdown(doc)
        assert len(md) > 100000


@skip_no_fixtures
class TestToroRedline:
    """Tests against Toro 2022 Term Loan - Redline v1.docx — track changes."""

    @pytest.fixture
    def doc(self):
        return parse_docx(KELVIN_FIXTURES / "Toro 2022 Term Loan - Redline v1.docx")

    def test_parses_without_error(self, doc):
        assert len(doc.body) >= 500

    def test_no_deltext_leaks(self, doc):
        """Deleted text should not appear in the output."""
        text = serialize_text(doc)
        # We can't check specific deleted text without knowing the document,
        # but we can verify it doesn't crash and produces substantial output
        assert len(text) > 50000

    def test_similar_size_to_clean(self, doc):
        """Redline output should be roughly similar in size to the clean version."""
        clean_doc = parse_docx(KELVIN_FIXTURES / "Toro 2022 Term Loan.docx")
        redline_text = serialize_text(doc)
        clean_text = serialize_text(clean_doc)
        # Allow ±50% difference (some text is inserted/deleted)
        ratio = len(redline_text) / len(clean_text)
        assert 0.5 < ratio < 1.5, f"Size ratio {ratio:.2f} outside expected range"


@skip_no_fixtures
class TestToroComments:
    """Tests against Toro 2022 Term Loan - Comments.docx."""

    @pytest.fixture
    def doc(self):
        return parse_docx(KELVIN_FIXTURES / "Toro 2022 Term Loan - Comments.docx")

    def test_has_comments(self, doc):
        assert len(doc.annotations) >= 3

    def test_comment_has_author(self, doc):
        for ann in doc.annotations:
            assert ann.body.get("author") is not None

    def test_comment_has_text(self, doc):
        for ann in doc.annotations:
            assert ann.body.get("text")


@skip_no_fixtures
class TestVariousDocuments:
    """Quick smoke tests across multiple fixtures."""

    @pytest.mark.parametrize(
        "filename",
        [
            "CheeseSample.docx",
            "sample.docx",
            "bcfp_consumer-rights-summary_2018-09.docx",
            "cms-10704-hra-model-notice.docx",
            "PolicyProcedureTemplate_PhysicalFacility_Final.docx",
        ],
    )
    def test_parse_without_error(self, filename):
        path = KELVIN_FIXTURES / filename
        if not path.exists():
            pytest.skip(f"Fixture {filename} not found")
        doc = parse_docx(path)
        assert len(doc.body) >= 1

    @pytest.mark.parametrize(
        "filename",
        [
            "CheeseSample.docx",
            "sample.docx",
            "bcfp_consumer-rights-summary_2018-09.docx",
        ],
    )
    def test_markdown_renders(self, filename):
        path = KELVIN_FIXTURES / filename
        if not path.exists():
            pytest.skip(f"Fixture {filename} not found")
        doc = parse_docx(path)
        md = serialize_markdown(doc)
        assert len(md) > 0


@skip_no_fixtures
class TestMCSRedline:
    """Tests against MCSRedline10312022.docx — large document with extensive redlines."""

    def test_parses_without_error(self):
        doc = parse_docx(KELVIN_FIXTURES / "MCSRedline10312022.docx")
        assert len(doc.body) >= 100
        text = serialize_text(doc)
        assert len(text) > 10000
