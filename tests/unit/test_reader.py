"""Tests for DOCX reader: paragraphs, headings, lists, tables, formatting, track changes."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from kaos_content.model.blocks import Heading, Table
from kaos_content.serializers.markdown import serialize_markdown
from kaos_content.serializers.text import serialize_text

from tests.conftest import make_minimal_docx

if TYPE_CHECKING:
    from kaos_content import ContentDocument

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _parse_from_body(body_xml: str, **kwargs) -> ContentDocument:
    """Helper: create a DOCX from body XML and parse it."""
    from kaos_office.docx.reader import parse_docx

    docx_bytes = make_minimal_docx(body_xml=body_xml, **kwargs)
    tmp = Path("/tmp/test_reader.docx")
    tmp.write_bytes(docx_bytes)
    try:
        return parse_docx(tmp)
    finally:
        tmp.unlink(missing_ok=True)


# ──────────────────────────── Paragraphs ────────────────────────────


class TestParagraphs:
    def test_single_paragraph(self):
        doc = _parse_from_body("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        assert len(doc.body) == 1
        assert doc.body[0].node_type == "paragraph"

    def test_multiple_paragraphs(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>"
        )
        assert len(doc.body) == 2

    def test_empty_paragraph_skipped(self):
        doc = _parse_from_body("<w:p></w:p><w:p><w:r><w:t>Content</w:t></w:r></w:p>")
        assert len(doc.body) == 1

    def test_whitespace_only_paragraph(self):
        doc = _parse_from_body('<w:p><w:r><w:t xml:space="preserve">   </w:t></w:r></w:p>')
        # Whitespace-only still creates a paragraph (it has content)
        text = serialize_text(doc)
        assert text.strip() == ""

    def test_paragraph_provenance(self):
        doc = _parse_from_body("<w:p><w:r><w:t>Test</w:t></w:r></w:p>")
        assert doc.body[0].provenance is not None
        assert doc.body[0].provenance.extractor == "kaos-office/docx"


# ──────────────────────────── Formatting ────────────────────────────


class TestFormatting:
    def test_bold(self):
        doc = _parse_from_body("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p>")
        md = serialize_markdown(doc)
        assert "**Bold**" in md

    def test_italic(self):
        doc = _parse_from_body("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r></w:p>")
        md = serialize_markdown(doc)
        assert "*Italic*" in md

    def test_bold_italic(self):
        doc = _parse_from_body("<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Both</w:t></w:r></w:p>")
        md = serialize_markdown(doc)
        assert "***Both***" in md or "**_Both_**" in md or "*__Both__*" in md

    def test_strikethrough(self):
        doc = _parse_from_body("<w:p><w:r><w:rPr><w:strike/></w:rPr><w:t>Struck</w:t></w:r></w:p>")
        md = serialize_markdown(doc)
        assert "~~Struck~~" in md

    def test_bold_val_false_not_bold(self):
        doc = _parse_from_body(
            '<w:p><w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>Not bold</w:t></w:r></w:p>'
        )
        md = serialize_markdown(doc)
        assert "**" not in md

    def test_tab_preserved(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>After</w:t></w:r></w:p>"
        )
        text = serialize_text(doc)
        assert "\t" in text

    def test_line_break(self):
        doc = _parse_from_body("<w:p><w:r><w:t>Line 1</w:t><w:br/><w:t>Line 2</w:t></w:r></w:p>")
        text = serialize_text(doc)
        assert "Line 1" in text
        assert "Line 2" in text

    def test_adjacent_text_merged(self):
        doc = _parse_from_body("<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>")
        text = serialize_text(doc)
        assert "Hello World" in text


# ──────────────────────────── Headings ────────────────────────────


class TestHeadings:
    STYLES_XML = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="SubHeading">
    <w:name w:val="Sub Heading"/>
    <w:basedOn w:val="Heading2"/>
  </w:style>
</w:styles>"""

    def test_heading_by_style(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
            styles_xml=self.STYLES_XML,
        )
        heading = doc.body[0]
        assert isinstance(heading, Heading)
        assert heading.depth == 1

    def test_heading_level2(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Section</w:t></w:r></w:p>',
            styles_xml=self.STYLES_XML,
        )
        heading = doc.body[0]
        assert isinstance(heading, Heading)
        assert heading.depth == 2

    def test_heading_by_inheritance(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:pStyle w:val="SubHeading"/></w:pPr>'
            "<w:r><w:t>Inherited</w:t></w:r></w:p>",
            styles_xml=self.STYLES_XML,
        )
        heading = doc.body[0]
        assert isinstance(heading, Heading)
        assert heading.depth == 2

    def test_heading_in_markdown(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>My Title</w:t></w:r></w:p>',
            styles_xml=self.STYLES_XML,
        )
        md = serialize_markdown(doc)
        assert md.strip() == "# My Title"


# ──────────────────────────── Lists ────────────────────────────


class TestLists:
    NUMBERING_XML = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"""

    def test_bullet_list(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>Item 1</w:t></w:r></w:p>"
            '<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>Item 2</w:t></w:r></w:p>",
            numbering_xml=self.NUMBERING_XML,
        )
        md = serialize_markdown(doc)
        assert "- Item 1" in md
        assert "- Item 2" in md

    def test_ordered_list(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:numPr><w:numId w:val="2"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>First</w:t></w:r></w:p>"
            '<w:p><w:pPr><w:numPr><w:numId w:val="2"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>Second</w:t></w:r></w:p>",
            numbering_xml=self.NUMBERING_XML,
        )
        md = serialize_markdown(doc)
        assert "1." in md or "1. " in md

    def test_list_followed_by_paragraph(self):
        doc = _parse_from_body(
            '<w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>Item</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Normal para</w:t></w:r></w:p>",
            numbering_xml=self.NUMBERING_XML,
        )
        # Should have a list then a paragraph
        assert doc.body[0].node_type == "bullet_list"
        assert doc.body[1].node_type == "paragraph"

    def test_numid_zero_not_list(self):
        """numId=0 means 'no list' — paragraph should not be in a list."""
        doc = _parse_from_body(
            '<w:p><w:pPr><w:numPr><w:numId w:val="0"/><w:ilvl w:val="0"/></w:numPr></w:pPr>'
            "<w:r><w:t>Not a list</w:t></w:r></w:p>",
            numbering_xml=self.NUMBERING_XML,
        )
        assert doc.body[0].node_type == "paragraph"


# ──────────────────────────── Tables ────────────────────────────


class TestTables:
    def test_simple_table(self):
        doc = _parse_from_body(
            "<w:tbl>"
            "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>"
            "<w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>"
        )
        table = doc.body[0]
        assert isinstance(table, Table)
        assert len(table.bodies) == 1
        assert len(table.bodies[0].rows) == 2

    def test_table_with_merged_cells(self):
        doc = _parse_from_body(
            "<w:tbl>"
            '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
            "<w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc></w:tr>"
            "<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>"
        )
        table = doc.body[0]
        assert isinstance(table, Table)
        first_row = table.bodies[0].rows[0]
        assert first_row.cells[0].col_span == 2

    def test_table_in_markdown(self):
        doc = _parse_from_body(
            "<w:tbl>"
            "<w:tr><w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>"
            "</w:tbl>"
        )
        md = serialize_markdown(doc)
        assert "Name" in md
        assert "Value" in md


# ──────────────────────────── Track Changes ────────────────────────────


class TestTrackChanges:
    def test_insertion_included(self):
        doc = _parse_from_body("<w:p><w:ins><w:r><w:t>Inserted text</w:t></w:r></w:ins></w:p>")
        text = serialize_text(doc)
        assert "Inserted text" in text

    def test_deletion_skipped(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>Keep</w:t></w:r>"
            "<w:del><w:r><w:delText>Deleted</w:delText></w:r></w:del></w:p>"
        )
        text = serialize_text(doc)
        assert "Keep" in text
        assert "Deleted" not in text

    def test_mixed_track_changes(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>Original </w:t></w:r>"
            "<w:del><w:r><w:delText>old</w:delText></w:r></w:del>"
            "<w:ins><w:r><w:t>new</w:t></w:r></w:ins>"
            "<w:r><w:t> text</w:t></w:r></w:p>"
        )
        text = serialize_text(doc)
        assert "Original" in text
        assert "new" in text
        assert "old" not in text

    def test_body_level_insertion(self):
        doc = _parse_from_body("<w:ins><w:p><w:r><w:t>New paragraph</w:t></w:r></w:p></w:ins>")
        text = serialize_text(doc)
        assert "New paragraph" in text

    def test_body_level_deletion(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>Keep</w:t></w:r></w:p>"
            "<w:del><w:p><w:r><w:delText>Remove</w:delText></w:r></w:p></w:del>"
        )
        text = serialize_text(doc)
        assert "Keep" in text
        # Deleted body-level paragraph should be skipped
        assert "Remove" not in text


# ──────────────────────────── SDT (Structured Document Tags) ────────────────────────────


class TestSDT:
    def test_sdt_unwrapped(self):
        doc = _parse_from_body(
            "<w:sdt><w:sdtContent>"
            "<w:p><w:r><w:t>Content inside SDT</w:t></w:r></w:p>"
            "</w:sdtContent></w:sdt>"
        )
        text = serialize_text(doc)
        assert "Content inside SDT" in text


# ──────────────────────────── Metadata ────────────────────────────


class TestMetadata:
    def test_metadata_from_core_xml(self):
        core = """\
<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
  <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
</cp:coreProperties>"""

        doc = _parse_from_body(
            "<w:p><w:r><w:t>Content</w:t></w:r></w:p>",
            core_xml=core,
        )
        assert doc.metadata.title == "Test Document"
        assert doc.metadata.authors == ("Test Author",)

    def test_metadata_from_app_xml(self):
        from kaos_office.docx.metadata import DocxMetadata

        app = """\
<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Words>1000</Words>
  <Pages>5</Pages>
  <Company>Test Corp</Company>
</Properties>"""
        meta = DocxMetadata.from_xml(app_xml=app.encode("utf-8"))
        assert meta.word_count == 1000
        assert meta.page_count == 5
        assert meta.company == "Test Corp"


# ──────────────────────────── Footnotes ────────────────────────────


class TestFootnotes:
    FOOTNOTES_XML = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="{W}">
  <w:footnote w:type="separator" w:id="0"/>
  <w:footnote w:type="continuationSeparator" w:id="-1"/>
  <w:footnote w:id="1">
    <w:p><w:r><w:t>This is footnote content.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>"""

    def test_footnote_extracted(self):
        doc = _parse_from_body(
            "<w:p><w:r><w:t>See note</w:t></w:r>"
            '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
            '<w:footnoteReference w:id="1"/></w:r></w:p>',
            footnotes_xml=self.FOOTNOTES_XML,
        )
        assert "1" in doc.footnotes
        md = serialize_markdown(doc)
        assert "footnote" in md.lower() or "[^1]" in md


# ──────────────────────────── Comments ────────────────────────────


class TestComments:
    COMMENTS_XML = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="{W}">
  <w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">
    <w:p><w:r><w:t>Please review this section.</w:t></w:r></w:p>
  </w:comment>
</w:comments>"""

    def test_comments_as_annotations(self):
        doc = _parse_from_body(
            '<w:p><w:commentRangeStart w:id="1"/>'
            "<w:r><w:t>Commented text</w:t></w:r>"
            '<w:commentRangeEnd w:id="1"/></w:p>',
            comments_xml=self.COMMENTS_XML,
        )
        assert len(doc.annotations) == 1
        assert doc.annotations[0].body["text"] == "Please review this section."
        assert doc.annotations[0].body["author"] == "Reviewer"
