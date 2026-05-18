"""Tests for the DOCX reader's numbering integration.

Validates that the rendered numbering label travels from
``numbering.xml`` through ``NumberingState`` and lands on the AST
node's ``numbering_label`` field. Builds minimal in-memory DOCX
packages so the test stays fast and hermetic; the heavier
"real-document" fixtures are exercised in Stage 5.
"""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

from kaos_content.model.blocks import Heading, ListItem, OrderedList, Paragraph

from kaos_office.docx.reader import parse_docx
from kaos_office.ooxml.namespace import W

# XML fixtures intentionally use long URIs (OOXML / package
# relationship namespaces) — line-breaking them would corrupt the
# fixtures, so the long-line lint is suppressed on these constants only.
_CT = """\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""  # noqa: E501

_PKG_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""  # noqa: E501

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""  # noqa: E501

_STYLES = f"""\
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
</w:styles>
"""


def _docx_bytes(document_xml: str, numbering_xml: str, styles_xml: str = _STYLES) -> bytes:
    """Build a minimal valid DOCX archive in memory."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _PKG_RELS)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/numbering.xml", numbering_xml)
        zf.writestr("word/styles.xml", styles_xml)
    return buf.getvalue()


def _write_docx(tmp_path: Path, name: str, document_xml: str, numbering_xml: str) -> Path:
    path = tmp_path / name
    path.write_bytes(_docx_bytes(document_xml, numbering_xml))
    return path


# Three-level decimal / lowerLetter / lowerRoman pattern, the legal
# section pattern ``Section 11(a)(i)``.
_LEGAL_NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="(%2)"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="(%3)"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""


class TestListItemNumberingLabel:
    """Numbered list items carry the rendered label on the AST."""

    DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>First top-level item.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Second top-level item.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Nested clause one.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Nested clause two.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Sub-sub clause.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

    def test_first_top_level(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        top = doc.body[0]
        assert isinstance(top, OrderedList)
        first_item = top.children[0]
        assert isinstance(first_item, ListItem)
        assert first_item.numbering_label == "1."

    def test_second_top_level(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        top = doc.body[0]
        assert isinstance(top, OrderedList)
        second_item = top.children[1]
        assert isinstance(second_item, ListItem)
        assert second_item.numbering_label == "2."

    def test_nested_first(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        top = doc.body[0]
        assert isinstance(top, OrderedList)
        # Locate nested clause one inside item 2.
        outer_second = top.children[1]
        assert isinstance(outer_second, ListItem)
        # Children: paragraph + nested OrderedList
        nested = None
        for child in outer_second.children:
            if isinstance(child, OrderedList):
                nested = child
                break
        assert nested is not None, "expected nested OrderedList under item 2"
        first_clause = nested.children[0]
        assert isinstance(first_clause, ListItem)
        assert first_clause.numbering_label == "(a)"

    def test_sub_sub_clause(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        labels = _collect_list_labels(doc.body)
        assert labels == ["1.", "2.", "(a)", "(b)", "(i)"]


class TestHeadingNumberingLabel:
    """Headings with auto-numbering pick up the rendered label."""

    DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>GOVERNING LAW</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>ASSIGNMENT</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

    def test_heading_carries_label(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        first_heading = doc.body[0]
        second_heading = doc.body[1]
        assert isinstance(first_heading, Heading)
        assert isinstance(second_heading, Heading)
        assert first_heading.numbering_label == "1."
        assert second_heading.numbering_label == "2."

    def test_heading_without_numbering_has_no_label(self, tmp_path: Path) -> None:
        doc_xml = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Recitals</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
        path = _write_docx(tmp_path, "doc.docx", doc_xml, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        heading = doc.body[0]
        assert isinstance(heading, Heading)
        assert heading.numbering_label is None


class TestUnnumberedParagraphStaysNone:
    """A paragraph with no ``numPr`` has ``numbering_label is None``."""

    def test_plain_paragraph(self, tmp_path: Path) -> None:
        doc_xml = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p><w:r><w:t>Just a sentence.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""
        path = _write_docx(tmp_path, "doc.docx", doc_xml, _LEGAL_NUMBERING)
        doc = parse_docx(path)
        p = doc.body[0]
        assert isinstance(p, Paragraph)
        assert p.numbering_label is None


class TestMixedHeadingAndListSequence:
    """Headings and list items in the same document advance independent
    counters when they reference different ``numId`` instances, and
    share counter state when they reference the same instance."""

    NUMBERING_TWO_INSTANCES = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="Section %1."/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="(%1)"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
"""

    DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>GOVERNING LAW</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>First sub-clause.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>Second sub-clause.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>ASSIGNMENT</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

    def test_heading_counter_advances_independently(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, self.NUMBERING_TWO_INSTANCES)
        doc = parse_docx(path)
        headings = [b for b in doc.body if isinstance(b, Heading)]
        assert headings[0].numbering_label == "Section 1."
        assert headings[1].numbering_label == "Section 2."

    def test_sublist_counter_starts_at_a(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, "doc.docx", self.DOCUMENT, self.NUMBERING_TWO_INSTANCES)
        doc = parse_docx(path)
        ordered_list = next(b for b in doc.body if isinstance(b, OrderedList))
        first_item = ordered_list.children[0]
        assert isinstance(first_item, ListItem)
        assert first_item.numbering_label == "(a)"


def _collect_list_labels(blocks: tuple) -> list[str]:
    """Walk blocks recursively and collect ``ListItem.numbering_label`` values."""
    labels: list[str] = []
    for block in blocks:
        if isinstance(block, ListItem):
            if block.numbering_label is not None:
                labels.append(block.numbering_label)
            labels.extend(_collect_list_labels(block.children))
        elif hasattr(block, "children"):
            labels.extend(_collect_list_labels(block.children))
    return labels
