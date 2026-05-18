"""Tests for Stage 7 extensions: pStyle-linked numbering + international formats."""

# ruff: noqa: RUF001
# Intentional non-Latin code-points populate the parametrize tables
# below (Hebrew / Arabic / Chinese / katakana).

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from kaos_content.model.blocks import Heading

from kaos_office.docx.numbering import (
    format_aiueo,
    format_arabic_alpha,
    format_chinese_counting,
    format_hebrew_1,
    format_iroha,
    format_number,
    parse_numbering_xml,
)
from kaos_office.docx.reader import parse_docx
from kaos_office.ooxml.namespace import W


class TestHebrewFormat:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "א"), (2, "ב"), (10, "י"), (22, "ת"), (23, "א")],
    )
    def test_values(self, value: int, expected: str) -> None:
        assert format_hebrew_1(value) == expected


class TestArabicAlphaFormat:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "أ"), (2, "ب"), (28, "ي"), (29, "أ")],
    )
    def test_values(self, value: int, expected: str) -> None:
        assert format_arabic_alpha(value) == expected


class TestChineseCounting:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (1, "一"),
            (11, "一一"),
            (23, "二三"),
            (100, "一〇〇"),
        ],
    )
    def test_values(self, value: int, expected: str) -> None:
        assert format_chinese_counting(value) == expected


class TestKatakanaAiueo:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "ア"), (2, "イ"), (5, "オ"), (20, "ト")],
    )
    def test_values(self, value: int, expected: str) -> None:
        assert format_aiueo(value) == expected


class TestKatakanaIroha:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [(1, "イ"), (2, "ロ"), (3, "ハ")],
    )
    def test_values(self, value: int, expected: str) -> None:
        assert format_iroha(value) == expected


class TestInternationalFormatDispatch:
    """The format_number dispatch picks up international converters."""

    @pytest.mark.parametrize(
        ("num_fmt", "value", "expected"),
        [
            ("hebrew1", 5, "ה"),
            ("arabicAlpha", 3, "ت"),
            ("chineseCounting", 7, "七"),
            ("aiueo", 6, "カ"),
            ("iroha", 4, "ニ"),
        ],
    )
    def test_dispatch(self, num_fmt: str, value: int, expected: str) -> None:
        assert format_number(value, num_fmt) == expected


# ── pStyle-linked numbering integration test ──────────────────────────


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


_STYLES_WITH_HEADING1 = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
</w:styles>
"""

# Numbering definition whose level 0 is style-linked to Heading1 —
# any paragraph whose pStyle is Heading1 inherits this numbering even
# without inline numPr.
_PSTYLE_NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="Article %1."/>
      <w:pStyle w:val="Heading1"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""

_PSTYLE_DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <!-- Headings carry pStyle=Heading1 but no inline numPr -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>SCOPE</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>DEFINITIONS</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""


def _write_docx(tmp_path: Path, document_xml: str, numbering_xml: str, styles: str) -> Path:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _PKG_RELS)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/numbering.xml", numbering_xml)
        zf.writestr("word/styles.xml", styles)
    path = tmp_path / "doc.docx"
    path.write_bytes(buf.getvalue())
    return path


def _has_numbering_label_support() -> bool:
    try:
        from kaos_content import Paragraph

        return "numbering_label" in Paragraph.model_fields  # type: ignore[attr-defined]
    except ImportError:
        return False


requires_label = pytest.mark.skipif(
    not _has_numbering_label_support(),
    reason="kaos-content release with numbering_label not yet installed",
)


class TestPStyleLinkedNumbering:
    """A paragraph that inherits numbering through its style — no inline
    numPr — still picks up the rendered label."""

    def test_resolve_pstyle_in_definitions(self) -> None:
        defs = parse_numbering_xml(_PSTYLE_NUMBERING.encode())
        resolved = defs.resolve_pstyle("Heading1")
        assert resolved == ("1", 0)

    def test_resolve_unknown_pstyle_returns_none(self) -> None:
        defs = parse_numbering_xml(_PSTYLE_NUMBERING.encode())
        assert defs.resolve_pstyle("UnknownStyle") is None

    @requires_label
    def test_heading_inherits_pstyle_numbering(self, tmp_path: Path) -> None:
        path = _write_docx(tmp_path, _PSTYLE_DOCUMENT, _PSTYLE_NUMBERING, _STYLES_WITH_HEADING1)
        doc = parse_docx(path)
        headings = [b for b in doc.body if isinstance(b, Heading)]
        assert headings[0].numbering_label == "Article 1."  # ty: ignore[unresolved-attribute]
        assert headings[1].numbering_label == "Article 2."  # ty: ignore[unresolved-attribute]
