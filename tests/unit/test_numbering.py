"""Tests for NumberingResolver: list type detection, overrides."""

from __future__ import annotations

from kaos_office.docx.numbering import NumberingResolver
from kaos_office.ooxml.namespace import W


def _numbering_xml(*defs: str) -> bytes:
    return f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  {"".join(defs)}
</w:numbering>""".encode()


class TestNumberingResolver:
    BASIC = _numbering_xml(
        '<w:abstractNum w:abstractNumId="0">'
        '  <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>'
        '  <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl>'
        "</w:abstractNum>",
        '<w:abstractNum w:abstractNumId="1">'
        '  <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>'
        '  <w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/></w:lvl>'
        "</w:abstractNum>",
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>',
    )

    def test_bullet_not_ordered(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.is_ordered("1", "0") is False

    def test_decimal_is_ordered(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.is_ordered("2", "0") is True

    def test_lower_letter_is_ordered(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.is_ordered("2", "1") is True

    def test_unknown_numid_defaults_to_bullet(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.is_ordered("99", "0") is False

    def test_get_format(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.get_format("1", "0") == "bullet"
        assert resolver.get_format("2", "0") == "decimal"

    def test_has_numbering(self):
        resolver = NumberingResolver.from_xml(self.BASIC)
        assert resolver.has_numbering("1") is True
        assert resolver.has_numbering("99") is False

    def test_none_xml(self):
        resolver = NumberingResolver.from_xml(None)
        assert resolver.is_ordered("1") is False
        assert resolver.has_numbering("1") is False

    def test_level_override(self):
        """Level overrides in num should take precedence over abstractNum."""
        xml = _numbering_xml(
            '<w:abstractNum w:abstractNumId="0">'
            '  <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>'
            "</w:abstractNum>",
            '<w:num w:numId="1">'
            '  <w:abstractNumId w:val="0"/>'
            '  <w:lvlOverride w:ilvl="0">'
            '    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>'
            "  </w:lvlOverride>"
            "</w:num>",
        )
        resolver = NumberingResolver.from_xml(xml)
        # Override changes bullet to decimal
        assert resolver.is_ordered("1", "0") is True
