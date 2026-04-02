"""Tests for StyleResolver: heading detection, inheritance, cycle detection."""

from __future__ import annotations

from kaos_office.docx.styles import StyleResolver
from kaos_office.ooxml.namespace import W


def _styles_xml(*style_defs: str) -> bytes:
    """Build a styles.xml with given style definitions."""
    return f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  {"".join(style_defs)}
</w:styles>""".encode()


def _heading_style(
    style_id: str, name: str, outline_lvl: int | None = None, based_on: str | None = None
) -> str:
    """Build a w:style element."""
    parts = [f'<w:style w:type="paragraph" w:styleId="{style_id}">']
    parts.append(f'  <w:name w:val="{name}"/>')
    if based_on:
        parts.append(f'  <w:basedOn w:val="{based_on}"/>')
    if outline_lvl is not None:
        parts.append(f'  <w:pPr><w:outlineLvl w:val="{outline_lvl}"/></w:pPr>')
    parts.append("</w:style>")
    return "\n".join(parts)


class TestStyleResolver:
    def test_heading_by_outline_level(self):
        xml = _styles_xml(_heading_style("Heading1", "heading 1", outline_lvl=0))
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("Heading1") == 1

    def test_heading_level_2(self):
        xml = _styles_xml(_heading_style("Heading2", "heading 2", outline_lvl=1))
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("Heading2") == 2

    def test_heading_by_name_pattern(self):
        xml = _styles_xml(
            '<w:style w:type="paragraph" w:styleId="H3">  <w:name w:val="heading 3"/></w:style>'
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("H3") == 3

    def test_heading_by_inheritance(self):
        xml = _styles_xml(
            _heading_style("Heading2", "heading 2", outline_lvl=1),
            _heading_style("CustomH", "Custom Heading", based_on="Heading2"),
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("CustomH") == 2

    def test_non_heading_returns_none(self):
        xml = _styles_xml(
            '<w:style w:type="paragraph" w:styleId="Normal">  <w:name w:val="Normal"/></w:style>'
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("Normal") is None

    def test_unknown_style_returns_none(self):
        resolver = StyleResolver.from_xml(None)
        assert resolver.heading_level("NonExistent") is None

    def test_none_style_returns_none(self):
        resolver = StyleResolver.from_xml(None)
        assert resolver.heading_level(None) is None

    def test_cycle_detection(self):
        """Styles that reference each other should not infinite loop."""
        xml = _styles_xml(
            _heading_style("StyleA", "Style A", based_on="StyleB"),
            _heading_style("StyleB", "Style B", based_on="StyleA"),
        )
        resolver = StyleResolver.from_xml(xml)
        # Should return None, not hang
        assert resolver.heading_level("StyleA") is None

    def test_heading_capped_at_6(self):
        xml = _styles_xml(_heading_style("H9", "heading 9", outline_lvl=8))
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("H9") == 6

    def test_heading_name_case_insensitive(self):
        xml = _styles_xml(
            '<w:style w:type="paragraph" w:styleId="H1">  <w:name w:val="HEADING 1"/></w:style>'
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("H1") == 1

    def test_result_cached(self):
        xml = _styles_xml(_heading_style("H1", "heading 1", outline_lvl=0))
        resolver = StyleResolver.from_xml(xml)
        r1 = resolver.heading_level("H1")
        r2 = resolver.heading_level("H1")
        assert r1 == r2 == 1

    def test_is_code_style(self):
        xml = _styles_xml(
            '<w:style w:type="paragraph" w:styleId="Code">  <w:name w:val="Code"/></w:style>'
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.is_code_style("Code") is True
        assert resolver.is_code_style("Normal") is False

    def test_toc_heading_not_heading(self):
        xml = _styles_xml(
            '<w:style w:type="paragraph" w:styleId="TOCHeading">'
            '  <w:name w:val="TOC Heading"/>'
            "</w:style>"
        )
        resolver = StyleResolver.from_xml(xml)
        assert resolver.heading_level("TOCHeading") is None

    def test_has_style(self):
        xml = _styles_xml(_heading_style("H1", "heading 1", outline_lvl=0))
        resolver = StyleResolver.from_xml(xml)
        assert resolver.has_style("H1")
        assert not resolver.has_style("Missing")
