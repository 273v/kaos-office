"""Unit tests for PPTX reader — synthetic PPTX files."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from tests.conftest import make_minimal_pptx

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _parse(pptx_bytes: bytes):
    """Helper: write bytes to temp file, parse, return doc."""
    from kaos_office.pptx.reader import parse_pptx

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
        f.write(pptx_bytes)
        f.flush()
        return parse_pptx(f.name)


class TestTitleExtraction:
    """Test title placeholder → heading mapping."""

    def test_title_becomes_heading1(self):
        data = make_minimal_pptx()
        doc = _parse(data)
        # Should have 1 slide div containing a heading
        assert len(doc.body) == 1
        div = doc.body[0]
        assert div.node_type == "div"
        # Find heading in children
        headings = [c for c in div.children if c.node_type == "heading"]
        assert len(headings) == 1
        assert headings[0].depth == 1

    def test_center_title_becomes_heading1(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>Centered Title</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        headings = [c for c in div.children if c.node_type == "heading"]
        assert len(headings) == 1
        assert headings[0].depth == 1

    def test_subtitle_becomes_heading2(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Subtitle 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="subTitle"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>My Subtitle</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        headings = [c for c in div.children if c.node_type == "heading"]
        assert len(headings) == 1
        assert headings[0].depth == 2

    def test_empty_title_skipped(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t></a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        headings = [c for c in div.children if c.node_type == "heading"]
        assert len(headings) == 0


class TestSkipPlaceholders:
    """Test that metadata placeholders are skipped."""

    def test_date_skipped(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Date 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="dt"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>01/01/2025</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        assert len(div.children) == 0

    def test_slide_number_skipped(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="SlideNum 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="sldNum"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>42</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        assert len(div.children) == 0

    def test_footer_skipped(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Footer 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="ftr"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>Footer text</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        assert len(div.children) == 0


class TestBodyText:
    """Test body text extraction with formatting."""

    def test_plain_text(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buNone/></a:pPr>
      <a:r><a:t>Hello world</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        paras = [c for c in div.children if c.node_type == "paragraph"]
        assert len(paras) == 1

    def test_bold_text(self):
        from kaos_content.serializers.markdown import serialize_markdown

        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buNone/></a:pPr>
      <a:r><a:rPr b="1"/><a:t>Bold text</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        md = serialize_markdown(doc)
        assert "**Bold text**" in md

    def test_italic_text(self):
        from kaos_content.serializers.markdown import serialize_markdown

        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buNone/></a:pPr>
      <a:r><a:rPr i="1"/><a:t>Italic text</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        md = serialize_markdown(doc)
        assert "*Italic text*" in md

    def test_bold_italic_text(self):
        from kaos_content.serializers.markdown import serialize_markdown

        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buNone/></a:pPr>
      <a:r><a:rPr b="1" i="1"/><a:t>Bold italic</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        md = serialize_markdown(doc)
        # Serializer may use ***...*** or **_..._** — both valid
        assert "Bold italic" in md
        assert "**" in md  # Has bold markers


class TestBulletLists:
    """Test bullet and list detection."""

    def test_bullet_list(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buChar char="&#x2022;"/></a:pPr>
      <a:r><a:t>First item</a:t></a:r>
    </a:p>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buChar char="&#x2022;"/></a:pPr>
      <a:r><a:t>Second item</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        lists = [c for c in div.children if c.node_type == "bullet_list"]
        assert len(lists) == 1
        assert len(lists[0].children) == 2  # 2 items

    def test_ordered_list(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr>
      <a:r><a:t>Step one</a:t></a:r>
    </a:p>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buAutoNum type="arabicPeriod"/></a:pPr>
      <a:r><a:t>Step two</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        lists = [c for c in div.children if c.node_type == "ordered_list"]
        assert len(lists) == 1
        assert len(lists[0].children) == 2

    def test_nested_bullets(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buChar char="&#x2022;"/></a:pPr>
      <a:r><a:t>Parent</a:t></a:r>
    </a:p>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="1"><a:buChar char="-"/></a:pPr>
      <a:r><a:t>Child</a:t></a:r>
    </a:p>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buChar char="&#x2022;"/></a:pPr>
      <a:r><a:t>Another parent</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(doc)
        assert "Parent" in md
        assert "Child" in md
        assert "Another parent" in md

    def test_buNone_no_list(self):
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}">
      <a:pPr lvl="0"><a:buNone/></a:pPr>
      <a:r><a:t>No bullet here</a:t></a:r>
    </a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        div = doc.body[0]
        lists = [c for c in div.children if "list" in c.node_type]
        assert len(lists) == 0
        paras = [c for c in div.children if c.node_type == "paragraph"]
        assert len(paras) == 1


class TestTableExtraction:
    """Test table processing."""

    def test_simple_table(self):
        """Test basic table extraction via real stress fixture."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "table_test2.pptx"
        if not fixture.exists():
            pytest.skip("table_test2.pptx fixture not available")
        doc = parse_pptx(fixture)

        # Find tables in the document
        def find_tables(blocks):
            tables = []
            for b in blocks:
                if b.node_type == "table":
                    tables.append(b)
                if hasattr(b, "children"):
                    tables.extend(find_tables(b.children))
            return tables

        tables = find_tables(doc.body)
        assert len(tables) >= 1


class TestChartExtraction:
    """Test chart linearization."""

    def test_bar_chart(self):
        """Test chart extraction from real fixture."""
        from kaos_content.serializers.markdown import serialize_markdown

        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "bar-chart.pptx"
        if not fixture.exists():
            pytest.skip("bar-chart.pptx fixture not available")
        doc = parse_pptx(fixture)
        md = serialize_markdown(doc)
        # Chart should be linearized as a table with category/series columns
        assert "1st Qtr" in md
        assert "2nd Qtr" in md

    def test_pie_chart(self):
        """Test pie chart extraction from real fixture."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "pie-chart.pptx"
        if not fixture.exists():
            pytest.skip("pie-chart.pptx fixture not available")
        doc = parse_pptx(fixture)

        def find_tables(blocks):
            tables = []
            for b in blocks:
                if b.node_type == "table":
                    tables.append(b)
                if hasattr(b, "children"):
                    tables.extend(find_tables(b.children))
            return tables

        tables = find_tables(doc.body)
        assert len(tables) >= 1


class TestSpatialOrdering:
    """Test shape sorting by position."""

    def test_shapes_sorted_by_top_then_left(self):
        # Two body placeholders: second has a lower top value, should be sorted first
        slide_xml = f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Body 1"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr><a:xfrm xmlns:a="{A_NS}"><a:off x="100" y="2000000"/><a:ext cx="5000000" cy="500000"/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:pPr lvl="0"><a:buNone/></a:pPr><a:r><a:t>Lower shape</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="3" name="Body 2"/>
    <p:cNvSpPr><a:spLocks xmlns:a="{A_NS}" noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="body" idx="2"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr><a:xfrm xmlns:a="{A_NS}"><a:off x="100" y="100000"/><a:ext cx="5000000" cy="500000"/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:pPr lvl="0"><a:buNone/></a:pPr><a:r><a:t>Upper shape</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        from kaos_content.serializers.text import serialize_text

        text = serialize_text(doc)
        # Upper shape should come before lower shape
        upper_pos = text.index("Upper shape")
        lower_pos = text.index("Lower shape")
        assert upper_pos < lower_pos


class TestMultipleSlides:
    """Test multi-slide presentations."""

    def test_two_slides(self):
        slide1 = f"""<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>Slide One</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        slide2 = f"""<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>Slide Two</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide1, slide2]))
        assert len(doc.body) == 2
        assert doc.body[0].attr.kv.get("slide_number") == "1"
        assert doc.body[1].attr.kv.get("slide_number") == "2"

    def test_empty_slide(self):
        # Slide with no content shapes
        slide_xml = ""
        doc = _parse(make_minimal_pptx(slide_xmls=[slide_xml]))
        assert len(doc.body) == 1
        div = doc.body[0]
        assert len(div.children) == 0


class TestMetadata:
    """Test metadata extraction."""

    def test_title_from_core_xml(self):
        core = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>My Presentation</dc:title>
  <dc:creator>Test Author</dc:creator>
</cp:coreProperties>"""
        doc = _parse(make_minimal_pptx(core_xml=core))
        assert doc.metadata.title == "My Presentation"

    def test_no_core_xml(self):
        doc = _parse(make_minimal_pptx(core_xml=None))
        # Should not crash
        assert doc is not None


class TestEdgeCases:
    """Test edge cases and error handling."""

    def test_no_slides(self):
        """Test PPTX with no slides."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "no-slides.pptx"
        if not fixture.exists():
            pytest.skip("no-slides.pptx fixture not available")
        doc = parse_pptx(fixture)
        assert len(doc.body) == 0

    def test_minimal_pptx(self):
        """Test minimal valid PPTX."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "minimal.pptx"
        if not fixture.exists():
            pytest.skip("minimal.pptx fixture not available")
        doc = parse_pptx(fixture)
        assert doc is not None

    def test_no_core_props(self):
        """Test PPTX without core.xml properties."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = (
            Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "no-core-props.pptx"
        )
        if not fixture.exists():
            pytest.skip("no-core-props.pptx fixture not available")
        doc = parse_pptx(fixture)
        assert doc is not None

    def test_missing_rels(self):
        """Test PPTX with missing relationships."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = (
            Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "missing_rels_item.pptx"
        )
        if not fixture.exists():
            pytest.skip("missing_rels_item.pptx fixture not available")
        doc = parse_pptx(fixture)
        assert doc is not None


class TestHelpers:
    """Test public helper functions."""

    def test_get_slide_count(self):
        import tempfile

        from kaos_office.pptx.reader import get_slide_count

        data = make_minimal_pptx(slide_xmls=["", ""])
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            f.write(data)
            f.flush()
            assert get_slide_count(f.name) == 2

    def test_get_slide_text(self):
        import tempfile

        from kaos_office.pptx.reader import get_slide_text

        data = make_minimal_pptx()
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            f.write(data)
            f.flush()
            text = get_slide_text(f.name, 1)
            assert "Test Title" in text

    def test_get_slide_text_out_of_range(self):
        import tempfile

        from kaos_office.pptx.reader import get_slide_text

        data = make_minimal_pptx()
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            f.write(data)
            f.flush()
            with pytest.raises(ValueError, match="out of range"):
                get_slide_text(f.name, 99)

    def test_list_slides(self):
        import tempfile

        from kaos_office.pptx.reader import list_slides

        data = make_minimal_pptx()
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            f.write(data)
            f.flush()
            slides = list_slides(f.name)
            assert len(slides) == 1
            assert slides[0]["slide_number"] == 1
            assert slides[0]["title"] == "Test Title"
