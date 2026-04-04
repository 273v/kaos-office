"""Unit tests for SmartArt OPC fallback extraction."""

from __future__ import annotations

from pathlib import Path

import pytest
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.ooxml.namespace import DGM, A
from kaos_office.pptx.smartart import _extract_texts_from_data_model


class TestExtractTextsFromDataModel:
    """Test the data model text extraction logic."""

    def test_simple_nodes(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="1">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>First Item</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="2">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Second Item</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == ["First Item", "Second Item"]

    def test_doc_node_skipped(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Root Doc</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == []

    def test_transition_nodes_skipped(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="1" type="sibTrans">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Transition</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="2" type="parTrans">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Parent Trans</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="3">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Real Content</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == ["Real Content"]

    def test_empty_text_nodes(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="1">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == []

    def test_multi_run_text(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="1">
      <dgm:t>
        <a:bodyPr/>
        <a:p><a:r><a:t>Hello</a:t></a:r><a:r><a:t>World</a:t></a:r></a:p>
      </dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == ["Hello World"]

    def test_no_ptlst(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == []

    def test_explicit_node_type(self):
        xml = f"""\
<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:t><a:bodyPr/><a:p><a:endParaRPr/></a:p></dgm:t>
    </dgm:pt>
    <dgm:pt modelId="1" type="node">
      <dgm:t><a:bodyPr/><a:p><a:r><a:t>Explicit Node</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>"""
        root = etree.fromstring(xml.encode())
        texts = _extract_texts_from_data_model(root)
        assert texts == ["Explicit Node"]


class TestSmartArtWithRealFixture:
    """Test SmartArt extraction with the real SmartArt.pptx fixture."""

    def test_smartart_pptx(self):
        """The SmartArt.pptx from Apache POI has placeholder text (empty)."""
        from kaos_office.pptx.reader import parse_pptx

        fixture = Path(__file__).parent.parent / "fixtures" / "pptx" / "stress" / "SmartArt.pptx"
        if not fixture.exists():
            pytest.skip("SmartArt.pptx fixture not available")
        doc = parse_pptx(fixture)
        # The SmartArt fixture has placeholder text — should not crash
        assert doc is not None
