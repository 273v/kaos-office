"""Tests for MCP tool metadata and basic functionality."""

from __future__ import annotations

from pathlib import Path

import pytest

from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    GetSlideNotesTool,
    ParseDocxTool,
    SearchDocxTool,
    SearchPptxTool,
    register_office_tools,
)
from tests.conftest import make_minimal_docx, make_minimal_pptx

# Namespace constants for building PPTX notes XML
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _make_notes_xml(text: str) -> str:
    """Build a minimal notesSlide XML with the given speaker notes text."""
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="{_P_NS}" xmlns:a="{_A_NS}" xmlns:r="{_R_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Notes Placeholder 2"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>"""


class TestToolMetadata:
    def test_parse_docx_tool_metadata(self):
        tool = ParseDocxTool()
        meta = tool.metadata
        assert meta.name == "kaos-office-parse-docx"
        assert meta.annotations is not None
        assert meta.annotations.readOnlyHint is True
        assert meta.annotations.destructiveHint is False
        assert meta.annotations.openWorldHint is False

    def test_get_text_tool_metadata(self):
        meta = GetDocxTextTool().metadata
        assert meta.name == "kaos-office-get-text"

    def test_get_markdown_tool_metadata(self):
        meta = GetDocxMarkdownTool().metadata
        assert meta.name == "kaos-office-get-markdown"

    def test_metadata_tool_metadata(self):
        meta = DocxMetadataTool().metadata
        assert meta.name == "kaos-office-metadata"

    def test_search_tool_metadata(self):
        meta = SearchDocxTool().metadata
        assert meta.name == "kaos-office-search"
        # Should have query parameter
        param_names = [p.name for p in meta.input_schema]
        assert "query" in param_names

    def test_search_pptx_tool_metadata(self):
        meta = SearchPptxTool().metadata
        assert meta.name == "kaos-office-search-pptx"
        assert meta.annotations is not None
        assert meta.annotations.readOnlyHint is True
        assert meta.annotations.destructiveHint is False
        assert meta.annotations.openWorldHint is False
        param_names = [p.name for p in meta.input_schema]
        assert "path" in param_names
        assert "query" in param_names
        assert "top_k" in param_names

    def test_get_slide_notes_tool_metadata(self):
        meta = GetSlideNotesTool().metadata
        assert meta.name == "kaos-office-get-slide-notes"
        assert meta.annotations is not None
        assert meta.annotations.readOnlyHint is True
        assert meta.annotations.destructiveHint is False
        assert meta.annotations.openWorldHint is False
        param_names = [p.name for p in meta.input_schema]
        assert "path" in param_names
        assert "slide" in param_names

    def test_all_tools_have_annotations(self):
        tools = [
            ParseDocxTool(),
            GetDocxTextTool(),
            GetDocxMarkdownTool(),
            DocxMetadataTool(),
            SearchDocxTool(),
            SearchPptxTool(),
            GetSlideNotesTool(),
        ]
        for tool in tools:
            assert tool.metadata.annotations is not None, (
                f"{tool.metadata.name} missing annotations"
            )


class TestToolExecution:
    @pytest.fixture
    def docx_path(self, tmp_path: Path) -> str:
        path = tmp_path / "test.docx"
        path.write_bytes(make_minimal_docx())
        return str(path)

    @pytest.fixture
    def pptx_path(self, tmp_path: Path) -> str:
        """Create a minimal PPTX with one slide (default title shape)."""
        path = tmp_path / "test.pptx"
        path.write_bytes(make_minimal_pptx())
        return str(path)

    @pytest.fixture
    def pptx_with_notes_path(self, tmp_path: Path) -> str:
        """Create a PPTX with speaker notes on slide 1."""
        path = tmp_path / "test_notes.pptx"
        notes_xml = _make_notes_xml("These are the speaker notes for slide one.")
        path.write_bytes(make_minimal_pptx(notes_xmls={0: notes_xml}))
        return str(path)

    @pytest.mark.asyncio
    async def test_parse_docx_file_not_found(self):
        tool = ParseDocxTool()
        result = await tool.execute({"path": "/nonexistent/file.docx"})
        assert result.isError is True

    @pytest.mark.asyncio
    async def test_get_text_basic(self, docx_path):
        tool = GetDocxTextTool()
        result = await tool.execute({"path": docx_path})
        assert result.isError is False
        assert "Hello" in str(result.content)

    @pytest.mark.asyncio
    async def test_get_markdown_basic(self, docx_path):
        tool = GetDocxMarkdownTool()
        result = await tool.execute({"path": docx_path})
        assert result.isError is False
        assert "Hello" in str(result.content)

    @pytest.mark.asyncio
    async def test_metadata_basic(self, docx_path):
        tool = DocxMetadataTool()
        result = await tool.execute({"path": docx_path})
        assert result.isError is False

    @pytest.mark.asyncio
    async def test_search_empty_query(self, docx_path):
        tool = SearchDocxTool()
        result = await tool.execute({"path": docx_path, "query": ""})
        assert result.isError is True

    @pytest.mark.asyncio
    async def test_search_basic(self, docx_path):
        tool = SearchDocxTool()
        result = await tool.execute({"path": docx_path, "query": "Hello"})
        assert result.isError is False

    # --- SearchPptxTool tests ---

    @pytest.mark.asyncio
    async def test_search_pptx_file_not_found(self):
        tool = SearchPptxTool()
        result = await tool.execute({"path": "/nonexistent/file.pptx", "query": "test"})
        assert result.isError is True
        assert "not found" in str(result.content).lower()

    @pytest.mark.asyncio
    async def test_search_pptx_empty_query(self, pptx_path):
        tool = SearchPptxTool()
        result = await tool.execute({"path": pptx_path, "query": ""})
        assert result.isError is True
        assert "required" in str(result.content).lower()

    @pytest.mark.asyncio
    async def test_search_pptx_basic(self, pptx_path):
        tool = SearchPptxTool()
        result = await tool.execute({"path": pptx_path, "query": "Test Title"})
        assert result.isError is False

    # --- GetSlideNotesTool tests ---

    @pytest.mark.asyncio
    async def test_get_slide_notes_file_not_found(self):
        tool = GetSlideNotesTool()
        result = await tool.execute({"path": "/nonexistent/file.pptx", "slide": 1})
        assert result.isError is True
        assert "not found" in str(result.content).lower()

    @pytest.mark.asyncio
    async def test_get_slide_notes_out_of_range(self, pptx_path):
        tool = GetSlideNotesTool()
        result = await tool.execute({"path": pptx_path, "slide": 999})
        assert result.isError is True
        assert "out of range" in str(result.content).lower()

    @pytest.mark.asyncio
    async def test_get_slide_notes_no_notes(self, pptx_path):
        tool = GetSlideNotesTool()
        result = await tool.execute({"path": pptx_path, "slide": 1})
        assert result.isError is False
        assert "no speaker notes" in str(result.content).lower()

    @pytest.mark.asyncio
    async def test_get_slide_notes_with_notes(self, pptx_with_notes_path):
        tool = GetSlideNotesTool()
        result = await tool.execute({"path": pptx_with_notes_path, "slide": 1})
        assert result.isError is False
        assert "speaker notes for slide one" in str(result.content).lower()


class TestToolRegistration:
    def test_register_tools(self):
        from kaos_core import KaosRuntime

        runtime = KaosRuntime.default()
        count = register_office_tools(runtime)
        assert count == 14  # 5 DOCX + 5 PPTX + 4 XLSX
