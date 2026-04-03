"""Unit tests for PPTX MCP tools."""

from __future__ import annotations

import pytest

from tests.conftest import make_minimal_pptx


class TestPptxToolMetadata:
    """Test PPTX tool metadata."""

    def test_parse_pptx_tool(self):
        from kaos_office.tools import ParsePptxTool

        tool = ParsePptxTool()
        meta = tool.metadata
        assert meta.name == "kaos-office-parse-pptx"
        assert meta.annotations.readOnlyHint is True
        assert meta.annotations.destructiveHint is False

    def test_list_slides_tool(self):
        from kaos_office.tools import ListSlidesTool

        tool = ListSlidesTool()
        meta = tool.metadata
        assert meta.name == "kaos-office-list-slides"
        assert meta.annotations.readOnlyHint is True

    def test_get_slide_tool(self):
        from kaos_office.tools import GetSlideTool

        tool = GetSlideTool()
        meta = tool.metadata
        assert meta.name == "kaos-office-get-slide"
        assert meta.annotations.readOnlyHint is True
        # Should have slide_number parameter
        params = {p.name for p in meta.input_schema}
        assert "path" in params
        assert "slide_number" in params


class TestPptxToolExecution:
    """Test PPTX tool execution."""

    @pytest.fixture
    def pptx_file(self, tmp_path):
        """Create a temp PPTX file for testing."""
        data = make_minimal_pptx()
        path = tmp_path / "test.pptx"
        path.write_bytes(data)
        return path

    async def test_parse_pptx_file_not_found(self):
        from kaos_office.tools import ParsePptxTool

        tool = ParsePptxTool()
        result = await tool.execute({"path": "/nonexistent/test.pptx"})
        assert result.isError is True
        assert "not found" in result.content[0].text.lower()

    async def test_parse_pptx_success(self, pptx_file):
        from kaos_office.tools import ParsePptxTool

        tool = ParsePptxTool()
        result = await tool.execute({"path": str(pptx_file)})
        assert result.isError is not True
        assert "Parsed" in result.content[0].text

    async def test_list_slides_success(self, pptx_file):
        from kaos_office.tools import ListSlidesTool

        tool = ListSlidesTool()
        result = await tool.execute({"path": str(pptx_file)})
        assert result.isError is not True
        # Summary in content, structured data in structuredContent
        assert "1 slide" in result.content[0].text
        assert result.structuredContent is not None
        slides = result.structuredContent["slides"]
        assert len(slides) == 1
        assert slides[0]["slide_number"] == 1

    async def test_get_slide_success(self, pptx_file):
        from kaos_office.tools import GetSlideTool

        tool = GetSlideTool()
        result = await tool.execute({"path": str(pptx_file), "slide_number": 1})
        assert result.isError is not True
        assert "Test Title" in result.content[0].text

    async def test_get_slide_out_of_range(self, pptx_file):
        from kaos_office.tools import GetSlideTool

        tool = GetSlideTool()
        result = await tool.execute({"path": str(pptx_file), "slide_number": 99})
        assert result.isError is True
        assert "out of range" in result.content[0].text.lower()

    async def test_list_slides_file_not_found(self):
        from kaos_office.tools import ListSlidesTool

        tool = ListSlidesTool()
        result = await tool.execute({"path": "/nonexistent/test.pptx"})
        assert result.isError is True


class TestToolRegistration:
    """Test tool registration includes PPTX tools."""

    def test_register_includes_pptx(self):
        from unittest.mock import MagicMock

        from kaos_office.tools import register_office_tools

        runtime = MagicMock()
        count = register_office_tools(runtime)
        assert count == 12  # 5 DOCX + 3 PPTX + 4 XLSX
