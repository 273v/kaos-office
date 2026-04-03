"""Tests for MCP tool metadata and basic functionality."""

from __future__ import annotations

from pathlib import Path

import pytest

from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    ParseDocxTool,
    SearchDocxTool,
    register_office_tools,
)
from tests.conftest import make_minimal_docx


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

    def test_all_tools_have_annotations(self):
        tools = [
            ParseDocxTool(),
            GetDocxTextTool(),
            GetDocxMarkdownTool(),
            DocxMetadataTool(),
            SearchDocxTool(),
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


class TestToolRegistration:
    def test_register_tools(self):
        from kaos_core import KaosRuntime

        runtime = KaosRuntime.default()
        count = register_office_tools(runtime)
        assert count == 12  # 5 DOCX + 3 PPTX + 4 XLSX
