"""Unit tests for the three MCP writer tools (DOCX / PPTX / XLSX)."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from kaos_content.model.blocks import Heading, Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.model.tabular import Column, ColumnType, TabularDocument
from kaos_content.model.tabular import Table as TabTable

from kaos_office.tools import WriteDocxTool, WritePptxTool, WriteXlsxTool


@pytest.fixture
def simple_content_doc() -> ContentDocument:
    return ContentDocument(
        metadata=DocumentMetadata(title="Hello"),
        body=(
            Heading(depth=1, children=(Text(value="Greeting"),)),
            Paragraph(children=(Text(value="Hello from the writer tool."),)),
        ),
    )


@pytest.fixture
def simple_tabular_doc() -> TabularDocument:
    table = TabTable(
        name="Sheet1",
        columns=(
            Column(name="id", column_type=ColumnType.INTEGER),
            Column(name="label", column_type=ColumnType.TEXT),
        ),
        rows=(
            (1, "alpha"),
            (2, "beta"),
        ),
    )
    return TabularDocument(tables=(table,))


class TestWriteDocxTool:
    @pytest.mark.asyncio
    async def test_writes_file_from_inline_json(
        self, simple_content_doc: ContentDocument, tmp_path: Path
    ) -> None:
        out = tmp_path / "out.docx"
        tool = WriteDocxTool()
        result = await tool.execute(
            {
                "document_json": simple_content_doc.model_dump_json(),
                "output_path": str(out),
            }
        )
        assert result.isError is False, result.content
        assert out.exists()
        assert out.stat().st_size > 0
        with zipfile.ZipFile(out) as zf:
            assert "word/document.xml" in zf.namelist()
        structured = result.require_structured()
        assert structured["format"] == "docx"
        assert structured["path"] == str(out)
        assert structured["size_bytes"] > 0
        assert structured["block_count"] == 2

    @pytest.mark.asyncio
    async def test_missing_document_returns_error(self, tmp_path: Path) -> None:
        tool = WriteDocxTool()
        result = await tool.execute({"output_path": str(tmp_path / "x.docx")})
        assert result.isError is True
        assert "Missing document" in result.require_text()

    @pytest.mark.asyncio
    async def test_refuses_to_overwrite_without_force(
        self, simple_content_doc: ContentDocument, tmp_path: Path
    ) -> None:
        out = tmp_path / "exists.docx"
        out.write_bytes(b"existing")
        tool = WriteDocxTool()
        result = await tool.execute(
            {
                "document_json": simple_content_doc.model_dump_json(),
                "output_path": str(out),
            }
        )
        assert result.isError is True
        assert "overwrite" in result.require_text().lower()
        assert out.read_bytes() == b"existing"  # untouched

    @pytest.mark.asyncio
    async def test_overwrites_with_force(
        self, simple_content_doc: ContentDocument, tmp_path: Path
    ) -> None:
        out = tmp_path / "exists.docx"
        out.write_bytes(b"old")
        tool = WriteDocxTool()
        result = await tool.execute(
            {
                "document_json": simple_content_doc.model_dump_json(),
                "output_path": str(out),
                "force": True,
            }
        )
        assert result.isError is False
        # Real DOCX bytes are far larger than the placeholder
        assert out.stat().st_size > 100

    @pytest.mark.asyncio
    async def test_invalid_json_returns_error(self, tmp_path: Path) -> None:
        tool = WriteDocxTool()
        result = await tool.execute(
            {
                "document_json": "{not valid json",
                "output_path": str(tmp_path / "x.docx"),
            }
        )
        assert result.isError is True
        assert "ContentDocument" in result.require_text()

    def test_metadata_annotations(self) -> None:
        tool = WriteDocxTool()
        meta = tool.metadata
        assert meta.name == "kaos-office-write-docx"
        ann = meta.annotations
        assert ann is not None
        assert ann.readOnlyHint is False
        assert ann.idempotentHint is False
        assert ann.openWorldHint is True


class TestWriteXlsxTool:
    @pytest.mark.asyncio
    async def test_writes_xlsx_from_tabular_json(
        self, simple_tabular_doc: TabularDocument, tmp_path: Path
    ) -> None:
        out = tmp_path / "out.xlsx"
        tool = WriteXlsxTool()
        result = await tool.execute(
            {
                "document_json": simple_tabular_doc.model_dump_json(),
                "output_path": str(out),
            }
        )
        assert result.isError is False, result.content
        assert out.exists()
        with zipfile.ZipFile(out) as zf:
            assert "xl/workbook.xml" in zf.namelist()
        structured = result.require_structured()
        assert structured["format"] == "xlsx"
        assert structured["table_count"] == 1

    @pytest.mark.asyncio
    async def test_rejects_invalid_json(self, tmp_path: Path) -> None:
        tool = WriteXlsxTool()
        result = await tool.execute(
            {
                "document_json": "not valid json at all",
                "output_path": str(tmp_path / "out.xlsx"),
            }
        )
        assert result.isError is True
        assert "TabularDocument" in result.require_text()


class TestWritePptxTool:
    @pytest.mark.asyncio
    async def test_writes_pptx(self, simple_content_doc: ContentDocument, tmp_path: Path) -> None:
        pytest.importorskip("pptx")
        out = tmp_path / "out.pptx"
        tool = WritePptxTool()
        result = await tool.execute(
            {
                "document_json": simple_content_doc.model_dump_json(),
                "output_path": str(out),
            }
        )
        assert result.isError is False, result.content
        assert out.exists()
        with zipfile.ZipFile(out) as zf:
            names = zf.namelist()
        assert "ppt/presentation.xml" in names
        structured = result.require_structured()
        assert structured["format"] == "pptx"

    @pytest.mark.asyncio
    async def test_missing_template_returns_error(
        self, simple_content_doc: ContentDocument, tmp_path: Path
    ) -> None:
        pytest.importorskip("pptx")
        tool = WritePptxTool()
        result = await tool.execute(
            {
                "document_json": simple_content_doc.model_dump_json(),
                "output_path": str(tmp_path / "out.pptx"),
                "template_path": str(tmp_path / "does-not-exist.pptx"),
            }
        )
        assert result.isError is True
        assert "Template not found" in result.require_text()


class TestRegisterIncludesWriters:
    def test_all_writers_registered(self) -> None:
        """register_office_tools returns 17 and includes the three writers."""
        from unittest.mock import MagicMock

        from kaos_office.tools import register_office_tools

        runtime = MagicMock()
        count = register_office_tools(runtime)
        # 14 readers + 3 writers = 17
        assert count == 17

        registered = [c.args[0] for c in runtime.tools.register_tool.call_args_list]
        names = {t.metadata.name for t in registered}
        assert "kaos-office-write-docx" in names
        assert "kaos-office-write-pptx" in names
        assert "kaos-office-write-xlsx" in names
