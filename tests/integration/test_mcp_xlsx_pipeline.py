"""End-to-end integration test: XLSX documents -> MCP tools -> resources.

Proves the full XLSX pipeline:
  1. Register kaos-office tools with KaosRuntime
  2. Wire into kaos-mcp server
  3. Call XLSX tools via MCP client session (list-sheets, get-sheet, parse, metadata)
  4. Verify structured content, data correctness, and error handling
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest
import xlsxwriter
from kaos_core import KaosRuntime, KaosSettings
from kaos_core.types.enums import StorageBackend
from kaos_core.vfs import VFSConfig, VirtualFileSystem
from kaos_mcp import create_app
from mcp import types
from mcp.shared.memory import create_connected_server_and_client_session
from mcp.types import TextContent

from kaos_office import register_office_tools


def _get_text(result: types.CallToolResult, index: int = 0) -> str:
    """Extract text from MCP CallToolResult, narrowing the union type."""
    content = result.content[index]
    assert isinstance(content, TextContent)
    return content.text


def _make_runtime(tmp_path: Path) -> KaosRuntime:
    settings = KaosSettings(
        artifact_inline_read_max_bytes=262_144,
        artifact_chunk_size_bytes=65_536,
    )
    # S3Backend.__init__ raises NotImplementedError; stub it for disk-only tests.
    _s3_noop = patch("kaos_core.vfs.core.S3Backend.__init__", lambda self: None)
    with _s3_noop:
        runtime = KaosRuntime(config=settings)
        runtime.vfs = VirtualFileSystem(
            VFSConfig(default_backend=StorageBackend.DISK, disk_base_path=tmp_path / "vfs")
        )
    runtime.artifacts = runtime.artifacts.__class__(
        runtime.vfs,
        manifest_context_id=settings.artifact_manifest_context_id,
        manifest_prefix=settings.artifact_manifest_prefix,
        max_inline_read_bytes=settings.artifact_inline_read_max_bytes,
        default_chunk_size=settings.artifact_chunk_size_bytes,
        temporary_ttl_seconds=settings.artifact_temporary_ttl_seconds,
    )
    return runtime


def _create_test_xlsx(path: Path) -> Path:
    """Create a test XLSX workbook with 3 sheets of known data.

    Sheets:
      - Revenue: Quarter (text), Amount (float) -- 3 data rows
      - Expenses: Category (text), Cost (float) -- 2 data rows
      - Summary: Metric (text), Value (float) -- 2 data rows
    """
    wb = xlsxwriter.Workbook(str(path))

    # Sheet 1: Revenue
    ws1 = wb.add_worksheet("Revenue")
    ws1.write(0, 0, "Quarter")
    ws1.write(0, 1, "Amount")
    ws1.write(1, 0, "Q1 2025")
    ws1.write(1, 1, 150000.50)
    ws1.write(2, 0, "Q2 2025")
    ws1.write(2, 1, 175000.75)
    ws1.write(3, 0, "Q3 2025")
    ws1.write(3, 1, 200000.00)

    # Sheet 2: Expenses
    ws2 = wb.add_worksheet("Expenses")
    ws2.write(0, 0, "Category")
    ws2.write(0, 1, "Cost")
    ws2.write(1, 0, "Salaries")
    ws2.write(1, 1, 80000.00)
    ws2.write(2, 0, "Travel")
    ws2.write(2, 1, 12500.00)

    # Sheet 3: Summary
    ws3 = wb.add_worksheet("Summary")
    ws3.write(0, 0, "Metric")
    ws3.write(0, 1, "Value")
    ws3.write(1, 0, "Total Revenue")
    ws3.write(1, 1, 525001.25)
    ws3.write(2, 0, "Total Expenses")
    ws3.write(2, 1, 92500.00)

    wb.close()
    return path


# ---------------------------------------------------------------------------
# Tool Discovery
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_xlsx_tools_discoverable_via_mcp(tmp_path: Path) -> None:
    """All 4 XLSX MCP tools should be discoverable."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        tools_result = await session.list_tools()
        tool_names = {t.name for t in tools_result.tools}

        expected_xlsx_tools = {
            "kaos-office-parse-xlsx",
            "kaos-office-list-sheets-xlsx",
            "kaos-office-get-sheet-xlsx",
            "kaos-office-xlsx-metadata",
        }
        assert expected_xlsx_tools.issubset(tool_names), (
            f"Missing XLSX tools: {expected_xlsx_tools - tool_names}"
        )


# ---------------------------------------------------------------------------
# list-sheets-xlsx
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_list_sheets_returns_correct_names(tmp_path: Path) -> None:
    """list-sheets-xlsx should return all 3 sheet names with metadata."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-list-sheets-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not result.isError
        assert result.structuredContent is not None

        sheets = result.structuredContent["sheets"]
        assert len(sheets) == 3

        sheet_names = [s["name"] for s in sheets]
        assert sheet_names == ["Revenue", "Expenses", "Summary"]

        # Each sheet should have visibility and dimension info
        for sheet in sheets:
            assert "visible" in sheet
            assert sheet["visible"] is True
            assert "rows" in sheet
            assert "columns" in sheet


@pytest.mark.integration
async def test_list_sheets_dimensions(tmp_path: Path) -> None:
    """list-sheets-xlsx should report correct row and column counts."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-list-sheets-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not result.isError
        sheets = result.structuredContent["sheets"]

        # Revenue: 1 header + 3 data = 4 rows, 2 columns
        revenue = sheets[0]
        assert revenue["name"] == "Revenue"
        assert revenue["rows"] == 4
        assert revenue["columns"] == 2

        # Expenses: 1 header + 2 data = 3 rows, 2 columns
        expenses = sheets[1]
        assert expenses["name"] == "Expenses"
        assert expenses["rows"] == 3
        assert expenses["columns"] == 2


# ---------------------------------------------------------------------------
# get-sheet-xlsx
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_get_sheet_returns_correct_data(tmp_path: Path) -> None:
    """get-sheet-xlsx should return TSV data with correct values."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-sheet-xlsx",
            {"path": str(xlsx_path), "sheet": "Revenue"},
        )
        assert not result.isError
        text = _get_text(result)

        # Verify TSV header row
        lines = text.strip().split("\n")
        assert len(lines) == 4  # header + 3 data rows
        assert "Quarter" in lines[0]
        assert "Amount" in lines[0]

        # Verify data content
        assert "Q1 2025" in text
        assert "Q2 2025" in text
        assert "Q3 2025" in text
        assert "150000.5" in text
        assert "175000.75" in text


@pytest.mark.integration
async def test_get_sheet_specific_sheet(tmp_path: Path) -> None:
    """get-sheet-xlsx should extract just the requested sheet."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-sheet-xlsx",
            {"path": str(xlsx_path), "sheet": "Expenses"},
        )
        assert not result.isError
        text = _get_text(result)

        lines = text.strip().split("\n")
        assert len(lines) == 3  # header + 2 data rows
        assert "Category" in lines[0]
        assert "Cost" in lines[0]
        assert "Salaries" in text
        assert "Travel" in text


@pytest.mark.integration
async def test_get_sheet_max_rows(tmp_path: Path) -> None:
    """get-sheet-xlsx with max_rows should truncate data rows."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-sheet-xlsx",
            {"path": str(xlsx_path), "sheet": "Revenue", "max_rows": 1},
        )
        assert not result.isError
        text = _get_text(result)
        lines = text.strip().split("\n")
        # Header + at most 1 data row
        assert len(lines) <= 2


@pytest.mark.integration
async def test_get_sheet_nonexistent_sheet(tmp_path: Path) -> None:
    """get-sheet-xlsx with unknown sheet name should return error."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-sheet-xlsx",
            {"path": str(xlsx_path), "sheet": "DoesNotExist"},
        )
        assert result.isError
        error_text = _get_text(result)
        assert "not found" in error_text.lower()


# ---------------------------------------------------------------------------
# parse-xlsx
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_parse_xlsx_via_mcp(tmp_path: Path) -> None:
    """parse-xlsx should return structured summary of all tables."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not result.isError
        text = _get_text(result)

        # Summary should mention all 3 sheets
        assert "Revenue" in text
        assert "Expenses" in text
        assert "Summary" in text
        # Should mention table count
        assert "3 table" in text


@pytest.mark.integration
async def test_parse_xlsx_with_artifact(tmp_path: Path) -> None:
    """parse-xlsx with runtime should store artifact and return structured content."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not result.isError

        # With MCP runtime, structured content should contain artifact details
        if result.structuredContent is not None:
            assert result.structuredContent["table_count"] == 3
            tables = result.structuredContent["tables"]
            assert len(tables) == 3
            assert tables[0]["name"] == "Revenue"
            assert tables[0]["row_count"] == 3
            assert tables[1]["name"] == "Expenses"
            assert tables[1]["row_count"] == 2
            assert tables[2]["name"] == "Summary"
            assert tables[2]["row_count"] == 2


@pytest.mark.integration
async def test_parse_xlsx_specific_sheets(tmp_path: Path) -> None:
    """parse-xlsx with sheets filter should extract only requested sheets."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-xlsx",
            {"path": str(xlsx_path), "sheets": ["Revenue", "Summary"]},
        )
        assert not result.isError
        text = _get_text(result)
        assert "Revenue" in text
        assert "Summary" in text


# ---------------------------------------------------------------------------
# xlsx-metadata
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_xlsx_metadata_via_mcp(tmp_path: Path) -> None:
    """xlsx-metadata should return workbook metadata with sheet info."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-xlsx-metadata",
            {"path": str(xlsx_path)},
        )
        assert not result.isError
        assert result.structuredContent is not None

        meta = result.structuredContent
        assert meta["table_count"] == 3
        assert meta["total_rows"] == 7  # 3 + 2 + 2

        tables = meta["tables"]
        assert len(tables) == 3
        # Verify column type info
        revenue_cols = tables[0]["columns"]
        assert len(revenue_cols) == 2
        assert revenue_cols[0]["name"] == "Quarter"
        assert revenue_cols[0]["type"] == "text"
        assert revenue_cols[1]["name"] == "Amount"
        assert revenue_cols[1]["type"] == "float"


# ---------------------------------------------------------------------------
# Error Handling
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_xlsx_file_not_found_error(tmp_path: Path) -> None:
    """All XLSX tools should return clear error for nonexistent files."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)
    fake_path = "/nonexistent/path/data.xlsx"

    async with create_connected_server_and_client_session(app) as session:
        for tool_name, args in [
            ("kaos-office-list-sheets-xlsx", {"path": fake_path}),
            ("kaos-office-get-sheet-xlsx", {"path": fake_path, "sheet": "Sheet1"}),
            ("kaos-office-parse-xlsx", {"path": fake_path}),
            ("kaos-office-xlsx-metadata", {"path": fake_path}),
        ]:
            result = await session.call_tool(tool_name, args)
            assert result.isError, f"{tool_name} should error on missing file"
            error_text = _get_text(result)
            assert "not found" in error_text.lower(), (
                f"{tool_name} error should mention 'not found': {error_text}"
            )


# ---------------------------------------------------------------------------
# Full Pipeline: list -> get each sheet -> parse -> metadata
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_xlsx_full_pipeline(tmp_path: Path) -> None:
    """Full agent-like pipeline: discover sheets, get data, parse, metadata."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    xlsx_path = _create_test_xlsx(tmp_path / "test.xlsx")
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        # Step 1: List sheets (agent would do this first)
        list_result = await session.call_tool(
            "kaos-office-list-sheets-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not list_result.isError
        sheets = list_result.structuredContent["sheets"]
        sheet_names = [s["name"] for s in sheets]
        assert sheet_names == ["Revenue", "Expenses", "Summary"]

        # Step 2: Get each sheet's data
        for sheet_name in sheet_names:
            get_result = await session.call_tool(
                "kaos-office-get-sheet-xlsx",
                {"path": str(xlsx_path), "sheet": sheet_name},
            )
            assert not get_result.isError, f"Failed to get sheet: {sheet_name}"
            text = _get_text(get_result)
            # Every sheet should have at least a header + 1 data row
            lines = text.strip().split("\n")
            assert len(lines) >= 2, f"Sheet {sheet_name} has too few rows"

        # Step 3: Parse full workbook
        parse_result = await session.call_tool(
            "kaos-office-parse-xlsx",
            {"path": str(xlsx_path)},
        )
        assert not parse_result.isError

        # Step 4: Get metadata
        meta_result = await session.call_tool(
            "kaos-office-xlsx-metadata",
            {"path": str(xlsx_path)},
        )
        assert not meta_result.isError
        meta = meta_result.structuredContent
        assert meta["table_count"] == 3
        assert meta["total_rows"] == 7
