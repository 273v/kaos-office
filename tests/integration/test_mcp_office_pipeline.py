"""End-to-end integration test: Office documents → MCP tools → resources.

Proves the full pipeline:
  1. Register kaos-office tools with KaosRuntime
  2. Wire into kaos-mcp server
  3. Call DOCX and PPTX tools via MCP client session
  4. Read markdown, outline, metadata via MCP resource templates
  5. Verify search works through MCP boundary
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from kaos_core import KaosContext, KaosRuntime, KaosSettings
from kaos_core.types.enums import StorageBackend
from kaos_core.vfs import VFSConfig, VirtualFileSystem
from kaos_mcp import create_app
from mcp import types
from mcp.shared.memory import create_connected_server_and_client_session
from pydantic import AnyUrl

from kaos_office import register_office_tools
from tests.conftest import KELVIN_FIXTURES, KELVIN_PPTX_FIXTURES, make_minimal_docx

BATTLE_DIR = Path(__file__).parent.parent / "fixtures" / "pptx" / "battle"


def _make_runtime(tmp_path: Path) -> KaosRuntime:
    settings = KaosSettings(
        artifact_inline_read_max_bytes=262_144,
        artifact_chunk_size_bytes=65_536,
    )
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


# ---------------------------------------------------------------------------
# Tool Discovery
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_list_office_tools_via_mcp(tmp_path: Path) -> None:
    """All 8 kaos-office tools should be discoverable via MCP."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        tools_result = await session.list_tools()
        tool_names = {t.name for t in tools_result.tools}

        # DOCX tools
        assert "kaos-office-parse-docx" in tool_names
        assert "kaos-office-get-text" in tool_names
        assert "kaos-office-get-markdown" in tool_names
        assert "kaos-office-metadata" in tool_names
        assert "kaos-office-search" in tool_names

        # PPTX tools
        assert "kaos-office-parse-pptx" in tool_names
        assert "kaos-office-list-slides" in tool_names
        assert "kaos-office-get-slide" in tool_names

        assert len(tool_names & {"kaos-office-parse-docx", "kaos-office-parse-pptx"}) == 2


# ---------------------------------------------------------------------------
# DOCX via MCP
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_parse_docx_via_mcp(tmp_path: Path) -> None:
    """Parse DOCX through MCP and verify structured result."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    # Create a test DOCX
    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx())

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-docx",
            {"path": str(docx_path)},
        )
        assert not result.isError
        assert len(result.content) >= 1
        text_contents = [c for c in result.content if isinstance(c, types.TextContent)]
        assert len(text_contents) >= 1


@pytest.mark.integration
async def test_get_text_via_mcp(tmp_path: Path) -> None:
    """Get plain text from DOCX through MCP."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx())

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-text",
            {"path": str(docx_path)},
        )
        assert not result.isError
        text = result.content[0].text
        assert "Hello" in text


@pytest.mark.integration
async def test_get_markdown_via_mcp(tmp_path: Path) -> None:
    """Get markdown from DOCX through MCP."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx())

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-markdown",
            {"path": str(docx_path)},
        )
        assert not result.isError
        text = result.content[0].text
        assert "Hello" in text


@pytest.mark.integration
async def test_metadata_via_mcp(tmp_path: Path) -> None:
    """Get DOCX metadata through MCP."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    core_xml = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>MCP Test Document</dc:title>
  <dc:creator>Test Author</dc:creator>
</cp:coreProperties>"""
    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx(core_xml=core_xml))

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-metadata",
            {"path": str(docx_path)},
        )
        assert not result.isError
        text = result.content[0].text
        meta = json.loads(text)
        assert meta["title"] == "MCP Test Document"
        assert meta["creator"] == "Test Author"


@pytest.mark.integration
async def test_search_docx_via_mcp(tmp_path: Path) -> None:
    """Search within DOCX through MCP."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx())

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-search",
            {"path": str(docx_path), "query": "Hello"},
        )
        assert not result.isError


@pytest.mark.integration
async def test_docx_error_handling_via_mcp(tmp_path: Path) -> None:
    """Verify error messages come through MCP correctly."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-docx",
            {"path": "/nonexistent/path/test.docx"},
        )
        assert result.isError
        error_text = result.content[0].text
        assert "not found" in error_text.lower()
        # Error should include recovery guidance
        assert "path" in error_text.lower()


# ---------------------------------------------------------------------------
# PPTX via MCP
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_parse_pptx_via_mcp(tmp_path: Path) -> None:
    """Parse PPTX through MCP and verify result."""
    pptx_path = BATTLE_DIR / "rich-text.pptx"
    if not pptx_path.exists():
        pytest.skip("Battle test fixtures not generated")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-parse-pptx",
            {"path": str(pptx_path)},
        )
        assert not result.isError
        assert len(result.content) >= 1


@pytest.mark.integration
async def test_list_slides_via_mcp(tmp_path: Path) -> None:
    """List PPTX slides through MCP."""
    pptx_path = BATTLE_DIR / "charts.pptx"
    if not pptx_path.exists():
        pytest.skip("Battle test fixtures not generated")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-list-slides",
            {"path": str(pptx_path)},
        )
        assert not result.isError
        slides = json.loads(result.content[0].text)
        assert len(slides) == 3
        assert slides[0]["slide_number"] == 1


@pytest.mark.integration
async def test_get_slide_via_mcp(tmp_path: Path) -> None:
    """Get specific slide text through MCP."""
    pptx_path = BATTLE_DIR / "notes.pptx"
    if not pptx_path.exists():
        pytest.skip("Battle test fixtures not generated")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-slide",
            {"path": str(pptx_path), "slide_number": 1},
        )
        assert not result.isError
        text = result.content[0].text
        assert "Slide With Notes" in text


@pytest.mark.integration
async def test_get_slide_out_of_range_via_mcp(tmp_path: Path) -> None:
    """Verify out-of-range slide number returns error via MCP."""
    pptx_path = BATTLE_DIR / "notes.pptx"
    if not pptx_path.exists():
        pytest.skip("Battle test fixtures not generated")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        result = await session.call_tool(
            "kaos-office-get-slide",
            {"path": str(pptx_path), "slide_number": 999},
        )
        assert result.isError
        assert "out of range" in result.content[0].text.lower()


# ---------------------------------------------------------------------------
# Artifact + Resource Templates
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_docx_artifact_resources_via_mcp(tmp_path: Path) -> None:
    """Parse DOCX, store artifact, then read via MCP resource templates."""
    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)

    # Parse and store as artifact directly
    context = KaosContext.create(session_id="office-mcp-test", runtime=runtime)
    from kaos_content.artifacts import store_document

    from kaos_office.docx.reader import parse_docx

    body_xml = """
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
    <w:p><w:r><w:t>This is the introduction paragraph.</w:t></w:r></w:p>
    """
    styles_xml = """<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
      </w:style>
    </w:styles>"""

    docx_path = tmp_path / "test.docx"
    docx_path.write_bytes(make_minimal_docx(body_xml=body_xml, styles_xml=styles_xml))
    doc = parse_docx(docx_path)
    manifest = await store_document(doc, runtime, context, name="test-doc")
    artifact_id = manifest.artifact_id

    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        # Verify resource templates exist
        templates = await session.list_resource_templates()
        template_uris = {t.uriTemplate for t in templates.resourceTemplates}
        assert "kaos://content/{artifact_id}/markdown" in template_uris
        assert "kaos://content/{artifact_id}/outline" in template_uris

        # Read markdown
        md_result = await session.read_resource(AnyUrl(f"kaos://content/{artifact_id}/markdown"))
        md_text = md_result.contents[0]
        assert isinstance(md_text, types.TextResourceContents)
        assert "Introduction" in md_text.text

        # Read outline
        outline_result = await session.read_resource(
            AnyUrl(f"kaos://content/{artifact_id}/outline")
        )
        outline_text = outline_result.contents[0]
        assert isinstance(outline_text, types.TextResourceContents)
        outline = json.loads(outline_text.text)
        assert isinstance(outline, list)
        assert len(outline) >= 1
        assert outline[0]["text"] == "Introduction"

        # Read metadata
        meta_result = await session.read_resource(AnyUrl(f"kaos://content/{artifact_id}/metadata"))
        meta_text = meta_result.contents[0]
        assert isinstance(meta_text, types.TextResourceContents)
        meta = json.loads(meta_text.text)
        assert isinstance(meta, dict)


# ---------------------------------------------------------------------------
# Real-world fixtures through full pipeline
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_real_docx_via_mcp_pipeline(tmp_path: Path) -> None:
    """Full pipeline with a real DOCX from kelvin fixtures."""
    docx_path = KELVIN_FIXTURES / "MultiParagraphSample.docx"
    if not docx_path.exists():
        pytest.skip("kelvin_office fixtures not available")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        # Parse
        result = await session.call_tool(
            "kaos-office-parse-docx",
            {"path": str(docx_path)},
        )
        assert not result.isError

        # Get markdown
        md_result = await session.call_tool(
            "kaos-office-get-markdown",
            {"path": str(docx_path)},
        )
        assert not md_result.isError
        md_text = md_result.content[0].text
        assert len(md_text) > 100

        # Search
        search_result = await session.call_tool(
            "kaos-office-search",
            {"path": str(docx_path), "query": "paragraph", "top_k": 3},
        )
        assert not search_result.isError


@pytest.mark.integration
async def test_real_pptx_via_mcp_pipeline(tmp_path: Path) -> None:
    """Full pipeline with a real PPTX from kelvin fixtures."""
    pptx_path = KELVIN_PPTX_FIXTURES / "Hello-World.pptx"
    if not pptx_path.exists():
        pytest.skip("kelvin_office PPTX fixtures not available")

    runtime = _make_runtime(tmp_path)
    register_office_tools(runtime)
    app = create_app(runtime)

    async with create_connected_server_and_client_session(app) as session:
        # List slides
        slides_result = await session.call_tool(
            "kaos-office-list-slides",
            {"path": str(pptx_path)},
        )
        assert not slides_result.isError
        slides = json.loads(slides_result.content[0].text)
        assert len(slides) == 9

        # Get slide 1
        slide_result = await session.call_tool(
            "kaos-office-get-slide",
            {"path": str(pptx_path), "slide_number": 1},
        )
        assert not slide_result.isError
        assert "Hello World" in slide_result.content[0].text

        # Parse full presentation
        parse_result = await session.call_tool(
            "kaos-office-parse-pptx",
            {"path": str(pptx_path)},
        )
        assert not parse_result.isError
