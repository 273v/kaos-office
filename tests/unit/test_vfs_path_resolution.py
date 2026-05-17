"""VFS-aware path resolution for every kaos-office file-input tool.

End-to-end coverage for the Stage-1 fix of the VFS-blind-tools audit
plan: files uploaded into ``KaosRuntime.vfs`` (the production layout
used by ``kaos-ui``'s single-user-chat SPA) must be visible to the
office tools when the agent passes a bare VFS path, not just absolute
filesystem paths.

For each of the three formats we:

1. Build a real Office file in-memory via the existing minimal
   builders in :mod:`tests.conftest`.
2. Write the bytes through the session-scoped VFS at a relative path
   (e.g. ``files/uploaded.docx``) — the same layout the SPA backend
   writes uploads to.
3. Call the matching tools with the bare VFS path (no
   ``kaos://`` prefix, no absolute path) and assert success.

Without the path-resolver routing the calls would all fail with
"File not found" because ``Path("files/uploaded.docx").exists()``
resolves against the test runner's CWD, not the session VFS.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_core import (
    ArtifactStore,
    KaosContext,
    KaosRuntime,
    KaosSettings,
    VFSConfig,
    VirtualFileSystem,
)
from kaos_core.types.enums import StorageBackend

from kaos_office.tools import (
    DocxMetadataTool,
    GetDocxMarkdownTool,
    GetDocxTextTool,
    GetSheetXlsxTool,
    GetSlideNotesTool,
    GetSlideTool,
    ListSheetsXlsxTool,
    ListSlidesTool,
    ParseDocxTool,
    ParsePptxTool,
    ParseXlsxTool,
    SearchDocxTool,
    SearchPptxTool,
    WritePptxTool,
    XlsxMetadataTool,
)
from tests.conftest import make_minimal_docx, make_minimal_pptx

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_runtime(tmp_path: Path) -> KaosRuntime:
    """Build a KaosRuntime with a real disk-backed VFS + artifact store.

    Mirrors the wiring used by ``kaos-mcp`` in production so the
    test exercises the same session-scoping boundary the SPA backend
    uses for uploads.
    """
    settings = KaosSettings(
        artifact_inline_read_max_bytes=262_144,
        artifact_chunk_size_bytes=65_536,
    )
    runtime = KaosRuntime(config=settings)
    runtime.vfs = VirtualFileSystem(
        VFSConfig(default_backend=StorageBackend.DISK, disk_base_path=tmp_path / "vfs")
    )
    runtime.artifacts = ArtifactStore(
        runtime.vfs,
        manifest_context_id=settings.artifact_manifest_context_id,
        manifest_prefix=settings.artifact_manifest_prefix,
        max_inline_read_bytes=settings.artifact_inline_read_max_bytes,
        default_chunk_size=settings.artifact_chunk_size_bytes,
        temporary_ttl_seconds=settings.artifact_temporary_ttl_seconds,
    )
    return runtime


def _context(runtime: KaosRuntime, *, session_id: str = "s-vfs-test") -> KaosContext:
    return KaosContext(session_id=session_id, runtime=runtime, vfs=runtime.vfs)


async def _upload_to_vfs(ctx: KaosContext, vfs_path: str, payload: bytes) -> None:
    """Write ``payload`` to the session-scoped VFS at ``vfs_path``.

    This is exactly the path the SPA upload endpoint takes — bytes go
    through ``context.get_vfs_path(...)`` so they land inside the
    per-session backend directory rather than the process CWD.
    """
    handle = ctx.get_vfs_path(vfs_path)
    await handle.write_bytes(payload)


def _make_minimal_xlsx() -> bytes:
    """Build a one-sheet XLSX in memory via the native lxml writer."""
    from kaos_content.model.tabular import Column, ColumnType, Table, TabularDocument

    from kaos_office.xlsx.writer import write_xlsx_bytes

    doc = TabularDocument(
        tables=(
            Table(
                name="Sheet1",
                columns=(
                    Column(name="id", column_type=ColumnType.INTEGER),
                    Column(name="label", column_type=ColumnType.TEXT),
                ),
                rows=((1, "alpha"), (2, "beta")),
            ),
        ),
    )
    return write_xlsx_bytes(doc)


# ---------------------------------------------------------------------------
# DOCX
# ---------------------------------------------------------------------------


@pytest.fixture
async def docx_vfs_context(tmp_path: Path):
    """Session context with a DOCX uploaded at a bare relative VFS path."""
    runtime = _make_runtime(tmp_path)
    ctx = _context(runtime, session_id="s-docx-vfs")
    vfs_path = "files/uploaded.docx"
    await _upload_to_vfs(ctx, vfs_path, make_minimal_docx())
    return ctx, vfs_path


class TestDocxToolsResolveVfsPaths:
    async def test_parse_docx_resolves_vfs_path(self, docx_vfs_context) -> None:
        ctx, vfs_path = docx_vfs_context
        result = await ParseDocxTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_get_text_resolves_vfs_path(self, docx_vfs_context) -> None:
        ctx, vfs_path = docx_vfs_context
        result = await GetDocxTextTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True
        assert "Hello" in str(result.content)

    async def test_get_markdown_resolves_vfs_path(self, docx_vfs_context) -> None:
        ctx, vfs_path = docx_vfs_context
        result = await GetDocxMarkdownTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True
        assert "Hello" in str(result.content)

    async def test_metadata_resolves_vfs_path(self, docx_vfs_context) -> None:
        ctx, vfs_path = docx_vfs_context
        result = await DocxMetadataTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_search_resolves_vfs_path(self, docx_vfs_context) -> None:
        ctx, vfs_path = docx_vfs_context
        result = await SearchDocxTool().execute({"path": vfs_path, "query": "Hello"}, context=ctx)
        assert result.isError is not True


# ---------------------------------------------------------------------------
# PPTX
# ---------------------------------------------------------------------------


@pytest.fixture
async def pptx_vfs_context(tmp_path: Path):
    """Session context with a PPTX uploaded at a bare relative VFS path."""
    runtime = _make_runtime(tmp_path)
    ctx = _context(runtime, session_id="s-pptx-vfs")
    vfs_path = "files/uploaded.pptx"
    await _upload_to_vfs(ctx, vfs_path, make_minimal_pptx())
    return ctx, vfs_path


class TestPptxToolsResolveVfsPaths:
    async def test_parse_pptx_resolves_vfs_path(self, pptx_vfs_context) -> None:
        ctx, vfs_path = pptx_vfs_context
        result = await ParsePptxTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_list_slides_resolves_vfs_path(self, pptx_vfs_context) -> None:
        ctx, vfs_path = pptx_vfs_context
        result = await ListSlidesTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_get_slide_resolves_vfs_path(self, pptx_vfs_context) -> None:
        ctx, vfs_path = pptx_vfs_context
        result = await GetSlideTool().execute({"path": vfs_path, "slide_number": 1}, context=ctx)
        assert result.isError is not True
        assert "Test Title" in str(result.content)

    async def test_search_pptx_resolves_vfs_path(self, pptx_vfs_context) -> None:
        ctx, vfs_path = pptx_vfs_context
        result = await SearchPptxTool().execute({"path": vfs_path, "query": "Test"}, context=ctx)
        assert result.isError is not True

    async def test_get_slide_notes_resolves_vfs_path(self, pptx_vfs_context) -> None:
        ctx, vfs_path = pptx_vfs_context
        result = await GetSlideNotesTool().execute({"path": vfs_path, "slide": 1}, context=ctx)
        # No notes attached; tool returns success with a "no speaker notes"
        # message — what matters is that the file was *found*.
        assert result.isError is not True


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------


@pytest.fixture
async def xlsx_vfs_context(tmp_path: Path):
    """Session context with an XLSX uploaded at a bare relative VFS path."""
    runtime = _make_runtime(tmp_path)
    ctx = _context(runtime, session_id="s-xlsx-vfs")
    vfs_path = "files/uploaded.xlsx"
    await _upload_to_vfs(ctx, vfs_path, _make_minimal_xlsx())
    return ctx, vfs_path


class TestXlsxToolsResolveVfsPaths:
    async def test_parse_xlsx_resolves_vfs_path(self, xlsx_vfs_context) -> None:
        ctx, vfs_path = xlsx_vfs_context
        result = await ParseXlsxTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_list_sheets_resolves_vfs_path(self, xlsx_vfs_context) -> None:
        ctx, vfs_path = xlsx_vfs_context
        result = await ListSheetsXlsxTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True

    async def test_get_sheet_resolves_vfs_path(self, xlsx_vfs_context) -> None:
        ctx, vfs_path = xlsx_vfs_context
        result = await GetSheetXlsxTool().execute(
            {"path": vfs_path, "sheet": "Sheet1"}, context=ctx
        )
        assert result.isError is not True

    async def test_metadata_resolves_vfs_path(self, xlsx_vfs_context) -> None:
        ctx, vfs_path = xlsx_vfs_context
        result = await XlsxMetadataTool().execute({"path": vfs_path}, context=ctx)
        assert result.isError is not True


# ---------------------------------------------------------------------------
# WritePptxTool — optional template_path also flows through the resolver.
# ---------------------------------------------------------------------------


class TestWritePptxTemplateResolvesVfsPath:
    async def test_template_path_accepts_vfs_path(self, tmp_path: Path) -> None:
        runtime = _make_runtime(tmp_path)
        ctx = _context(runtime, session_id="s-template-vfs")
        template_vfs = "templates/brand.pptx"
        await _upload_to_vfs(ctx, template_vfs, make_minimal_pptx())

        # Build a trivial inline ContentDocument so the writer has body.
        from kaos_content.model.document import ContentDocument
        from kaos_content.model.metadata import DocumentMetadata

        doc = ContentDocument(metadata=DocumentMetadata(title="t"), body=())
        out = tmp_path / "out.pptx"

        result = await WritePptxTool().execute(
            {
                "document_json": doc.model_dump_json(),
                "output_path": str(out),
                "template_path": template_vfs,
                "force": True,
            },
            context=ctx,
        )
        # The template only needs to *resolve*; whether python-pptx actually
        # writes successfully depends on the optional dep being installed.
        # If the writer dep is missing we accept that error; what we're
        # asserting is that the path resolver did NOT fail with
        # "Template not found".
        if result.isError:
            assert "Template not found" not in str(result.content)
