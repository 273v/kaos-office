"""Unit tests for the CompareDocxTool MCP tool (DOCX redline)."""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text

from kaos_office.docx.writer import write_docx
from kaos_office.tools import CompareDocxTool


def _make_docx(path: Path, *paras: str) -> Path:
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=tuple(Paragraph(children=(Text(value=p),)) for p in paras),
    )
    write_docx(doc, path)
    return path


@pytest.fixture
def doc_pair(tmp_path: Path) -> tuple[Path, Path]:
    original = _make_docx(tmp_path / "a.docx", "Shared opening.", "Old middle line.")
    revised = _make_docx(tmp_path / "b.docx", "Shared opening.", "New middle line.")
    return original, revised


class TestCompareDocxTool:
    @pytest.mark.asyncio
    async def test_writes_redline_and_reports_revisions(
        self, doc_pair: tuple[Path, Path], tmp_path: Path
    ) -> None:
        original, revised = doc_pair
        out = tmp_path / "redline.docx"
        tool = CompareDocxTool()
        result = await tool.execute(
            {
                "original_path": original.as_uri(),
                "revised_path": revised.as_uri(),
                "output_path": str(out),
                "author": "Counsel",
            }
        )
        assert result.isError is False, result.content
        assert out.exists() and out.stat().st_size > 0
        structured = result.require_structured()
        assert structured["format"] == "docx"
        assert structured["revision_count"] >= 1
        assert isinstance(structured["revisions_by_type"], dict)

    @pytest.mark.asyncio
    async def test_missing_inputs_returns_error(self, tmp_path: Path) -> None:
        tool = CompareDocxTool()
        result = await tool.execute({"output_path": str(tmp_path / "x.docx")})
        assert result.isError is True

    @pytest.mark.asyncio
    async def test_refuses_overwrite_without_force(
        self, doc_pair: tuple[Path, Path], tmp_path: Path
    ) -> None:
        original, revised = doc_pair
        out = tmp_path / "exists.docx"
        out.write_bytes(b"placeholder")
        tool = CompareDocxTool()
        result = await tool.execute(
            {
                "original_path": original.as_uri(),
                "revised_path": revised.as_uri(),
                "output_path": str(out),
            }
        )
        assert result.isError is True

    def test_metadata_is_authoring_transform(self) -> None:
        meta = CompareDocxTool().metadata
        assert meta.name == "kaos-office-redline-docx"
        assert {p.name for p in meta.input_schema} >= {
            "original_path",
            "revised_path",
            "output_path",
        }
