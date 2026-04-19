"""Tests for the write-docx / write-pptx / write-xlsx CLI subcommands."""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import pytest
from kaos_content.model.blocks import Heading, Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.model.tabular import Column, ColumnType, TabularDocument
from kaos_content.model.tabular import Table as TabTable

from kaos_office.cli import main


def _write_content_doc_json(tmp_path: Path, name: str = "doc.json") -> Path:
    doc = ContentDocument(
        metadata=DocumentMetadata(title="CLI"),
        body=(
            Heading(depth=1, children=(Text(value="Title"),)),
            Paragraph(children=(Text(value="Body text."),)),
        ),
    )
    p = tmp_path / name
    p.write_text(doc.model_dump_json(), encoding="utf-8")
    return p


def _write_tabular_doc_json(tmp_path: Path, name: str = "tab.json") -> Path:
    doc = TabularDocument(
        tables=(
            TabTable(
                name="Sheet1",
                columns=(
                    Column(name="id", column_type=ColumnType.INTEGER),
                    Column(name="label", column_type=ColumnType.TEXT),
                ),
                rows=((1, "alpha"), (2, "beta")),
            ),
        )
    )
    p = tmp_path / name
    p.write_text(doc.model_dump_json(), encoding="utf-8")
    return p


class TestWriteDocxCommand:
    def test_writes_file(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "out.docx"
        main(["write-docx", str(src), str(out)])
        assert out.exists()
        with zipfile.ZipFile(out) as zf:
            assert "word/document.xml" in zf.namelist()

    def test_json_envelope(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "out.docx"
        main(["write-docx", str(src), str(out), "--json"])
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["command"] == "write-docx"
        assert envelope["format"] == "docx"
        assert envelope["block_count"] == 2
        assert envelope["size_bytes"] > 0

    def test_refuses_overwrite_without_force(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "existing.docx"
        out.write_bytes(b"existing")
        with pytest.raises(SystemExit):
            main(["write-docx", str(src), str(out)])
        assert out.read_bytes() == b"existing"
        err = capsys.readouterr().err
        assert "overwrite" in err.lower()

    def test_overwrite_with_force(self, tmp_path: Path) -> None:
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "existing.docx"
        out.write_bytes(b"stub")
        main(["write-docx", str(src), str(out), "--force"])
        # Real DOCX is well over 100 bytes
        assert out.stat().st_size > 100

    def test_stdin_input(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        doc = ContentDocument(
            metadata=DocumentMetadata(title="Stdin"),
            body=(Paragraph(children=(Text(value="from stdin"),)),),
        )
        monkeypatch.setattr("sys.stdin", io.StringIO(doc.model_dump_json()))
        out = tmp_path / "out.docx"
        main(["write-docx", "-", str(out)])
        assert out.exists()
        assert out.stat().st_size > 100


class TestWriteXlsxCommand:
    def test_writes_xlsx(self, tmp_path: Path) -> None:
        src = _write_tabular_doc_json(tmp_path)
        out = tmp_path / "out.xlsx"
        main(["write-xlsx", str(src), str(out)])
        assert out.exists()
        with zipfile.ZipFile(out) as zf:
            assert "xl/workbook.xml" in zf.namelist()

    def test_json_envelope(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        src = _write_tabular_doc_json(tmp_path)
        out = tmp_path / "out.xlsx"
        main(["write-xlsx", str(src), str(out), "--json"])
        envelope = json.loads(capsys.readouterr().out)
        assert envelope["command"] == "write-xlsx"
        assert envelope["format"] == "xlsx"
        assert envelope["table_count"] == 1


class TestWritePptxCommand:
    def test_writes_pptx(self, tmp_path: Path) -> None:
        pytest.importorskip("pptx")
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "out.pptx"
        main(["write-pptx", str(src), str(out)])
        assert out.exists()
        with zipfile.ZipFile(out) as zf:
            assert "ppt/presentation.xml" in zf.namelist()

    def test_missing_template_raises(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        pytest.importorskip("pptx")
        src = _write_content_doc_json(tmp_path)
        out = tmp_path / "out.pptx"
        with pytest.raises(SystemExit):
            main(
                [
                    "write-pptx",
                    str(src),
                    str(out),
                    "--template",
                    str(tmp_path / "nope.pptx"),
                ]
            )
        err = capsys.readouterr().err
        assert "not found" in err.lower()
