"""Tests for CLI subcommands."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from kaos_office.cli import main
from tests.conftest import make_minimal_docx


@pytest.fixture
def docx_path(tmp_path: Path) -> str:
    path = tmp_path / "test.docx"
    core_xml = """\
<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>CLI Test Doc</dc:title>
  <dc:creator>Test Author</dc:creator>
</cp:coreProperties>"""
    path.write_bytes(make_minimal_docx(core_xml=core_xml))
    return str(path)


class TestExtractCommand:
    def test_extract_markdown(self, docx_path, capsys):
        main(["extract", docx_path])
        output = capsys.readouterr().out
        assert "Hello" in output

    def test_extract_text(self, docx_path, capsys):
        main(["extract", docx_path, "--format", "text"])
        output = capsys.readouterr().out
        assert "Hello" in output

    def test_extract_json_envelope(self, docx_path, capsys):
        main(["extract", docx_path, "--json"])
        output = capsys.readouterr().out
        data = json.loads(output)
        assert data["command"] == "extract"
        assert data["blocks"] >= 1

    def test_extract_to_file(self, docx_path, tmp_path):
        output_file = str(tmp_path / "output.md")
        main(["extract", docx_path, "--output", output_file])
        assert Path(output_file).exists()
        assert "Hello" in Path(output_file).read_text()

    def test_extract_file_not_found(self):
        with pytest.raises(SystemExit):
            main(["extract", "/nonexistent/file.docx"])


class TestSearchCommand:
    def test_search_basic(self, docx_path, capsys):
        main(["search", docx_path, "Hello"])
        output = capsys.readouterr().out
        assert "Hello" in output or "Found" in output

    def test_search_json(self, docx_path, capsys):
        main(["search", docx_path, "Hello", "--json"])
        output = capsys.readouterr().out
        data = json.loads(output)
        assert data["command"] == "search"
        assert "results" in data


class TestMetadataCommand:
    def test_metadata_human(self, docx_path, capsys):
        main(["metadata", docx_path])
        output = capsys.readouterr().out
        assert "CLI Test Doc" in output or "Test Author" in output

    def test_metadata_json(self, docx_path, capsys):
        main(["metadata", docx_path, "--json"])
        output = capsys.readouterr().out
        data = json.loads(output)
        assert data["command"] == "metadata"
        assert data.get("title") == "CLI Test Doc" or data.get("creator") == "Test Author"
