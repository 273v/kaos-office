"""Unit tests for PPTX CLI commands."""

from __future__ import annotations

import json

import pytest

from tests.conftest import make_minimal_pptx


@pytest.fixture
def pptx_file(tmp_path):
    """Create a temp PPTX file for testing."""
    data = make_minimal_pptx()
    path = tmp_path / "test.pptx"
    path.write_bytes(data)
    return path


class TestPptxExtractCommand:
    """Test pptx-extract subcommand."""

    def test_extract_markdown(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-extract", str(pptx_file)])
        out = capsys.readouterr().out
        assert "Test Title" in out

    def test_extract_text(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-extract", str(pptx_file), "--format", "text"])
        out = capsys.readouterr().out
        assert "Test Title" in out

    def test_extract_json_envelope(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-extract", str(pptx_file), "--json"])
        out = capsys.readouterr().out
        envelope = json.loads(out)
        assert envelope["command"] == "pptx-extract"
        assert "slides" in envelope
        assert "content" in envelope

    def test_extract_to_file(self, pptx_file, tmp_path, capsys):
        from kaos_office.cli import main

        outfile = tmp_path / "output.md"
        main(["pptx-extract", str(pptx_file), "--output", str(outfile)])
        assert outfile.exists()
        assert "Test Title" in outfile.read_text()

    def test_extract_not_found(self, capsys):
        from kaos_office.cli import main

        with pytest.raises(SystemExit):
            main(["pptx-extract", "/nonexistent.pptx"])


class TestPptxSlidesCommand:
    """Test pptx-slides subcommand."""

    def test_list_slides(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-slides", str(pptx_file)])
        out = capsys.readouterr().out
        assert "Slides: 1" in out
        assert "Test Title" in out

    def test_list_slides_json(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-slides", str(pptx_file), "--json"])
        out = capsys.readouterr().out
        envelope = json.loads(out)
        assert envelope["command"] == "pptx-slides"
        assert envelope["slide_count"] == 1
        assert len(envelope["slides"]) == 1


class TestPptxSlideCommand:
    """Test pptx-slide subcommand."""

    def test_get_slide(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-slide", str(pptx_file), "1"])
        out = capsys.readouterr().out
        assert "Test Title" in out

    def test_get_slide_json(self, pptx_file, capsys):
        from kaos_office.cli import main

        main(["pptx-slide", str(pptx_file), "1", "--json"])
        out = capsys.readouterr().out
        envelope = json.loads(out)
        assert envelope["command"] == "pptx-slide"
        assert envelope["slide_number"] == 1
        assert "Test Title" in envelope["content"]

    def test_get_slide_out_of_range(self, pptx_file, capsys):
        from kaos_office.cli import main

        with pytest.raises(SystemExit):
            main(["pptx-slide", str(pptx_file), "99"])
