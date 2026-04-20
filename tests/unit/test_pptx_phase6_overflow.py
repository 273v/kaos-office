"""Tests for Phase 6.3 — PPTX text overflow handling.

Pre-6.3 the PPTX writer set no auto-size, so PowerPoint silently
cropped LLM-generated summary slides that outgrew their original
shape. The new ``overflow`` parameter gives callers three choices:

- ``"warn"`` (default) — no auto-size, emits a logger warning when
  character density suggests overflow. Never silently truncates
  because the warning is visible in logs.
- ``"autofit"`` — PowerPoint shrinks the font to fit.
- ``"extend"`` — PowerPoint grows the shape to fit.

These tests pin the public contract and verify the XML that reaches
PowerPoint carries the right auto-size values per mode.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_content.model.blocks import Heading, Paragraph
from kaos_content.model.document import ContentDocument
from kaos_content.model.inlines import Text

from kaos_office.pptx.writer import write_pptx, write_pptx_bytes


def _long_paragraph(char_count: int) -> Paragraph:
    """One paragraph with exactly char_count characters of text."""
    return Paragraph(children=(Text(value="x" * char_count),))


def _short_doc() -> ContentDocument:
    return ContentDocument(
        body=(
            Heading(depth=1, children=(Text(value="Title"),)),
            Paragraph(children=(Text(value="Short body."),)),
        )
    )


def _overflow_doc(chars: int = 7000) -> ContentDocument:
    """Doc whose body text is guaranteed to exceed the warn threshold.

    Default body shape is 9x5 inches = 45 sq in. Threshold is 150
    chars/sq in = 6750 chars. A 7000-char payload should warn.
    """
    return ContentDocument(
        body=(
            Heading(depth=1, children=(Text(value="Overflow test"),)),
            _long_paragraph(chars),
        )
    )


# --- Overflow mode XML contract -------------------------------------------


class TestAutofitMode:
    """``overflow="autofit"`` writes ``normAutofit`` into the shape."""

    def test_autofit_emits_norm_autofit_element(self, tmp_path: Path) -> None:
        doc = _overflow_doc()
        out = tmp_path / "autofit.pptx"
        write_pptx(doc, out, overflow="autofit")

        import zipfile

        with zipfile.ZipFile(out) as zf:
            slide_xml = zf.read("ppt/slides/slide1.xml").decode()
        # python-pptx maps MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE to
        # ``<a:normAutofit/>`` on the text frame's bodyPr.
        assert "normAutofit" in slide_xml, (
            "autofit mode must emit <a:normAutofit/> so PowerPoint shrinks font"
        )


class TestExtendMode:
    """``overflow="extend"`` writes ``spAutoFit`` into the shape."""

    def test_extend_emits_sp_auto_fit(self, tmp_path: Path) -> None:
        doc = _overflow_doc()
        out = tmp_path / "extend.pptx"
        write_pptx(doc, out, overflow="extend")

        import zipfile

        with zipfile.ZipFile(out) as zf:
            slide_xml = zf.read("ppt/slides/slide1.xml").decode()
        assert "spAutoFit" in slide_xml, (
            "extend mode must emit <a:spAutoFit/> so PowerPoint grows shape"
        )


class TestWarnMode:
    """``overflow="warn"`` (default) emits neither auto-size element
    and logs a structured warning when text density is high."""

    def test_warn_emits_no_autosize_element(self, tmp_path: Path) -> None:
        doc = _short_doc()
        out = tmp_path / "warn_short.pptx"
        write_pptx(doc, out)  # default

        import zipfile

        with zipfile.ZipFile(out) as zf:
            slide_xml = zf.read("ppt/slides/slide1.xml").decode()
        assert "normAutofit" not in slide_xml
        assert "spAutoFit" not in slide_xml

    def test_warn_logs_when_density_exceeds_threshold(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """kaos-core's logging bootstrap installs a stderr handler with
        ``propagate=False`` on the ``kaos`` logger, so caplog/capsys/
        capfd all miss it. Patch the module logger's ``warning`` method
        directly and capture call args — the simplest reliable signal."""
        from kaos_office.pptx import writer as pptx_writer

        captured: list[tuple[str, tuple[object, ...]]] = []

        def _spy(msg: str, *args: object) -> None:
            captured.append((msg, args))

        monkeypatch.setattr(pptx_writer.logger, "warning", _spy)

        doc = _overflow_doc(chars=8000)
        out = tmp_path / "warn_overflow.pptx"
        write_pptx(doc, out)
        assert any("likely overflows" in msg for msg, _ in captured), (
            f"expected overflow warning, got: {captured!r}"
        )

    def test_warn_silent_below_threshold(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from kaos_office.pptx import writer as pptx_writer

        captured: list[tuple[str, tuple[object, ...]]] = []

        def _spy(msg: str, *args: object) -> None:
            captured.append((msg, args))

        monkeypatch.setattr(pptx_writer.logger, "warning", _spy)

        doc = _short_doc()
        out = tmp_path / "warn_clean.pptx"
        write_pptx(doc, out)
        assert not any("likely overflows" in msg for msg, _ in captured), (
            f"short doc should not warn: {captured!r}"
        )


class TestDefaultIsWarn:
    """The default value of ``overflow`` is ``"warn"``, not silent."""

    def test_default_matches_explicit_warn(self, tmp_path: Path) -> None:
        doc = _short_doc()
        a = write_pptx_bytes(doc)  # default
        b = write_pptx_bytes(doc, overflow="warn")
        # Byte-identical except for ZIP timestamps / ordering; the
        # embedded slide XML must be the same shape.
        import io
        import zipfile

        def _slide(data: bytes) -> str:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                return zf.read("ppt/slides/slide1.xml").decode()

        assert _slide(a) == _slide(b)


class TestFileRoundTrip:
    """Written file opens cleanly — LibreOffice rendering confirms the
    autosize XML is valid. Skipped when libreoffice is absent."""

    def test_autofit_opens_in_libreoffice(self, tmp_path: Path) -> None:
        import shutil
        import subprocess

        if not shutil.which("libreoffice"):
            pytest.skip("libreoffice not installed")

        doc = _overflow_doc(chars=8000)
        out = tmp_path / "auto.pptx"
        write_pptx(doc, out, overflow="autofit")
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                str(out),
                "--outdir",
                str(tmp_path),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        assert result.returncode == 0, result.stderr
        assert (tmp_path / "auto.pdf").exists()
