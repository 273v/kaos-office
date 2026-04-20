"""Tests for Phase 6.2 — DOCX Structured Document Tag preservation.

Pre-6.2 the reader unwrapped ``<w:sdt>`` silently, dropping tag /
alias / lock / control-type on every round-trip. Word templates with
form controls, content controls, or protected regions degenerated to
plain text after one KAOS pass.

This file pins the new preservation contract:
- Reader emits ``Div(classes=("sdt",), attr=Attr(kv=...))`` for
  block-level SDT and ``Span(classes=("sdt",), ...)`` for inline.
- Writer detects the class and re-emits the ``<w:sdt>`` wrapper with
  preserved ``<w:sdtPr>`` metadata.
- Round-trip (parse → write → parse) preserves ``sdt.tag`` / ``sdt.alias``
  / ``sdt.lock`` / ``sdt.control_type`` verbatim.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from kaos_content.model.blocks import Div, Paragraph
from kaos_content.model.inlines import Span
from kaos_content.traversal import find_by_type
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx
from kaos_office.ooxml.namespace import W, qn

# --- Helpers: hand-author a DOCX whose body contains specific SDTs ---------


_STYLES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)
_CT = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    b'<Default Extension="rels" ContentType='
    b'"application/vnd.openxmlformats-package.relationships+xml"/>'
    b'<Default Extension="xml" ContentType="application/xml"/>'
    b'<Override PartName="/word/document.xml" ContentType='
    b'"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    b'<Override PartName="/word/styles.xml" ContentType='
    b'"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    b"</Types>"
)
_ROOT_RELS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    b'Target="word/document.xml"/></Relationships>'
)
_DOC_RELS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    b'<Relationship Id="rId1" '
    b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
    b'Target="styles.xml"/></Relationships>'
)


def _pack_docx(body_inner_xml: str, tmp_path: Path, name: str = "sdt.docx") -> Path:
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<w:body>{body_inner_xml}</w:body></w:document>"
    ).encode()
    path = tmp_path / name
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/styles.xml", _STYLES_XML)
    return path


def _block_sdt(*, tag: str, alias: str, lock: str, control: str, text: str) -> str:
    """A minimal block-level SDT wrapping one paragraph."""
    return (
        "<w:sdt>"
        "<w:sdtPr>"
        f'<w:alias w:val="{alias}"/>'
        f'<w:tag w:val="{tag}"/>'
        '<w:id w:val="1234567"/>'
        f'<w:lock w:val="{lock}"/>'
        f"<w:{control}/>"
        "</w:sdtPr>"
        "<w:sdtContent>"
        f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
        "</w:sdtContent>"
        "</w:sdt>"
    )


def _inline_sdt(*, tag: str, text: str) -> str:
    """A minimal inline-level SDT wrapping one run inside a paragraph."""
    return (
        "<w:p>"
        "<w:r><w:t>Before </w:t></w:r>"
        "<w:sdt>"
        "<w:sdtPr>"
        f'<w:tag w:val="{tag}"/>'
        '<w:id w:val="8675309"/>'
        "<w:text/>"
        "</w:sdtPr>"
        "<w:sdtContent>"
        f"<w:r><w:t>{text}</w:t></w:r>"
        "</w:sdtContent>"
        "</w:sdt>"
        "<w:r><w:t> after</w:t></w:r>"
        "</w:p>"
    )


# --- Reader tests ----------------------------------------------------------


class TestReaderPreservesBlockSDT:
    def test_block_sdt_emits_div_with_sdt_class(self, tmp_path: Path) -> None:
        body = _block_sdt(
            tag="client.name",
            alias="Client Name",
            lock="sdtContentLocked",
            control="text",
            text="Acme Corp",
        )
        doc = parse_docx(_pack_docx(body, tmp_path))
        divs = list(find_by_type(doc, Div))
        sdts = [d for d in divs if "sdt" in (d.attr.classes or ())]
        assert len(sdts) == 1
        sdt_div = sdts[0]
        kv = sdt_div.attr.kv
        assert kv["sdt.tag"] == "client.name"
        assert kv["sdt.alias"] == "Client Name"
        assert kv["sdt.lock"] == "sdtContentLocked"
        assert kv["sdt.control_type"] == "text"
        # Content survives unchanged.
        assert len(sdt_div.children) == 1
        inner = sdt_div.children[0]
        assert isinstance(inner, Paragraph)

    def test_block_sdt_with_no_metadata_falls_back_to_unwrap(self, tmp_path: Path) -> None:
        """An anonymous SDT with no tag/alias/lock/control_type has
        nothing worth preserving — reader unwraps (legacy behavior)."""
        body = (
            "<w:sdt>"
            "<w:sdtPr/>"
            "<w:sdtContent>"
            "<w:p><w:r><w:t>anon</w:t></w:r></w:p>"
            "</w:sdtContent>"
            "</w:sdt>"
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
        )
        doc = parse_docx(_pack_docx(body, tmp_path))
        divs = list(find_by_type(doc, Div))
        assert not [d for d in divs if "sdt" in (d.attr.classes or ())]
        # But the text still reached the body as a Paragraph.
        from kaos_content.traversal.visitor import extract_text

        paras = list(find_by_type(doc, Paragraph))
        assert any("anon" in extract_text(p) for p in paras)


class TestReaderPreservesInlineSDT:
    def test_inline_sdt_emits_span_with_sdt_class(self, tmp_path: Path) -> None:
        body = _inline_sdt(tag="match.clause", text="shall survive")
        doc = parse_docx(_pack_docx(body, tmp_path))
        spans = list(find_by_type(doc, Span))
        sdt_spans = [s for s in spans if "sdt" in (s.attr.classes or ())]
        assert len(sdt_spans) == 1
        kv = sdt_spans[0].attr.kv
        assert kv["sdt.tag"] == "match.clause"
        assert kv["sdt.control_type"] == "text"
        # The wrapped run's text survives.
        from kaos_content.traversal.visitor import extract_text

        assert "shall survive" in extract_text(sdt_spans[0])


# --- Writer + round-trip tests --------------------------------------------


class TestRoundTrip:
    """read → write → read preserves the SDT wrapper + metadata."""

    def test_block_sdt_roundtrip(self, tmp_path: Path) -> None:
        body = _block_sdt(
            tag="party.1.name",
            alias="Party 1",
            lock="contentLocked",
            control="richText",
            text="Contracting Party",
        )
        original_path = _pack_docx(body, tmp_path, name="original.docx")
        doc = parse_docx(original_path)

        rewritten = tmp_path / "rt.docx"
        write_docx(doc, rewritten)

        reloaded = parse_docx(rewritten)
        sdts = [d for d in find_by_type(reloaded, Div) if "sdt" in (d.attr.classes or ())]
        assert len(sdts) == 1
        kv = sdts[0].attr.kv
        assert kv["sdt.tag"] == "party.1.name"
        assert kv["sdt.alias"] == "Party 1"
        assert kv["sdt.lock"] == "contentLocked"
        assert kv["sdt.control_type"] == "richText"

        # Structural check: the rewritten DOCX actually contains <w:sdt>.
        with zipfile.ZipFile(rewritten) as zf:
            xml = zf.read("word/document.xml")
        root = etree.fromstring(xml)
        sdt_els = root.findall(f".//{qn(W, 'sdt')}")
        assert len(sdt_els) == 1
        sdt_pr = sdt_els[0].find(qn(W, "sdtPr"))
        assert sdt_pr is not None
        tag_el = sdt_pr.find(qn(W, "tag"))
        assert tag_el is not None and tag_el.get(qn(W, "val")) == "party.1.name"

    def test_inline_sdt_roundtrip(self, tmp_path: Path) -> None:
        body = _inline_sdt(tag="assignment.clause", text="the following clause")
        original_path = _pack_docx(body, tmp_path)
        doc = parse_docx(original_path)

        rewritten = tmp_path / "rt.docx"
        write_docx(doc, rewritten)

        reloaded = parse_docx(rewritten)
        sdt_spans = [s for s in find_by_type(reloaded, Span) if "sdt" in (s.attr.classes or ())]
        assert len(sdt_spans) == 1
        assert sdt_spans[0].attr.kv["sdt.tag"] == "assignment.clause"

    def test_multiple_block_sdts_preserve_order(self, tmp_path: Path) -> None:
        body = (
            _block_sdt(tag="a", alias="A", lock="unlocked", control="text", text="one")
            + "<w:p><w:r><w:t>between</w:t></w:r></w:p>"
            + _block_sdt(tag="b", alias="B", lock="unlocked", control="text", text="two")
        )
        doc = parse_docx(_pack_docx(body, tmp_path))
        sdts = [d for d in find_by_type(doc, Div) if "sdt" in (d.attr.classes or ())]
        assert [s.attr.kv["sdt.tag"] for s in sdts] == ["a", "b"]


class TestLibreOfficeAccepts:
    """The regenerated DOCX is still a valid OOXML file — LibreOffice
    --convert-to pdf is the cheapest independent validator we have."""

    def test_libreoffice_converts(self, tmp_path: Path) -> None:
        import shutil
        import subprocess

        if not shutil.which("libreoffice"):
            import pytest

            pytest.skip("libreoffice not installed")

        body = _block_sdt(
            tag="effective.date",
            alias="Effective Date",
            lock="sdtContentLocked",
            control="date",
            text="2026-04-20",
        )
        original_path = _pack_docx(body, tmp_path)
        doc = parse_docx(original_path)
        rt = tmp_path / "rt.docx"
        write_docx(doc, rt)

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                str(rt),
                "--outdir",
                str(tmp_path),
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
        assert result.returncode == 0, result.stderr
        pdf = tmp_path / "rt.pdf"
        assert pdf.exists() and pdf.stat().st_size > 1000
