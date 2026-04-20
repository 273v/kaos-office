"""Tests for DOCX Phase 4B.1 — titlePg + evenAndOddHeaders gating.

Phase 4 shipped header/footer reference emission with the ``w:type``
attribute for ``"first"`` and ``"even"`` variants, but neither of the
gating elements was emitted:

- ``<w:titlePg/>`` in sectPr gates ``w:type="first"`` — without it, Word
  silently ignores the first-page header/footer reference.
- ``<w:evenAndOddHeaders/>`` in word/settings.xml is the document-wide
  gate for ``w:type="even"`` — without it, Word uses the default header
  on every page regardless of parity.

These tests lock both gates in.
"""

from __future__ import annotations

import io
import zipfile

from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument
from kaos_content.model.inlines import Text
from lxml import etree  # ty: ignore[unresolved-import]

from kaos_office.docx.writer import write_docx_bytes
from kaos_office.ooxml.namespace import (
    CT_SETTINGS,
    W_BODY,
    W_EVEN_AND_ODD_HEADERS,
    W_SECTPR,
    W_TITLEPG,
)


def _body_sectpr(docx_bytes: bytes) -> etree._Element:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        xml = zf.read("word/document.xml")
    root = etree.fromstring(xml)
    body = root.find(W_BODY)
    assert body is not None
    sect_pr = body.find(W_SECTPR)
    assert sect_pr is not None
    return sect_pr


def _part(docx_bytes: bytes, name: str) -> str | None:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        if name not in zf.namelist():
            return None
        return zf.read(name).decode("utf-8")


class TestTitlePgGate:
    """`<w:titlePg/>` appears in sectPr iff a first header/footer exists."""

    def test_emitted_when_first_header_present(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={
                "default": (Paragraph(children=(Text(value="D"),)),),
                "first": (Paragraph(children=(Text(value="F"),)),),
            },
        )
        sect_pr = _body_sectpr(write_docx_bytes(doc))
        assert sect_pr.find(W_TITLEPG) is not None

    def test_emitted_when_first_footer_present(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            footers={"first": (Paragraph(children=(Text(value="F"),)),)},
        )
        sect_pr = _body_sectpr(write_docx_bytes(doc))
        assert sect_pr.find(W_TITLEPG) is not None

    def test_absent_when_only_default_header(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={"default": (Paragraph(children=(Text(value="D"),)),)},
        )
        sect_pr = _body_sectpr(write_docx_bytes(doc))
        assert sect_pr.find(W_TITLEPG) is None

    def test_absent_when_no_headers(self) -> None:
        doc = ContentDocument(body=(Paragraph(children=(Text(value="body"),)),))
        sect_pr = _body_sectpr(write_docx_bytes(doc))
        assert sect_pr.find(W_TITLEPG) is None


class TestEvenAndOddHeadersGate:
    """`<w:evenAndOddHeaders/>` lives in word/settings.xml (document-wide)."""

    def test_settings_emitted_when_even_header_present(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={
                "default": (Paragraph(children=(Text(value="D"),)),),
                "even": (Paragraph(children=(Text(value="E"),)),),
            },
        )
        data = write_docx_bytes(doc)
        settings_xml = _part(data, "word/settings.xml")
        assert settings_xml is not None, "settings.xml must exist when an even header is present"
        root = etree.fromstring(settings_xml.encode("utf-8"))
        assert root.find(W_EVEN_AND_ODD_HEADERS) is not None

    def test_settings_emitted_when_even_footer_present(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            footers={"even": (Paragraph(children=(Text(value="E"),)),)},
        )
        data = write_docx_bytes(doc)
        settings_xml = _part(data, "word/settings.xml")
        assert settings_xml is not None
        root = etree.fromstring(settings_xml.encode("utf-8"))
        assert root.find(W_EVEN_AND_ODD_HEADERS) is not None

    def test_settings_absent_when_no_even_variant(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={"default": (Paragraph(children=(Text(value="D"),)),)},
        )
        data = write_docx_bytes(doc)
        assert _part(data, "word/settings.xml") is None

    def test_settings_content_type_and_rel_registered(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={"even": (Paragraph(children=(Text(value="E"),)),)},
        )
        data = write_docx_bytes(doc)
        ct_xml = _part(data, "[Content_Types].xml")
        rels_xml = _part(data, "word/_rels/document.xml.rels")
        assert ct_xml is not None and CT_SETTINGS in ct_xml
        assert rels_xml is not None and "relationships/settings" in rels_xml


class TestCombinedGates:
    """Title-page + odd/even together is a common legal-document layout."""

    def test_both_gates_emitted(self) -> None:
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={
                "default": (Paragraph(children=(Text(value="D"),)),),
                "first": (Paragraph(children=(Text(value="F"),)),),
                "even": (Paragraph(children=(Text(value="E"),)),),
            },
            footers={
                "default": (Paragraph(children=(Text(value="DF"),)),),
                "first": (Paragraph(children=(Text(value="FF"),)),),
                "even": (Paragraph(children=(Text(value="EF"),)),),
            },
        )
        data = write_docx_bytes(doc)
        sect_pr = _body_sectpr(data)
        assert sect_pr.find(W_TITLEPG) is not None
        settings_xml = _part(data, "word/settings.xml")
        assert settings_xml is not None
        root = etree.fromstring(settings_xml.encode("utf-8"))
        assert root.find(W_EVEN_AND_ODD_HEADERS) is not None

    def test_titlepg_appears_after_references(self) -> None:
        """ECMA-376 §17.6.17 requires sectPr child order:
        headerReference / footerReference / titlePg / pgSz / pgMar.
        """
        doc = ContentDocument(
            body=(Paragraph(children=(Text(value="body"),)),),
            headers={
                "default": (Paragraph(children=(Text(value="D"),)),),
                "first": (Paragraph(children=(Text(value="F"),)),),
            },
        )
        sect_pr = _body_sectpr(write_docx_bytes(doc))
        # Walk direct children and capture local tag names in order.
        tags = [etree.QName(child).localname for child in sect_pr]
        # Every headerReference must come before titlePg.
        titlepg_idx = tags.index("titlePg")
        for i, t in enumerate(tags):
            if t in {"headerReference", "footerReference"}:
                assert i < titlepg_idx, f"{t} at {i} must precede titlePg at {titlepg_idx}"
        # titlePg must come before pgSz.
        if "pgSz" in tags:
            assert titlepg_idx < tags.index("pgSz")
