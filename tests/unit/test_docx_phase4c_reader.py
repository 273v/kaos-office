"""Tests for DOCX Phase 4C — reader populates ``doc.sections``.

The pre-4C reader read only the final ``<w:sectPr>`` and discarded
every non-terminal section's geometry. This was silent data loss for
multi-section DOCX documents (cover + landscape tables, per-region
headers, mid-document orientation changes). 4C extends the reader to
emit one ``Section`` per sectPr, in document order.

Backward compat: ``doc.metadata.page_setup`` continues to surface the
final section's geometry — existing consumers that read page_setup
keep working.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from kaos_office.docx.reader import parse_docx

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "docx"


# --- Helpers: hand-author a multi-section DOCX without python-docx ---------


_MIN_STYLES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)
# Long URIs deliberately: OOXML namespace constants. noqa locally.
_CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
_CT_DOC = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
_CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_RT_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
_RT_STYLES_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

_MIN_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    f'<Default Extension="rels" ContentType="{_CT_RELS}"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/word/document.xml" ContentType="{_CT_DOC}"/>'
    f'<Override PartName="/word/styles.xml" ContentType="{_CT_STYLES}"/>'
    "</Types>"
).encode()
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{_RT_OFFICE}" Target="word/document.xml"/>'
    "</Relationships>"
).encode()
_DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{_RT_STYLES_URI}" Target="styles.xml"/>'
    "</Relationships>"
).encode()


def _make_docx(body_inner_xml: str, tmp_path: Path, name: str = "multi.docx") -> Path:
    """Pack a minimal DOCX whose body XML is the supplied fragment.

    ``body_inner_xml`` goes directly between ``<w:body>`` tags and must
    include whatever paragraphs + sectPrs the test wants. Everything
    else is the smallest valid DOCX wrapper: one style sheet, the root
    rels, and the matching content types. This keeps test fixtures
    auditable line-by-line rather than opaque binary blobs.
    """
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<w:body>{body_inner_xml}</w:body>"
        "</w:document>"
    ).encode()

    path = tmp_path / name
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _MIN_CONTENT_TYPES)
        zf.writestr("_rels/.rels", _ROOT_RELS)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", _MIN_STYLES_XML)
    return path


def _p(text: str, sect_pr_inner: str | None = None) -> str:
    """Build a `<w:p>` with optional nested sectPr in its pPr."""
    ppr = ""
    if sect_pr_inner:
        ppr = f"<w:pPr><w:sectPr>{sect_pr_inner}</w:sectPr></w:pPr>"
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _sect_pr(*, w: int = 12240, h: int = 15840, type_val: str | None = None) -> str:
    """Inline sectPr geometry helper."""
    type_xml = f'<w:type w:val="{type_val}"/>' if type_val else ""
    pg_sz = f'<w:pgSz w:w="{w}" w:h="{h}"/>'
    pg_mar = (
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
        'w:left="1440" w:header="720" w:footer="720"/>'
    )
    return f"{type_xml}{pg_sz}{pg_mar}"


# --- Tests ----------------------------------------------------------------


class TestSingleSection:
    """Legacy shape: one body-direct sectPr, no nested. doc.sections has one
    entry. doc.metadata.page_setup continues to reflect that section."""

    def test_one_section(self, tmp_path: Path) -> None:
        body = _p("hello") + f"<w:sectPr>{_sect_pr()}</w:sectPr>"
        doc = parse_docx(_make_docx(body, tmp_path))
        assert len(doc.sections) == 1
        assert doc.sections[0].end_block_index == len(doc.body)
        assert doc.sections[0].end_block_index == 1
        assert doc.sections[0].page_setup is not None
        # Backward compat: page_setup still populated.
        assert doc.metadata.page_setup is not None

    def test_no_sect_pr_empty_sections(self, tmp_path: Path) -> None:
        """A body with no <w:sectPr> at all produces an empty sections tuple
        — the "implicit single section from metadata.page_setup" shape."""
        body = _p("just one paragraph")
        doc = parse_docx(_make_docx(body, tmp_path))
        assert doc.sections == ()


class TestMultiSection:
    """A body with N sectPrs produces N Sections in document order."""

    def test_two_sections(self, tmp_path: Path) -> None:
        # Section 1: portrait US Letter, closes at paragraph 2 (landscape
        # via nested sectPr). Section 2: landscape, body-direct sectPr.
        body = (
            _p("section 1 para 1")
            + _p("section 1 para 2", sect_pr_inner=_sect_pr(w=12240, h=15840))
            + _p("section 2 para 1")
            + _p("section 2 para 2")
            + f"<w:sectPr>{_sect_pr(w=15840, h=12240)}</w:sectPr>"
        )
        doc = parse_docx(_make_docx(body, tmp_path))
        assert len(doc.sections) == 2
        assert doc.sections[0].end_block_index == 2
        assert doc.sections[1].end_block_index == 4
        # Section 1 is portrait (w < h), section 2 is landscape (w > h).
        ps1 = doc.sections[0].page_setup
        ps2 = doc.sections[1].page_setup
        assert ps1 is not None and ps2 is not None
        assert ps1.page_width_pt is not None and ps1.page_height_pt is not None
        assert ps2.page_width_pt is not None and ps2.page_height_pt is not None
        assert ps1.page_width_pt < ps1.page_height_pt  # portrait
        assert ps2.page_width_pt > ps2.page_height_pt  # landscape

    def test_break_type_round_trips(self, tmp_path: Path) -> None:
        body = (
            _p(
                "cover",
                sect_pr_inner=_sect_pr(type_val="nextPage"),
            )
            + _p("body 1")
            + _p("body 2", sect_pr_inner=_sect_pr(type_val="continuous"))
            + _p("tail")
            + f"<w:sectPr>{_sect_pr(type_val='nextPage')}</w:sectPr>"
        )
        doc = parse_docx(_make_docx(body, tmp_path))
        assert len(doc.sections) == 3
        assert doc.sections[0].break_type == "nextPage"
        assert doc.sections[1].break_type == "continuous"
        assert doc.sections[2].break_type == "nextPage"

    def test_unknown_break_type_defaults_to_next_page(self, tmp_path: Path) -> None:
        """Readers must be forgiving: an unknown w:type/@w:val collapses
        to the OOXML default rather than raising."""
        body = (
            _p("x", sect_pr_inner=_sect_pr(type_val="madeUpType"))
            + f"<w:sectPr>{_sect_pr()}</w:sectPr>"
        )
        doc = parse_docx(_make_docx(body, tmp_path))
        assert doc.sections[0].break_type == "nextPage"


class TestBackwardCompatPageSetup:
    """Phase 4 consumers reading ``doc.metadata.page_setup`` must keep working.

    The contract is: page_setup surfaces the **final** section's geometry
    (the OOXML document-default value).
    """

    def test_final_section_wins_for_page_setup(self, tmp_path: Path) -> None:
        body = (
            _p("portrait", sect_pr_inner=_sect_pr(w=12240, h=15840))
            + _p("landscape body")
            + f"<w:sectPr>{_sect_pr(w=15840, h=12240)}</w:sectPr>"
        )
        doc = parse_docx(_make_docx(body, tmp_path))
        ps = doc.metadata.page_setup
        assert ps is not None
        assert ps.page_width_pt is not None and ps.page_height_pt is not None
        assert ps.page_width_pt > ps.page_height_pt  # landscape (final section)


class TestFixtures:
    """Smoke test against whichever fixture in tests/fixtures/docx/ has
    multiple sectPrs. Skipped if none exist."""

    def _find_multi_section_fixture(self) -> Path | None:
        for p in sorted(FIXTURES.glob("*.docx")):
            try:
                with zipfile.ZipFile(p) as zf:
                    xml = zf.read("word/document.xml")
            except (zipfile.BadZipFile, KeyError):
                continue
            # Cheap substring check — all sectPr variants match.
            if xml.count(b"<w:sectPr") >= 2 or xml.count(b":sectPr ") >= 2:
                return p
        return None

    def test_fixture_round_trip(self) -> None:
        fixture = self._find_multi_section_fixture()
        if fixture is None:
            pytest.skip("no multi-section fixture available")
        doc = parse_docx(fixture)
        # If the reader found multiple sectPrs, doc.sections is non-empty
        # and ordered; end_block_index strictly increases.
        if len(doc.sections) > 1:
            for a, b in zip(doc.sections, doc.sections[1:], strict=False):
                assert a.end_block_index <= b.end_block_index
