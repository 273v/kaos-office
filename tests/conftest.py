"""Shared test fixtures for kaos-office."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest

KELVIN_FIXTURES = Path(
    "/home/mjbommar/projects/273v/kelvin-modules/kelvin_office/tests/resources/docx"
)


def has_kelvin_fixtures() -> bool:
    """Check if kelvin_office test fixtures are available."""
    return KELVIN_FIXTURES.exists()


skip_no_fixtures = pytest.mark.skipif(
    not has_kelvin_fixtures(),
    reason="kelvin_office test fixtures not available",
)


def make_minimal_docx(
    *,
    body_xml: str = "<w:p><w:r><w:t>Hello</w:t></w:r></w:p>",
    styles_xml: str | None = None,
    numbering_xml: str | None = None,
    comments_xml: str | None = None,
    footnotes_xml: str | None = None,
    core_xml: str | None = None,
    app_xml: str | None = None,
) -> bytes:
    """Create a minimal valid DOCX file in memory.

    Args:
        body_xml: XML fragment for w:body content.
        styles_xml: Full styles.xml content, or None to omit.
        numbering_xml: Full numbering.xml content, or None to omit.
        comments_xml: Full comments.xml content, or None to omit.
        footnotes_xml: Full footnotes.xml content, or None to omit.
        core_xml: Full docProps/core.xml content, or None to omit.
        app_xml: Full docProps/app.xml content, or None to omit.

    Returns:
        Bytes of a valid DOCX ZIP file.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    document_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}">
  <w:body>
    {body_xml}
  </w:body>
</w:document>"""

    # Build relationship entries
    doc_rels_entries = []
    rel_id = 1

    if styles_xml:
        doc_rels_entries.append(
            f'<Relationship Id="rId{rel_id}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            f'Target="styles.xml"/>'
        )
        rel_id += 1

    if numbering_xml:
        doc_rels_entries.append(
            f'<Relationship Id="rId{rel_id}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" '
            f'Target="numbering.xml"/>'
        )
        rel_id += 1

    if comments_xml:
        doc_rels_entries.append(
            f'<Relationship Id="rId{rel_id}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
            f'Target="comments.xml"/>'
        )
        rel_id += 1

    if footnotes_xml:
        doc_rels_entries.append(
            f'<Relationship Id="rId{rel_id}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
            f'Target="footnotes.xml"/>'
        )
        rel_id += 1

    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    root_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  {"".join(doc_rels_entries)}
</Relationships>"""

    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    overrides = [
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    ]
    if styles_xml:
        overrides.append(
            '<Override PartName="/word/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        )
    if numbering_xml:
        overrides.append(
            '<Override PartName="/word/numbering.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        )

    content_types = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT_NS}">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  {"".join(overrides)}
</Types>"""

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)

        if styles_xml:
            zf.writestr("word/styles.xml", styles_xml)
        if numbering_xml:
            zf.writestr("word/numbering.xml", numbering_xml)
        if comments_xml:
            zf.writestr("word/comments.xml", comments_xml)
        if footnotes_xml:
            zf.writestr("word/footnotes.xml", footnotes_xml)
        if core_xml:
            zf.writestr("docProps/core.xml", core_xml)
        if app_xml:
            zf.writestr("docProps/app.xml", app_xml)

    return buf.getvalue()
