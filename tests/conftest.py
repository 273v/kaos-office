"""Shared test fixtures for kaos-office.

Fixture roots (resolution order):

* ``tests/fixtures/{docx,pptx,xlsx}/`` — vendored real documents that
  ship with this repo. CI and any contributor's checkout always sees
  them, so the integration suite is reproducible without opt-in.
* ``KAOS_OFFICE_EXTERNAL_FIXTURES_DIR`` env var — optional override for
  fixtures too large to vendor (e.g. multi-MB legal decks). If set, it
  must point at a directory shaped like the vendored layout
  (``<root>/docx/``, ``<root>/pptx/``, ...). When unset, tests
  depending on these large fixtures auto-skip.

The legacy ``KELVIN_FIXTURES`` / ``KELVIN_PPTX_FIXTURES`` symbols are
preserved as aliases for the local fixture directories so existing
test imports keep working without a sweep.
"""

from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

import pytest

# --- Local (vendored) fixture roots -----------------------------------------

_LOCAL_FIXTURES_ROOT = Path(__file__).parent / "fixtures"
LOCAL_DOCX_FIXTURES = _LOCAL_FIXTURES_ROOT / "docx"
LOCAL_PPTX_FIXTURES = _LOCAL_FIXTURES_ROOT / "pptx"
LOCAL_XLSX_FIXTURES = _LOCAL_FIXTURES_ROOT / "xlsx"
PPTX_STRESS_FIXTURES = LOCAL_PPTX_FIXTURES / "stress"

# Backward-compatible aliases — kept so the integration tests don't
# need a rename sweep. Both now resolve to the in-repo vendored copies
# (the same files we used to read from kelvin-modules), not to an
# absolute path on a developer machine.
KELVIN_FIXTURES = LOCAL_DOCX_FIXTURES
KELVIN_PPTX_FIXTURES = LOCAL_PPTX_FIXTURES


# --- External (opt-in) fixture root -----------------------------------------

_EXTERNAL_ROOT_ENV = "KAOS_OFFICE_EXTERNAL_FIXTURES_DIR"


def external_fixtures_root() -> Path | None:
    """Return the configured external fixtures root, if any.

    Set via the ``KAOS_OFFICE_EXTERNAL_FIXTURES_DIR`` env var. Used for
    real documents too large to vendor in-repo (multi-MB decks /
    contracts). Returns ``None`` when the env var is unset, so callers
    can use ``pytest.skip(...)`` cleanly.
    """
    raw = os.environ.get(_EXTERNAL_ROOT_ENV)
    if not raw:
        return None
    root = Path(raw).expanduser()
    return root if root.is_dir() else None


def external_fixture(*parts: str) -> Path | None:
    """Resolve a path under the external fixtures root, or ``None``.

    ``parts`` is the relative path within the root (e.g. ``"pptx",
    "CIPLA_CLEVELAND_BAR_DEC_2023.pptx"``).
    """
    root = external_fixtures_root()
    if root is None:
        return None
    candidate = root.joinpath(*parts)
    return candidate if candidate.is_file() else None


# --- Skip helpers -----------------------------------------------------------


def has_local_docx_fixtures() -> bool:
    return LOCAL_DOCX_FIXTURES.is_dir() and any(LOCAL_DOCX_FIXTURES.glob("*.docx"))


def has_local_pptx_fixtures() -> bool:
    return LOCAL_PPTX_FIXTURES.is_dir() and any(LOCAL_PPTX_FIXTURES.glob("*.pptx"))


# Backward-compatible names — same shape as before, but the predicate
# now checks the in-repo vendored fixtures, not a hard-coded user path.
def has_kelvin_fixtures() -> bool:
    return has_local_docx_fixtures()


def has_kelvin_pptx_fixtures() -> bool:
    return has_local_pptx_fixtures()


skip_no_fixtures = pytest.mark.skipif(
    not has_local_docx_fixtures(),
    reason="vendored DOCX fixtures missing under tests/fixtures/docx/",
)

skip_no_pptx_fixtures = pytest.mark.skipif(
    not has_local_pptx_fixtures(),
    reason="vendored PPTX fixtures missing under tests/fixtures/pptx/",
)


def skip_without_external_fixture(*parts: str) -> pytest.MarkDecorator:
    """Skip a test unless the named external fixture resolves.

    Set ``KAOS_OFFICE_EXTERNAL_FIXTURES_DIR`` to a directory with the
    vendored layout (``<root>/docx/``, ``<root>/pptx/``, ...) to
    enable tests that depend on multi-MB documents.
    """
    return pytest.mark.skipif(
        external_fixture(*parts) is None,
        reason=(
            f"external fixture {'/'.join(parts)} unavailable; set ${_EXTERNAL_ROOT_ENV} to enable"
        ),
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
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
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


def make_minimal_pptx(
    *,
    slide_xmls: list[str] | None = None,
    core_xml: str | None = None,
    app_xml: str | None = None,
    notes_xmls: dict[int, str] | None = None,
    diagram_data_xmls: dict[str, str] | None = None,
    extra_parts: dict[str, str] | None = None,
    extra_slide_rels: dict[int, list[str]] | None = None,
) -> bytes:
    """Create a minimal valid PPTX file in memory.

    Args:
        slide_xmls: List of slide body XML (shape tree content).
        core_xml: Full docProps/core.xml, or None to omit.
        app_xml: Full docProps/app.xml, or None to omit.
        notes_xmls: Dict of slide_index (0-based) → notes XML.
        diagram_data_xmls: Dict of part_path → diagram data XML.
        extra_parts: Dict of part_path → content (for charts, etc.).
        extra_slide_rels: Dict of slide_index (0-based) → list of rel XML strings.

    Returns:
        Bytes of a valid PPTX ZIP file.
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

    if slide_xmls is None:
        slide_xmls = [
            f"""<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Title 1"/>
    <p:cNvSpPr/>
    <p:nvPr><p:ph type="title"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr><a:xfrm xmlns:a="{A_NS}"><a:off x="0" y="0"/><a:ext cx="9144000"
    cy="1143000"/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr xmlns:a="{A_NS}"/>
    <a:p xmlns:a="{A_NS}"><a:r><a:t>Test Title</a:t></a:r></a:p>
  </p:txBody>
</p:sp>"""
        ]

    # Build presentation.xml
    slide_refs = []
    pres_rels = []
    for i in range(len(slide_xmls)):
        rid = f"rId{i + 1}"
        slide_refs.append(f'<p:sldId id="{256 + i}" r:id="{rid}"/>')
        pres_rels.append(
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i + 1}.xml"/>'
        )

    # Add slide master rel
    master_rid = f"rId{len(slide_xmls) + 1}"
    pres_rels.append(
        f'<Relationship Id="{master_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        f'Target="slideMasters/slideMaster1.xml"/>'
    )

    pres_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:a="{A_NS}">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="{master_rid}"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {"".join(slide_refs)}
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>"""

    pres_rels_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  {"".join(pres_rels)}
</Relationships>"""

    # Minimal slide layout and master
    slide_layout_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"""

    slide_master_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"""

    slide_layout_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

    slide_master_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""

    root_rels = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="ppt/presentation.xml"/>
</Relationships>"""

    # Build content types
    overrides = [
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    ]
    for i in range(len(slide_xmls)):
        overrides.append(
            f'<Override PartName="/ppt/slides/slide{i + 1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    overrides.append(
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
    )
    overrides.append(
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
    )
    if notes_xmls:
        for i in notes_xmls:
            overrides.append(
                f'<Override PartName="/ppt/notesSlides/notesSlide{i + 1}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
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
        zf.writestr("ppt/presentation.xml", pres_xml)
        zf.writestr("ppt/_rels/presentation.xml.rels", pres_rels_xml)

        for i, slide_xml in enumerate(slide_xmls):
            slide_content = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      {slide_xml}
    </p:spTree>
  </p:cSld>
</p:sld>"""
            zf.writestr(f"ppt/slides/slide{i + 1}.xml", slide_content)

            # Build slide rels — always include slide layout
            slide_rel_entries = [
                '<Relationship Id="rId99" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slideLayout" '
                'Target="../slideLayouts/slideLayout1.xml"/>',
            ]
            if extra_slide_rels and i in extra_slide_rels:
                slide_rel_entries.extend(extra_slide_rels[i])
            if notes_xmls and i in notes_xmls:
                notes_rid = f"rId{100 + i}"
                slide_rel_entries.append(
                    f'<Relationship Id="{notes_rid}" '
                    "Type="
                    '"http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/notesSlide" '
                    f'Target="../notesSlides/notesSlide{i + 1}.xml"/>'
                )
                zf.writestr(f"ppt/notesSlides/notesSlide{i + 1}.xml", notes_xmls[i])

            slide_rels_xml = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{RELS_NS}">
  {"".join(slide_rel_entries)}
</Relationships>"""
            zf.writestr(f"ppt/slides/_rels/slide{i + 1}.xml.rels", slide_rels_xml)

        # Add slide layout and master
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels)

        if core_xml:
            zf.writestr("docProps/core.xml", core_xml)
        if app_xml:
            zf.writestr("docProps/app.xml", app_xml)

        if diagram_data_xmls:
            for part_path, xml_content in diagram_data_xmls.items():
                zf.writestr(part_path, xml_content)

        if extra_parts:
            for part_path, content in extra_parts.items():
                zf.writestr(part_path, content)

    return buf.getvalue()
