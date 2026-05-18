"""Generate the on-disk DOCX fixtures used by the numbering tests.

Run once to materialize the ``.docx`` files in this directory; the
fixtures are committed to the repo so CI doesn't need to regenerate
them. Regenerate when you intentionally change a fixture's shape (and
keep its ``.expected.md`` sibling in lockstep).

The fixtures cover the cases called out in
``kaos-modules/docs/plans/docx-numbering-resolution.md`` Stage 5:

* ``numbering_simple.docx`` — three-level decimal / lowerLetter /
  lowerRoman pattern (textbook attorney citation).
* ``numbering_nda_governing_law.docx`` — NDA-style numbered headings
  ("Section 11. GOVERNING LAW") with sub-clause (a)/(b)/(c).
* ``numbering_legal_outline.docx`` — 1./1.1/1.1.1 with ``<w:isLgl/>``
  forcing decimal at every level.

Run:
    uv run python tests/fixtures/docx/numbering/generate.py

Authorship: AI-generated test fixtures. No client / private content;
no PII; redistributable. See AGENTS.md for the AI-authorship
discipline.
"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

# Allow standalone invocation (``python tests/fixtures/...``).
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

from kaos_office.ooxml.namespace import W

HERE = Path(__file__).resolve().parent

_CT = """\
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""  # noqa: E501

_PKG_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""  # noqa: E501

_DOC_RELS = """\
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""  # noqa: E501

_STYLES = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:outlineLvl w:val="1"/></w:pPr>
  </w:style>
</w:styles>
"""


def _build_docx(document_xml: str, numbering_xml: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CT)
        zf.writestr("_rels/.rels", _PKG_RELS)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/numbering.xml", numbering_xml)
        zf.writestr("word/styles.xml", _STYLES)
    return buf.getvalue()


# ── Fixture 1: simple three-level pattern ────────────────────────────


SIMPLE_NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="(%2)"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="(%3)"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""

SIMPLE_DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>First top item.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Second top item.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Nested clause.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Sub-sub clause.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

SIMPLE_EXPECTED_MD = """\
1. First top item.
2. Second top item.
   (a) Nested clause.
       (i) Sub-sub clause.
"""


# ── Fixture 2: NDA "Section 11. GOVERNING LAW" pattern ───────────────

NDA_NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="Section %1."/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="(%1)"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
"""

NDA_DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>CONFIDENTIAL INFORMATION</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>The Receiving Party agrees not to disclose...</w:t></w:r></w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>TERM</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>This Agreement is effective for two years.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>The Term may be extended in writing.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>GOVERNING LAW</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>The validity, interpretation, and performance of this Agreement shall be governed by the laws of the State of Delaware.</w:t></w:r></w:p>
  </w:body>
</w:document>
"""  # noqa: E501

NDA_EXPECTED_MD = """\
# Section 1. CONFIDENTIAL INFORMATION

The Receiving Party agrees not to disclose...

# Section 2. TERM

(a) This Agreement is effective for two years.
(b) The Term may be extended in writing.

# Section 3. GOVERNING LAW

The validity, interpretation, and performance of this Agreement shall be governed by the laws of the State of Delaware.
"""  # noqa: E501


# ── Fixture 3: legal outline with isLgl ─────────────────────────────

LEGAL_NUMBERING = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%1.%2"/>
      <w:isLgl/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%1.%2.%3"/>
      <w:isLgl/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""

LEGAL_DOCUMENT = f"""\
<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{W}">
  <w:body>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Definitions.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Affiliate.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Confidential Information.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Customer data.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Source code.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

# isLgl forces decimal at every referenced level even though levels 1
# and 2 have lowerLetter / lowerRoman numFmts. So we see ``1.1`` rather
# than ``1.a`` and ``1.1.1`` rather than ``1.a.i``.
LEGAL_EXPECTED_MD = """\
1 Definitions.
  1.1 Affiliate.
  1.2 Confidential Information.
      1.2.1 Customer data.
      1.2.2 Source code.
"""


def main() -> None:
    fixtures = {
        "numbering_simple.docx": _build_docx(SIMPLE_DOCUMENT, SIMPLE_NUMBERING),
        "numbering_nda_governing_law.docx": _build_docx(NDA_DOCUMENT, NDA_NUMBERING),
        "numbering_legal_outline.docx": _build_docx(LEGAL_DOCUMENT, LEGAL_NUMBERING),
    }
    expected = {
        "numbering_simple.expected.md": SIMPLE_EXPECTED_MD,
        "numbering_nda_governing_law.expected.md": NDA_EXPECTED_MD,
        "numbering_legal_outline.expected.md": LEGAL_EXPECTED_MD,
    }
    for name, content in fixtures.items():
        path = HERE / name
        path.write_bytes(content)
        print(f"wrote {path}  ({len(content)} bytes)")
    for name, text in expected.items():
        path = HERE / name
        path.write_text(text, encoding="utf-8")
        print(f"wrote {path}  ({len(text)} chars)")


if __name__ == "__main__":
    main()
