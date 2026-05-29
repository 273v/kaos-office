"""Headings must preserve inline structure (revisions + formatting).

Regression: the DOCX reader built headings from a flattened plain-text
string, discarding inline content. That silently dropped tracked changes
inside an edited heading (the redline showed no change and accept/reject
could not reproduce either side) and also lost run formatting (bold,
links) in headings. The reader now builds headings from their collected
inlines.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from kaos_content.model.blocks import Heading
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Strong, Text
from kaos_content.revision import Revisions, accept_all, reject_all
from kaos_content.traversal import walk
from kaos_content.traversal.visitor import extract_text

from kaos_office.docx import parse_docx, write_docx, write_redline


def _norm(doc: ContentDocument) -> str:
    return " ".join(" ".join(extract_text(b).split()) for b in doc.body)


def _heading_doc(path: Path, text: str) -> Path:
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(Heading(depth=1, children=(Text(value=text),)),),
    )
    write_docx(doc, path)
    return path


def test_heading_text_edit_redline_round_trips() -> None:
    tmp = Path(tempfile.mkdtemp())
    a = _heading_doc(tmp / "a.docx", "Governing Law and Venue")
    b = _heading_doc(tmp / "b.docx", "Governing Law and Jurisdiction")
    out = write_redline(a, b, tmp / "r.docx")
    reparsed = parse_docx(out, track_changes=True)

    assert len(Revisions.from_document(reparsed)) >= 1
    assert reparsed.body[0].node_type == "heading"
    assert _norm(accept_all(reparsed)) == _norm(parse_docx(b))
    assert _norm(reject_all(reparsed)) == _norm(parse_docx(a))


def test_plain_heading_still_single_text_child() -> None:
    tmp = Path(tempfile.mkdtemp())
    a = _heading_doc(tmp / "p.docx", "Simple Heading")
    heading = parse_docx(a).body[0]
    assert heading.node_type == "heading"
    assert extract_text(heading) == "Simple Heading"


def test_heading_inline_formatting_preserved_on_round_trip() -> None:
    tmp = Path(tempfile.mkdtemp())
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(
            Heading(
                depth=1,
                children=(
                    Text(value="The "),
                    Strong(children=(Text(value="Material"),)),
                    Text(value=" Terms"),
                ),
            ),
        ),
    )
    out = tmp / "fmt.docx"
    write_docx(doc, out)
    reparsed = parse_docx(out)
    assert any(type(n).__name__ == "Strong" for block in reparsed.body for n in walk(block))
    assert extract_text(reparsed.body[0]) == "The Material Terms"
