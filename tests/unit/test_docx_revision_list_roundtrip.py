"""Regression tests: body-level revision blocks adjacent to lists.

A ``<w:ins>`` / ``<w:del>`` block that follows list paragraphs must close
the open list before opening its revision ``Div`` — otherwise the reader
splices the revision wrapper into the still-open ``OrderedList`` /
``BulletList`` as a stray ``Div`` child, which fails ``ListItem``
validation. This was surfaced by the redline engine, which routinely
emits block-level revision Divs next to lists. See
``reader._handle_body_revision``.
"""

from __future__ import annotations

from pathlib import Path

from kaos_content.model.blocks import ListItem, OrderedList, Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.revision import Revisions, make_block_deletion, make_block_insertion

from kaos_office.docx import parse_docx, write_docx


def _li(text: str) -> ListItem:
    return ListItem(children=(Paragraph(children=(Text(value=text),)),))


def _roundtrip(doc: ContentDocument, path: Path) -> ContentDocument:
    write_docx(doc, path)
    return parse_docx(path, track_changes=True)


def test_deleted_paragraph_after_list_round_trips(tmp_path: Path) -> None:
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(
            OrderedList(children=(_li("Keep one."), _li("Keep two."))),
            make_block_deletion(
                Paragraph(children=(Text(value="Loose deleted para."),)),
                author="X",
                revision_id="0",
            ),
        ),
    )
    reparsed = _roundtrip(doc, tmp_path / "del_after_list.docx")
    # The list must not absorb the revision Div: two top-level blocks survive.
    types = [b.node_type for b in reparsed.body]
    assert "ordered_list" in types
    assert "div" in types
    assert len(Revisions.from_document(reparsed)) == 1


def test_inserted_paragraph_between_two_lists_round_trips(tmp_path: Path) -> None:
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(
            OrderedList(children=(_li("A one."), _li("A two."))),
            make_block_insertion(
                Paragraph(children=(Text(value="Inserted between lists."),)),
                author="X",
                revision_id="0",
            ),
            OrderedList(children=(_li("B one."), _li("B two."))),
        ),
    )
    reparsed = _roundtrip(doc, tmp_path / "ins_between_lists.docx")
    assert [b.node_type for b in reparsed.body].count("ordered_list") == 2
    assert len(Revisions.from_document(reparsed)) == 1
