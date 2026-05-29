"""Comprehensive OOXML symmetry for every tracked-change type.

Authoring one of each ``rev-*`` shape (inline insert / delete /
move-from / move-to, and a block move pair), writing to DOCX, and
re-parsing must preserve every type, the shared move name, and emit
``w:delText`` for deletions and move-froms (OOXML §17.16.2).
"""

from __future__ import annotations

import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from kaos_content.model.blocks import Paragraph
from kaos_content.model.document import ContentDocument, DocumentMetadata
from kaos_content.model.inlines import Text
from kaos_content.revision import (
    Revisions,
    RevisionType,
    make_block_move_from,
    make_block_move_to,
    make_inline_deletion,
    make_inline_insertion,
    make_inline_move_from,
    make_inline_move_to,
)

from kaos_office.docx import parse_docx, write_docx


def test_all_inline_revision_types_round_trip() -> None:
    tmp = Path(tempfile.mkdtemp())
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(
            Paragraph(
                children=(
                    Text(value="keep "),
                    make_inline_insertion(Text(value="ins "), author="A", revision_id="0"),
                    make_inline_deletion(Text(value="del "), author="A", revision_id="1"),
                    make_inline_move_from(
                        Text(value="mfrom "), author="A", move_name="m1", revision_id="2"
                    ),
                    make_inline_move_to(
                        Text(value="mto"), author="A", move_name="m1", revision_id="3"
                    ),
                )
            ),
        ),
    )
    out = tmp / "all_inline.docx"
    write_docx(doc, out)

    xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8", "replace")
    assert xml.count("<w:ins ") == 1
    assert xml.count("<w:del ") == 1
    assert xml.count("<w:moveFrom ") == 1
    assert xml.count("<w:moveTo ") == 1
    # delText for the deletion AND the move-from (both remove text).
    assert xml.count("<w:delText") == 2

    revs = Revisions.from_document(parse_docx(out, track_changes=True))
    assert dict(Counter(r.change_type.value for r in revs)) == {
        "insertion": 1,
        "deletion": 1,
        "move_from": 1,
        "move_to": 1,
    }
    assert {r.move_name for r in revs if r.move_name} == {"m1"}


def test_block_move_pair_round_trips() -> None:
    tmp = Path(tempfile.mkdtemp())
    doc = ContentDocument(
        metadata=DocumentMetadata(title=""),
        body=(
            make_block_move_from(
                Paragraph(children=(Text(value="moved clause"),)),
                author="A",
                move_name="bm",
                revision_id="0",
            ),
            Paragraph(children=(Text(value="stable"),)),
            make_block_move_to(
                Paragraph(children=(Text(value="moved clause"),)),
                author="A",
                move_name="bm",
                revision_id="1",
            ),
        ),
    )
    out = tmp / "block_move.docx"
    write_docx(doc, out)
    revs = Revisions.from_document(parse_docx(out, track_changes=True))
    kinds = {r.change_type for r in revs}
    assert RevisionType.MOVE_FROM in kinds
    assert RevisionType.MOVE_TO in kinds
    assert {r.move_name for r in revs if r.move_name} == {"bm"}
