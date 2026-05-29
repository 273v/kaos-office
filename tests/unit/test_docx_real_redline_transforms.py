"""Integration: accept/reject transforms on a real parsed DOCX redline.

The kaos-content transform tests exercise accept/reject on synthetic
documents. This test runs the same transforms on revisions the DOCX
reader produced from a real tracked-changes file (Toro Term Loan redline,
generated in LibreOffice), closing the loop the compare engine relies on:
reader → typed revisions → resolve.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_content.model.document import ContentDocument
from kaos_content.revision import (
    Revisions,
    RevisionType,
    accept_all,
    reject_all,
)
from kaos_content.traversal.visitor import extract_text

from kaos_office.docx import compare_docx, parse_docx

_FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "docx"
_REDLINE = _FIXTURES / "Toro 2022 Term Loan - Redline v1.docx"
_CLEAN = _FIXTURES / "Toro 2022 Term Loan.docx"


def _text_len(doc: ContentDocument) -> int:
    return len("".join(extract_text(b) for b in doc.body))


@pytest.mark.skipif(not _REDLINE.exists(), reason="Toro redline fixture missing")
class TestRealRedlineTransforms:
    def test_reader_surfaces_insertions_and_deletions(self) -> None:
        doc = parse_docx(_REDLINE, track_changes=True)
        revs = Revisions.from_document(doc)
        assert len(revs) >= 1
        types = {r.change_type for r in revs}
        assert RevisionType.INSERTION in types
        assert RevisionType.DELETION in types

    def test_accept_and_reject_resolve_all_revisions(self) -> None:
        doc = parse_docx(_REDLINE, track_changes=True)
        before = len(Revisions.from_document(doc))
        assert before >= 1

        accepted = accept_all(doc)
        rejected = reject_all(doc)
        # Both views fully resolve the markup.
        assert len(Revisions.from_document(accepted)) == 0
        assert len(Revisions.from_document(rejected)) == 0

    def test_accept_drops_deletions_reject_restores_them(self) -> None:
        doc = parse_docx(_REDLINE, track_changes=True)
        # Accepting drops deleted text; rejecting restores it. On a redline
        # with deletions the rejected (original) text is longer than the
        # accepted (final) text.
        assert _text_len(reject_all(doc)) > _text_len(accept_all(doc))


def _norm(doc: ContentDocument) -> str:
    return " ".join(" ".join(extract_text(b).split()) for b in doc.body)


@pytest.mark.skipif(not (_CLEAN.exists() and _REDLINE.exists()), reason="Toro fixtures missing")
def test_compare_clean_against_already_redlined_input() -> None:
    """Real workflow: redline a clean doc against a marked-up version.

    ``compare_docx`` parses inputs with ``track_changes=False``, so the
    already-redlined side contributes its *accepted* (effective) text. The
    generated redline must then accept to that effective text and reject to
    the clean original.
    """
    redline = compare_docx(_CLEAN, _REDLINE, author="Reviewer")
    assert Revisions.from_document(redline)
    # parse_docx(track_changes=False) is exactly the effective text of each side.
    assert _norm(accept_all(redline)) == _norm(parse_docx(_REDLINE))
    assert _norm(reject_all(redline)) == _norm(parse_docx(_CLEAN))
