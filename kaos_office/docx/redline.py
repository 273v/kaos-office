"""Generate a DOCX redline by comparing two Word documents.

This is the DOCX-facing convenience layer over
:func:`kaos_content.compare_documents`: it parses two ``.docx`` files into
``ContentDocument`` trees, compares them, and (for :func:`write_redline`)
serializes the result back to a ``.docx`` with native tracked changes.

The inputs are parsed with ``track_changes=False`` so each side is its
effective current text — any tracked changes already present in an input
are flattened (insertions accepted, deletions dropped) before the
comparison. The generated redline therefore reflects the difference
between the two documents' *content*, not a merge of pre-existing
revisions.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

from kaos_content import compare_documents

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx

if TYPE_CHECKING:
    from kaos_content import ContentDocument


def compare_docx(
    original: str | Path,
    revised: str | Path,
    *,
    author: str = "Reviewer",
    date: datetime | None = None,
    detect_moves: bool = True,
) -> ContentDocument:
    """Compare two DOCX files and return a redlined ``ContentDocument``.

    Args:
        original: Path to the baseline ``.docx``.
        revised: Path to the edited ``.docx``.
        author: Author name recorded on every generated revision.
        date: Timestamp recorded on every generated revision (``None`` =
            undated).
        detect_moves: When True, a deleted block that closely matches an
            inserted block elsewhere is tagged as a move pair.

    Returns:
        A ``ContentDocument`` whose ``body`` expresses the differences as
        tracked-change ``rev-*`` wrappers. Pass it to
        :func:`kaos_office.docx.writer.write_docx` to produce a Word file
        with native tracked changes, or use
        :func:`kaos_content.revision.view` to render original / final /
        markup.
    """
    original_doc = parse_docx(original)
    revised_doc = parse_docx(revised)
    return compare_documents(
        original_doc,
        revised_doc,
        author=author,
        date=date,
        detect_moves=detect_moves,
    )


def write_redline(
    original: str | Path,
    revised: str | Path,
    output: str | Path,
    *,
    author: str = "Reviewer",
    date: datetime | None = None,
    detect_moves: bool = True,
) -> Path:
    """Compare two DOCX files and write a tracked-changes redline ``.docx``.

    Returns the path the redline was written to. See :func:`compare_docx`
    for the comparison semantics and argument meanings.
    """
    redline = compare_docx(original, revised, author=author, date=date, detect_moves=detect_moves)
    return write_docx(redline, output)
