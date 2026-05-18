"""Round-trip test for ``numbering_label`` through the DOCX writer.

Read a fixture DOCX, write it back out, read the result, and verify
the rendered numbering label is preserved on the second pass.

The writer's "bake as plain text" approach (Stage 6 of the numbering
plan) means the round-tripped file no longer carries
``numbering.xml``-driven labels for the imported sections; instead
each block's rendered numeral lives in the run text. The reader picks
this up as the first token of the paragraph rather than via the
numbering resolver — so we verify the label survives at the visible
markdown layer, which is the user-facing contract.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from kaos_office.docx.reader import parse_docx
from kaos_office.docx.writer import write_docx_bytes

serialize_markdown = pytest.importorskip(
    "kaos_content", reason="kaos-content not installed"
).serialize_markdown


def _has_numbering_label_support() -> bool:
    try:
        from kaos_content import Paragraph

        return "numbering_label" in Paragraph.model_fields  # type: ignore[attr-defined]
    except ImportError:
        return False


requires_label = pytest.mark.skipif(
    not _has_numbering_label_support(),
    reason="kaos-content release with numbering_label not yet installed",
)


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "docx" / "numbering"


@requires_label
@pytest.mark.parametrize(
    "fixture",
    [
        "numbering_simple",
        "numbering_nda_governing_law",
        "numbering_legal_outline",
    ],
)
def test_round_trip_preserves_labels(fixture: str, tmp_path: Path) -> None:
    src_doc = parse_docx(FIXTURES / f"{fixture}.docx")
    src_markdown = serialize_markdown(src_doc).strip()

    out_bytes = write_docx_bytes(src_doc)
    out_path = tmp_path / f"{fixture}.roundtrip.docx"
    out_path.write_bytes(out_bytes)

    round_tripped_doc = parse_docx(out_path)
    round_tripped_markdown = serialize_markdown(round_tripped_doc).strip()

    # The "visible label" contract: every line of the original
    # rendered markdown must appear in the round-tripped rendering.
    # Order is the same; minor whitespace variation is tolerated.
    for line in src_markdown.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            # Heading-rendering may shift between # and explicit
            # heading style at the boundary; we still want the body
            # token to match.
            stripped = stripped.lstrip("#").strip()
        if not stripped:
            continue
        assert stripped in round_tripped_markdown, (
            f"line {stripped!r} from source markdown missing in round-trip:\n"
            f"--- source ---\n{src_markdown}\n--- round-trip ---\n{round_tripped_markdown}"
        )
