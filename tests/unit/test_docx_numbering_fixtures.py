"""Fixture-driven verification of attorney-citable numbering output.

Each fixture under ``tests/fixtures/docx/numbering/`` is a real
on-disk DOCX archive paired with a ``.expected.md`` ground-truth
sibling produced by hand. The test reads the DOCX, renders the AST to
markdown via kaos-content, and asserts a token-by-token match. A diff
indicates either:

* a regression in the numbering resolver (state.py / parser.py), or
* an intentional behavior change that requires updating the ground
  truth (regenerate via ``tests/fixtures/docx/numbering/generate.py``
  and review the diff before committing).

Per the verification gate in
``kaos-modules/docs/plans/docx-numbering-resolution.md``, the check is
that the rendered output reads exactly the way an attorney would cite
the source — not just that a serializer ran without error.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from kaos_office.docx.reader import parse_docx

# kaos-content ≥ 0.1.0a11 introduces the serialize_markdown signature
# we depend on. Tests are skipped (not failed) when the locked
# kaos-content predates the field — that keeps CI green during the
# release-coordination window. Once kaos-office bumps the floor,
# remove the skip block in a single follow-up commit.
serialize_markdown = pytest.importorskip(
    "kaos_content", reason="kaos-content not installed"
).serialize_markdown


FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "docx" / "numbering"


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


@requires_label
@pytest.mark.parametrize(
    "name",
    [
        "numbering_simple",
        "numbering_nda_governing_law",
        "numbering_legal_outline",
    ],
)
def test_docx_to_markdown_matches_ground_truth(name: str) -> None:
    docx_path = FIXTURES / f"{name}.docx"
    expected_path = FIXTURES / f"{name}.expected.md"
    assert docx_path.exists(), f"missing fixture {docx_path}"
    assert expected_path.exists(), f"missing ground truth {expected_path}"

    doc = parse_docx(docx_path)
    rendered = serialize_markdown(doc).strip()
    expected = expected_path.read_text(encoding="utf-8").strip()

    if rendered != expected:
        # Provide a diff that's actionable: regenerate via
        # tests/fixtures/docx/numbering/generate.py if the new shape
        # is intentional, otherwise fix the resolver / serializer.
        import difflib

        diff = "\n".join(
            difflib.unified_diff(
                expected.splitlines(),
                rendered.splitlines(),
                fromfile=str(expected_path),
                tofile=f"parse_docx({docx_path.name}) → serialize_markdown",
                lineterm="",
            )
        )
        pytest.fail(f"\n{diff}")
