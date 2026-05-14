"""Top-level ``parse_*`` API naming, per docs/guides/python-api-naming.md.

kaos-office already used ``parse_docx`` / ``parse_pptx``; this guards the
addition of ``parse_xlsx`` at the top level so all three sibling readers
have the same call shape.
"""

from __future__ import annotations

from pathlib import Path

import pytest

_FIXTURES = Path(__file__).parent.parent / "fixtures"


def _pick(ext: str) -> Path:
    candidates = sorted((_FIXTURES / ext.lstrip(".")).glob(f"*{ext}"))
    if not candidates:
        pytest.skip(f"No {ext} fixture available")
    candidates.sort(key=lambda p: p.stat().st_size)
    return candidates[0]


def test_parse_docx_importable_from_top_level() -> None:
    from kaos_office import parse_docx

    assert callable(parse_docx)


def test_parse_pptx_importable_from_top_level() -> None:
    from kaos_office import parse_pptx

    assert callable(parse_pptx)


def test_parse_xlsx_importable_from_top_level() -> None:
    """PA3: parse_xlsx is exposed at the top level alongside parse_docx/parse_pptx."""
    from kaos_office import parse_xlsx

    assert callable(parse_xlsx)


def test_parse_docx_returns_content_document() -> None:
    from kaos_content.model.document import ContentDocument

    from kaos_office import parse_docx

    doc = parse_docx(_pick(".docx"))
    assert isinstance(doc, ContentDocument)


def test_parse_pptx_returns_content_document() -> None:
    from kaos_content.model.document import ContentDocument

    from kaos_office import parse_pptx

    doc = parse_pptx(_pick(".pptx"))
    assert isinstance(doc, ContentDocument)


def test_parse_xlsx_returns_tabular_document() -> None:
    from kaos_content.model.tabular import TabularDocument

    from kaos_office import parse_xlsx

    doc = parse_xlsx(_pick(".xlsx"))
    assert isinstance(doc, TabularDocument)
