"""Regression test for audit-04 F-002: top-level writer re-exports.

audit-04/kaos-office.md flagged that `kaos_office.__all__` exposed the
parser functions (`parse_docx`, `parse_pptx`, `parse_xlsx`) and 17 tool
classes / registration functions, but did NOT expose the writer
functions (`write_docx`, `write_pptx`, `write_xlsx`) or `list_sheets`.

That created an asymmetric top-level surface: the README and tool
registry both name authoring tools (`WriteDocxTool`, `WritePptxTool`,
`WriteXlsxTool`, `register_office_authoring_tools`) and the per-format
subpackages re-export the writer functions, but the top-level facade
didn't. Downstream callers had to know whether to import from
`kaos_office.docx` vs `kaos_office.pptx` vs `kaos_office.xlsx`, while
the parser side worked through the flat `kaos_office` import.

This test pins the symmetric surface so a future refactor cannot
quietly drop the re-export.
"""

from __future__ import annotations

import kaos_office


def test_writer_functions_are_reexported_at_top_level() -> None:
    """`kaos_office.write_*` must be the same callable as the subpackage's.

    The fix re-exports through each format subpackage (which preserves
    the lazy `python-pptx` ImportError wrapper for `write_pptx`), so the
    top-level binding must be identity-equal to the subpackage binding.
    """
    from kaos_office.docx import write_docx
    from kaos_office.pptx import write_pptx
    from kaos_office.xlsx import list_sheets, write_xlsx

    assert kaos_office.write_docx is write_docx
    assert kaos_office.write_pptx is write_pptx
    assert kaos_office.write_xlsx is write_xlsx
    assert kaos_office.list_sheets is list_sheets


def test_writer_functions_in_dunder_all() -> None:
    """`__all__` is the wildcard-import contract; pin the new entries."""
    for name in ("write_docx", "write_pptx", "write_xlsx", "list_sheets"):
        assert name in kaos_office.__all__, (
            f"audit-04 F-002 regression: {name!r} missing from kaos_office.__all__"
        )


def test_parser_and_writer_surfaces_are_symmetric() -> None:
    """A reader for each format must have a corresponding writer.

    This is the deeper finding behind F-002: callers who learn the
    parse_X name expect a write_X name on the same facade.
    """
    for fmt in ("docx", "pptx", "xlsx"):
        parse_name = f"parse_{fmt}"
        write_name = f"write_{fmt}"
        assert parse_name in kaos_office.__all__, (
            f"{parse_name!r} unexpectedly missing — fixture assumption broken"
        )
        assert write_name in kaos_office.__all__, (
            f"audit-04 F-002 regression: parser {parse_name!r} is re-exported "
            f"but writer {write_name!r} is not — asymmetric facade"
        )
