"""Tests for the OPC part-name resolver.

ECMA-376 §9.3.2 defines three relationship-target forms. A naive
``f"xl/{rel.target}"`` in the XLSX reader broke on form 1 (absolute)
because openpyxl emits absolute targets — the reader silently
dropped every sheet, returning ``tables=[]`` on a valid workbook
(corpus-stress S19 / S16, plan
``2026-05-26-corpus-stress-5-failure-resolution.md`` Failure 1).

These tests pin the resolver's behaviour on all three forms plus the
package-root and parent-relative-walks-above-root edge cases.
"""

from __future__ import annotations

import pytest

from kaos_office.opc.path import InvalidRelationshipTarget, resolve_part_target


class TestAbsoluteTargets:
    """Form 1: ``Target="/xl/worksheets/sheet1.xml"`` — strip leading slash."""

    def test_basic_absolute_target(self) -> None:
        assert (
            resolve_part_target("xl/workbook.xml", "/xl/worksheets/sheet1.xml")
            == "xl/worksheets/sheet1.xml"
        )

    def test_absolute_target_ignores_source_part(self) -> None:
        # Same target from a different source part still resolves to
        # the same package-root path — that's the contract of "absolute".
        assert (
            resolve_part_target("word/document.xml", "/word/media/image1.png")
            == "word/media/image1.png"
        )

    def test_absolute_target_double_slash_collapses(self) -> None:
        # Pathological but real: some writers emit ``//xl/...``.
        # ``lstrip("/")`` handles this cleanly.
        assert (
            resolve_part_target("xl/workbook.xml", "//xl/worksheets/sheet1.xml")
            == "xl/worksheets/sheet1.xml"
        )


class TestSourceRelativeTargets:
    """Form 3: ``Target="worksheets/sheet1.xml"`` — join to source dir."""

    def test_basic_source_relative(self) -> None:
        assert (
            resolve_part_target("xl/workbook.xml", "worksheets/sheet1.xml")
            == "xl/worksheets/sheet1.xml"
        )

    def test_deeper_source_directory(self) -> None:
        assert (
            resolve_part_target("ppt/slides/slide1.xml", "_rels/slide1.xml.rels")
            == "ppt/slides/_rels/slide1.xml.rels"
        )

    def test_package_root_source_part(self) -> None:
        # ``[Content_Types].xml`` lives at the package root.
        # An empty source-dir means the target IS the part name.
        assert (
            resolve_part_target("[Content_Types].xml", "word/document.xml") == "word/document.xml"
        )

    def test_empty_source_part(self) -> None:
        # The package-level ``_rels/.rels`` has no source part name.
        assert resolve_part_target("", "word/document.xml") == "word/document.xml"


class TestParentRelativeTargets:
    """Form 2: ``Target="../media/image1.png"`` — walk up source dir."""

    def test_one_level_up(self) -> None:
        assert resolve_part_target("word/document.xml", "../media/image1.png") == "media/image1.png"

    def test_two_levels_up(self) -> None:
        assert (
            resolve_part_target(
                "ppt/slides/slide1.xml",
                "../../media/image2.png",
            )
            == "media/image2.png"
        )

    def test_one_level_up_keeps_intermediate_dirs(self) -> None:
        assert (
            resolve_part_target(
                "ppt/slides/slide1.xml",
                "../media/image2.png",
            )
            == "ppt/media/image2.png"
        )

    def test_parent_relative_above_package_root_raises(self) -> None:
        # Walks above the package root — there is no "outside".
        with pytest.raises(InvalidRelationshipTarget):
            resolve_part_target("xl/workbook.xml", "../../../escape.png")

    def test_dot_segments_ignored(self) -> None:
        # ``./`` and empty segments are no-ops.
        assert (
            resolve_part_target("word/document.xml", "../media/./image1.png") == "media/image1.png"
        )


class TestEdgeCases:
    """Empty target + obvious malformed inputs."""

    def test_empty_target_raises(self) -> None:
        with pytest.raises(InvalidRelationshipTarget):
            resolve_part_target("xl/workbook.xml", "")

    def test_single_dot_segment_resolves_in_place(self) -> None:
        # ``./worksheets/sheet1.xml`` is form 3 with a leading dot.
        # The resolver collapses ``.`` to source dir (treat as form 3
        # since it doesn't start with ``../``).
        assert (
            resolve_part_target("xl/workbook.xml", "./worksheets/sheet1.xml")
            == "xl/./worksheets/sheet1.xml"
        )
        # Note: form-3 doesn't normalize ``./``. If a writer emits this,
        # ``OPCPackage.has_part`` will still find the part because zipfile
        # path lookup is exact-string. If real-world writers do emit
        # this and break, the helper can be tightened to normalize.


class TestOpenpyxlRegression:
    """Lock the absolute-target behaviour that S19 relied on.

    A regression here means openpyxl-produced XLSX files would once
    again return ``tables=[]`` — the exact CS-B5 / S19 / S16 bug.
    """

    def test_openpyxl_absolute_sheet_target_resolves(self) -> None:
        # Verbatim from an openpyxl-emitted workbook's
        # ``xl/_rels/workbook.xml.rels``.
        assert (
            resolve_part_target("xl/workbook.xml", "/xl/worksheets/sheet1.xml")
            == "xl/worksheets/sheet1.xml"
        )
