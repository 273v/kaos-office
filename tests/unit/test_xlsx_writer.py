"""Unit tests for XLSX writer.

Round-trip tests: TabularDocument → write_xlsx → parse_xlsx → verify.
Modification round-trips: build → modify → write → re-parse → verify.
Column type mapping: verify each ColumnType produces correct Excel formatting.
Performance: large tables under time budgets.

Follows the same testing patterns as test_docx_writer.py.
"""

from __future__ import annotations

import datetime
import time
import zipfile
from io import BytesIO
from pathlib import Path

import pytest
from kaos_content.model.tabular import Column, ColumnType, Table, TabularDocument

from kaos_office.xlsx.reader import parse_xlsx
from kaos_office.xlsx.writer import write_xlsx, write_xlsx_bytes

FIXTURES = Path(__file__).resolve().parent.parent / "fixtures" / "xlsx"


def _zip_parts(xlsx_bytes: bytes) -> list[str]:
    """Extract sorted part names from XLSX bytes."""
    return sorted(zipfile.ZipFile(BytesIO(xlsx_bytes)).namelist())


def _simple_doc(
    name: str = "Sheet1",
    columns: tuple[Column, ...] | None = None,
    rows: tuple | None = None,
) -> TabularDocument:
    """Helper to build a simple TabularDocument."""
    if columns is None:
        columns = (
            Column(name="Name", column_type=ColumnType.TEXT),
            Column(name="Value", column_type=ColumnType.INTEGER),
        )
    if rows is None:
        rows = (("Alice", 100), ("Bob", 200))
    return TabularDocument(
        tables=(Table(name=name, columns=columns, rows=rows),),
    )


# ---------------------------------------------------------------------------
# OPC structure tests
# ---------------------------------------------------------------------------


class TestOPCStructure:
    """Verify output is a valid OPC package with correct parts."""

    def test_required_parts(self) -> None:
        data = write_xlsx_bytes(_simple_doc())
        parts = _zip_parts(data)
        assert "[Content_Types].xml" in parts
        assert "_rels/.rels" in parts
        assert "xl/workbook.xml" in parts
        assert "xl/worksheets/sheet1.xml" in parts
        assert "xl/sharedStrings.xml" in parts
        assert "xl/styles.xml" in parts

    def test_multi_sheet(self) -> None:
        doc = TabularDocument(
            tables=(
                Table(
                    name="First",
                    columns=(Column(name="A", column_type=ColumnType.TEXT),),
                    rows=(("a",),),
                ),
                Table(
                    name="Second",
                    columns=(Column(name="B", column_type=ColumnType.INTEGER),),
                    rows=((1,),),
                ),
            ),
        )
        parts = _zip_parts(write_xlsx_bytes(doc))
        assert "xl/worksheets/sheet1.xml" in parts
        assert "xl/worksheets/sheet2.xml" in parts

    def test_empty_table(self) -> None:
        doc = TabularDocument(
            tables=(
                Table(
                    name="Empty",
                    columns=(Column(name="A", column_type=ColumnType.TEXT),),
                    rows=(),
                ),
            ),
        )
        data = write_xlsx_bytes(doc)
        assert len(data) > 0
        parts = _zip_parts(data)
        assert "xl/worksheets/sheet1.xml" in parts


# ---------------------------------------------------------------------------
# Column type mapping tests
# ---------------------------------------------------------------------------


class TestColumnTypeMapping:
    """Verify each ColumnType produces correct values after round-trip."""

    def test_text_column(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="T", column_type=ColumnType.TEXT),),
            rows=(("hello",), ("world",)),
        )
        out = tmp_path / "text.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0] == ("hello",)
        assert doc2.tables[0].rows[1] == ("world",)

    def test_integer_column(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="N", column_type=ColumnType.INTEGER),),
            rows=((42,), (0,), (-7,)),
        )
        out = tmp_path / "int.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0] == (42,)
        assert doc2.tables[0].rows[1] == (0,)
        assert doc2.tables[0].rows[2] == (-7,)

    def test_float_column(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="F", column_type=ColumnType.FLOAT),),
            rows=((3.14,), (0.0,)),
        )
        out = tmp_path / "float.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert abs(doc2.tables[0].rows[0][0] - 3.14) < 0.001
        assert doc2.tables[0].rows[1][0] == 0.0

    def test_boolean_column(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="B", column_type=ColumnType.BOOLEAN),),
            rows=((True,), (False,)),
        )
        out = tmp_path / "bool.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0][0] is True
        assert doc2.tables[0].rows[1][0] is False

    def test_date_column(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="D", column_type=ColumnType.DATE),),
            rows=(("2024-01-15",), ("2024-12-31",)),
        )
        out = tmp_path / "date.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0][0] == datetime.date(2024, 1, 15)
        assert doc2.tables[0].rows[1][0] == datetime.date(2024, 12, 31)

    def test_none_values(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(
                Column(name="T", column_type=ColumnType.TEXT),
                Column(name="N", column_type=ColumnType.INTEGER),
            ),
            rows=(("a", None), (None, 5)),
        )
        out = tmp_path / "none.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        # None values should round-trip as None
        assert doc2.tables[0].rows[0][1] is None
        assert doc2.tables[0].rows[1][0] is None

    def test_mixed_types_document(self, tmp_path: Path) -> None:
        """A table with multiple column types in one sheet."""
        doc = TabularDocument(
            tables=(
                Table(
                    name="Mixed",
                    columns=(
                        Column(name="Name", column_type=ColumnType.TEXT),
                        Column(name="Age", column_type=ColumnType.INTEGER),
                        Column(name="Score", column_type=ColumnType.FLOAT),
                        Column(name="Active", column_type=ColumnType.BOOLEAN),
                        Column(name="Joined", column_type=ColumnType.DATE),
                    ),
                    rows=(
                        ("Alice", 30, 95.5, True, "2024-01-15"),
                        ("Bob", 25, 87.3, False, "2024-06-01"),
                    ),
                ),
            ),
        )
        out = tmp_path / "mixed.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].columns[0].name == "Name"
        assert doc2.tables[0].rows[0][0] == "Alice"
        assert doc2.tables[0].rows[0][1] == 30
        assert abs(doc2.tables[0].rows[0][2] - 95.5) < 0.01
        assert doc2.tables[0].rows[0][3] is True
        assert doc2.tables[0].rows[0][4] == datetime.date(2024, 1, 15)


# ---------------------------------------------------------------------------
# Round-trip tests (fixture-based)
# ---------------------------------------------------------------------------


class TestRoundTrip:
    """Parse real XLSX → write → re-parse → verify content."""

    def _roundtrip(
        self, fixture_name: str, tmp_path: Path
    ) -> tuple[TabularDocument, TabularDocument]:
        """Parse fixture, write, re-parse, return both."""
        src = parse_xlsx(FIXTURES / fixture_name)
        data = write_xlsx_bytes(src)
        out = tmp_path / "output.xlsx"
        out.write_bytes(data)
        dst = parse_xlsx(out)
        return src, dst

    def test_payment_report(self, tmp_path: Path) -> None:
        src, dst = self._roundtrip("payment-report-07-01-20-thru-07-15-20.xlsx", tmp_path)
        assert len(dst.tables) == len(src.tables)
        assert len(dst.tables[0].columns) == len(src.tables[0].columns)
        # All rows (including empty separator rows) now survive round-trip.
        assert len(dst.tables[0].rows) == len(src.tables[0].rows)

    def test_cbs_multi_sheet(self, tmp_path: Path) -> None:
        src, dst = self._roundtrip("CBS-BNSF-2015-Q4.xlsx", tmp_path)
        assert len(dst.tables) == len(src.tables)
        # CBS has all-None separator rows; all now survive round-trip.
        for i in range(len(src.tables)):
            assert len(dst.tables[i].columns) == len(src.tables[i].columns), (
                f"Sheet {i}: column count mismatch"
            )
            assert len(dst.tables[i].rows) == len(src.tables[i].rows), (
                f"Sheet {i}: row count mismatch "
                f"(src={len(src.tables[i].rows)} dst={len(dst.tables[i].rows)})"
            )

    def test_ppp_large(self, tmp_path: Path) -> None:
        src, dst = self._roundtrip("ppplf-transaction-specific-disclosures-07-13-21.xlsx", tmp_path)
        assert len(dst.tables) >= 1
        assert len(dst.tables[0].rows) == len(src.tables[0].rows)

    def test_states(self, tmp_path: Path) -> None:
        if not (FIXTURES / "states.xlsx").exists():
            pytest.skip("states.xlsx not available")
        src, dst = self._roundtrip("states.xlsx", tmp_path)
        assert len(dst.tables[0].rows) == len(src.tables[0].rows)


# ---------------------------------------------------------------------------
# Modification round-trip tests
# ---------------------------------------------------------------------------


class TestModificationRoundTrip:
    """Build → modify → write → re-parse → verify edit.

    TabularDocument uses frozen dataclasses, so we construct new instances
    directly rather than using model_copy.
    """

    def test_add_row(self, tmp_path: Path) -> None:
        """Add a row to a synthetic document and verify it persists."""
        src = _simple_doc(rows=(("Alice", 100), ("Bob", 200)))
        table = src.tables[0]
        new_table = Table(
            name=table.name,
            columns=table.columns,
            rows=(*table.rows, ("Charlie", 300)),
        )
        modified = TabularDocument(tables=(new_table,))

        out = tmp_path / "added_row.xlsx"
        write_xlsx(modified, out)
        doc2 = parse_xlsx(out)
        assert len(doc2.tables[0].rows) == 3
        assert doc2.tables[0].rows[-1][0] == "Charlie"
        assert doc2.tables[0].rows[-1][1] == 300

    def test_change_cell_value(self, tmp_path: Path) -> None:
        """Change a cell value and verify it persists."""
        src = _simple_doc(rows=(("Alice", 100), ("Bob", 200)))
        table = src.tables[0]
        rows_list = list(table.rows)
        rows_list[0] = ("Alice", 999)
        new_table = Table(
            name=table.name,
            columns=table.columns,
            rows=tuple(rows_list),
        )
        modified = TabularDocument(tables=(new_table,))

        out = tmp_path / "changed.xlsx"
        write_xlsx(modified, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0][1] == 999
        assert doc2.tables[0].rows[1][1] == 200  # Bob unchanged

    def test_modify_fixture_preserves_content(self, tmp_path: Path) -> None:
        """Modify a fixture document and verify non-modified content survives."""
        if not (FIXTURES / "states.xlsx").exists():
            pytest.skip("states.xlsx not available")
        src = parse_xlsx(FIXTURES / "states.xlsx")
        table = src.tables[0]
        rows_list = list(table.rows)
        n_cols = len(table.columns)
        rows_list[0] = tuple("MODIFIED" if i == 0 else rows_list[0][i] for i in range(n_cols))
        new_table = Table(name=table.name, columns=table.columns, rows=tuple(rows_list))
        modified = TabularDocument(tables=(new_table, *src.tables[1:]))

        out = tmp_path / "mod_fixture.xlsx"
        write_xlsx(modified, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0][0] == "MODIFIED"
        assert len(doc2.tables[0].rows) == len(table.rows)

    def test_add_sheet(self, tmp_path: Path) -> None:
        """Add a second sheet and verify both survive."""
        doc = _simple_doc()
        extra = Table(
            name="Extra",
            columns=(Column(name="X", column_type=ColumnType.TEXT),),
            rows=(("extra1",), ("extra2",)),
        )
        modified = TabularDocument(tables=(*doc.tables, extra))

        out = tmp_path / "multi.xlsx"
        write_xlsx(modified, out)
        doc2 = parse_xlsx(out)
        assert len(doc2.tables) == 2
        assert doc2.tables[1].rows[0] == ("extra1",)


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------


class TestErrorHandling:
    def test_empty_table_produces_valid_xlsx(self) -> None:
        doc = TabularDocument(
            tables=(
                Table(
                    name="Empty",
                    columns=(Column(name="A", column_type=ColumnType.TEXT),),
                    rows=(),
                ),
            ),
        )
        data = write_xlsx_bytes(doc)
        assert len(data) > 0
        zf = zipfile.ZipFile(BytesIO(data))
        assert "xl/worksheets/sheet1.xml" in zf.namelist()

    def test_write_creates_parent_dirs(self, tmp_path: Path) -> None:
        doc = _simple_doc()
        out = tmp_path / "a" / "b" / "output.xlsx"
        result = write_xlsx(doc, out)
        assert result == out
        assert out.exists()

    def test_unicode_content(self, tmp_path: Path) -> None:
        doc = _simple_doc(
            columns=(Column(name="Text", column_type=ColumnType.TEXT),),
            rows=(("日本語テスト",), ("émojis 🎉",), ("café résumé",)),
        )
        out = tmp_path / "unicode.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert doc2.tables[0].rows[0][0] == "日本語テスト"
        assert doc2.tables[0].rows[2][0] == "café résumé"

    def test_long_sheet_name_truncated(self, tmp_path: Path) -> None:
        """Sheet names > 31 chars should be handled."""
        doc = TabularDocument(
            tables=(
                Table(
                    name="A" * 50,
                    columns=(Column(name="X", column_type=ColumnType.TEXT),),
                    rows=(("v",),),
                ),
            ),
        )
        out = tmp_path / "longname.xlsx"
        write_xlsx(doc, out)
        doc2 = parse_xlsx(out)
        assert len(doc2.tables) == 1
        assert len(doc2.tables[0].name) <= 31


# ---------------------------------------------------------------------------
# Performance
# ---------------------------------------------------------------------------


class TestPerformance:
    def test_1000_rows_under_3s(self) -> None:
        """1000 rows should write in under 3 seconds."""
        rows = tuple((f"item_{i}", i, i * 1.5) for i in range(1000))
        doc = TabularDocument(
            tables=(
                Table(
                    name="Big",
                    columns=(
                        Column(name="Name", column_type=ColumnType.TEXT),
                        Column(name="ID", column_type=ColumnType.INTEGER),
                        Column(name="Score", column_type=ColumnType.FLOAT),
                    ),
                    rows=rows,
                ),
            ),
        )
        start = time.monotonic()
        data = write_xlsx_bytes(doc)
        elapsed = time.monotonic() - start
        assert elapsed < 3.0, f"write_xlsx_bytes took {elapsed:.2f}s (budget 3s)"
        assert len(data) > 0

    def test_fixture_roundtrip_under_5s(self, tmp_path: Path) -> None:
        """Parse → write → re-parse of largest fixture under 5s."""
        fixture = FIXTURES / "ppplf-transaction-specific-disclosures-07-13-21.xlsx"
        if not fixture.exists():
            pytest.skip("PPP fixture not available")
        start = time.monotonic()
        doc = parse_xlsx(fixture)
        data = write_xlsx_bytes(doc)
        out = tmp_path / "perf.xlsx"
        out.write_bytes(data)
        parse_xlsx(out)
        elapsed = time.monotonic() - start
        assert elapsed < 5.0, f"Full round-trip took {elapsed:.2f}s (budget 5s)"
