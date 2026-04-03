"""Tests for XLSX reader — parse_xlsx, list_sheets, real fixtures."""

from __future__ import annotations

import json
import time
from io import StringIO
from pathlib import Path
from unittest import mock

import pytest

from kaos_office.xlsx.reader import list_sheets, parse_xlsx

FIXTURES = Path(__file__).parent.parent / "fixtures" / "xlsx"
PAYMENT = FIXTURES / "payment-report-07-01-20-thru-07-15-20.xlsx"
CBS = FIXTURES / "CBS-BNSF-2015-Q4.xlsx"
PPP = FIXTURES / "ppplf-transaction-specific-disclosures-07-13-21.xlsx"


# ---------------------------------------------------------------------------
# parse_xlsx — basic
# ---------------------------------------------------------------------------


class TestParseXlsx:
    def test_payment_report_basic(self) -> None:
        doc = parse_xlsx(PAYMENT, header_row=2)
        assert len(doc.tables) == 1
        t = doc.tables[0]
        assert t.name == "July"
        assert t.row_count > 100

    def test_payment_report_columns(self) -> None:
        doc = parse_xlsx(PAYMENT, header_row=2)
        names = doc.tables[0].column_names()
        assert len(names) >= 10

    def test_cbs_multi_sheet(self) -> None:
        doc = parse_xlsx(CBS)
        assert len(doc.tables) == 2
        names = doc.table_names()
        assert "4Q15 CBS" in names
        assert "Page 2" in names

    def test_cbs_row_counts(self) -> None:
        doc = parse_xlsx(CBS)
        for t in doc.tables:
            assert t.row_count > 0

    def test_ppp_large_file(self) -> None:
        doc = parse_xlsx(PPP)
        assert len(doc.tables) >= 1
        t = doc.tables[0]
        assert t.row_count > 10_000

    def test_ppp_columns(self) -> None:
        doc = parse_xlsx(PPP)
        t = doc.tables[0]
        assert len(t.columns) >= 10

    def test_specific_sheet(self) -> None:
        doc = parse_xlsx(CBS, sheets=["Page 2"])
        assert len(doc.tables) == 1
        assert doc.tables[0].name == "Page 2"

    def test_max_rows(self) -> None:
        doc = parse_xlsx(PPP, max_rows=10)
        t = doc.tables[0]
        assert len(t.rows) == 10
        assert t.row_count > 10  # Knows the full count

    def test_provenance(self) -> None:
        doc = parse_xlsx(PAYMENT, header_row=2)
        assert doc.provenance is not None
        assert doc.provenance.extractor.startswith("kaos-office/xlsx/")
        assert doc.metadata.document_type == "xlsx"

    def test_missing_file(self) -> None:
        with pytest.raises(FileNotFoundError):
            parse_xlsx("/nonexistent/file.xlsx")


# ---------------------------------------------------------------------------
# list_sheets
# ---------------------------------------------------------------------------


class TestListSheets:
    def test_payment_single_sheet(self) -> None:
        sheets = list_sheets(PAYMENT)
        assert len(sheets) == 1
        assert sheets[0]["name"] == "July"
        assert sheets[0]["visible"] is True
        assert sheets[0]["rows"] > 100

    def test_cbs_multi_sheet(self) -> None:
        sheets = list_sheets(CBS)
        assert len(sheets) == 2
        names = [s["name"] for s in sheets]
        assert "4Q15 CBS" in names
        assert "Page 2" in names

    def test_ppp_dimensions(self) -> None:
        sheets = list_sheets(PPP)
        assert sheets[0]["rows"] > 10_000
        assert sheets[0]["columns"] >= 10

    def test_missing_file(self) -> None:
        with pytest.raises(FileNotFoundError):
            list_sheets("/nonexistent/file.xlsx")


# ---------------------------------------------------------------------------
# Type inference from real data
# ---------------------------------------------------------------------------


class TestTypeInference:
    def test_ppp_has_typed_columns(self) -> None:
        """PPP file should have numeric and text columns."""
        doc = parse_xlsx(PPP, max_rows=100)
        t = doc.tables[0]
        types = {c.name: c.column_type.value for c in t.columns}
        # Should have a mix of types
        type_values = set(types.values())
        assert len(type_values) >= 2  # At least text + something

    def test_data_values_are_python_native(self) -> None:
        """Calamine returns Python native types — verify no wrappers."""
        doc = parse_xlsx(CBS, max_rows=5)
        t = doc.tables[0]
        for row in t.rows:
            for val in row:
                if val is not None:
                    assert isinstance(val, (int, float, str, bool)), (
                        f"Unexpected type: {type(val)} for {val!r}"
                    )


# ---------------------------------------------------------------------------
# TabularDocument integration
# ---------------------------------------------------------------------------


class TestTabularDocumentIntegration:
    def test_serializers_work(self) -> None:
        """Verify kaos-content serializers work with XLSX-produced TabularDocument."""
        from kaos_content.serializers.tabular import (
            serialize_markdown_table,
            serialize_tabular_summary,
            serialize_tsv,
        )

        doc = parse_xlsx(CBS)
        t = doc.tables[0]

        tsv = serialize_tsv(t)
        assert len(tsv) > 0
        assert "\t" in tsv

        md = serialize_markdown_table(t, max_rows=5)
        assert "|" in md

        summary = serialize_tabular_summary(doc)
        assert "CBS" in summary or "4Q15" in summary

    def test_json_round_trip(self) -> None:
        """TabularDocument JSON round-trip preserves structure."""
        from kaos_content.artifacts import _tabular_from_json, _tabular_to_json

        doc = parse_xlsx(CBS)
        json_str = _tabular_to_json(doc)
        restored = _tabular_from_json(json_str)

        assert len(restored.tables) == len(doc.tables)
        for orig, rest in zip(doc.tables, restored.tables, strict=True):
            assert orig.name == rest.name
            assert orig.row_count == rest.row_count
            assert len(orig.columns) == len(rest.columns)

    def test_duckdb_registration(self) -> None:
        """XLSX → TabularDocument → DuckDB → SQL query."""
        duckdb = pytest.importorskip("duckdb")
        from kaos_content.bridges.duckdb import query_to_table, register_document

        doc = parse_xlsx(CBS)
        con = duckdb.connect()
        names = register_document(con, doc)
        assert len(names) == 2

        result = query_to_table(con, 'SELECT COUNT(*) FROM "4Q15 CBS"', name="count")
        assert result.rows[0][0] > 0


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


class TestXlsxCLI:
    def test_xlsx_extract_tsv(self) -> None:
        from kaos_office.cli import main

        with mock.patch("sys.stdout", new_callable=StringIO) as out:
            main(["xlsx-extract", str(CBS), "--format", "tsv"])
        output = out.getvalue()
        assert len(output) > 100
        assert "\t" in output

    def test_xlsx_sheets_json(self) -> None:
        from kaos_office.cli import main

        with mock.patch("sys.stdout", new_callable=StringIO) as out:
            main(["xlsx-sheets", str(CBS), "--json"])
        data = json.loads(out.getvalue())
        assert data["command"] == "xlsx-sheets"
        assert data["count"] == 2

    def test_xlsx_sheet_specific(self) -> None:
        from kaos_office.cli import main

        with mock.patch("sys.stdout", new_callable=StringIO) as out:
            main(["xlsx-sheet", str(CBS), "Page 2", "--format", "markdown"])
        output = out.getvalue()
        assert "|" in output

    def test_xlsx_extract_json_envelope(self) -> None:
        from kaos_office.cli import main

        with mock.patch("sys.stdout", new_callable=StringIO) as out:
            main(["xlsx-extract", str(CBS), "--json"])
        data = json.loads(out.getvalue())
        assert data["command"] == "xlsx-extract"
        assert data["table_count"] == 2


# ---------------------------------------------------------------------------
# Performance benchmark
# ---------------------------------------------------------------------------


class TestXlsxPerformance:
    def test_ppp_parse_under_3s(self) -> None:
        """937KB XLSX file (14K+ rows) should parse in < 3s with calamine."""
        start = time.monotonic()
        doc = parse_xlsx(PPP)
        elapsed = time.monotonic() - start
        assert doc.tables[0].row_count > 10_000
        assert elapsed < 3.0, f"PPP parse took {elapsed:.2f}s (limit: 3s)"

    def test_ppp_list_sheets_under_500ms(self) -> None:
        start = time.monotonic()
        list_sheets(PPP)
        elapsed = time.monotonic() - start
        assert elapsed < 0.5, f"List sheets took {elapsed:.2f}s (limit: 0.5s)"

    def test_cbs_parse_under_500ms(self) -> None:
        start = time.monotonic()
        parse_xlsx(CBS)
        elapsed = time.monotonic() - start
        assert elapsed < 0.5, f"CBS parse took {elapsed:.2f}s (limit: 0.5s)"
