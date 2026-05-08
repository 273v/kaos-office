"""Integration tests for PPTX extraction against real presentation files.

Uses kelvin_office PPTX fixtures and local stress test fixtures.
"""

from __future__ import annotations

import pytest

from tests.conftest import (
    KELVIN_PPTX_FIXTURES,
    PPTX_STRESS_FIXTURES,
    external_fixture,
    skip_no_pptx_fixtures,
    skip_without_external_fixture,
)


@skip_no_pptx_fixtures
class TestHelloWorld:
    """Test Hello-World.pptx — simple presentation with titles and groups."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(KELVIN_PPTX_FIXTURES / "Hello-World.pptx")

    def test_slide_count(self):
        assert len(self.doc.body) == 9

    def test_title_extracted(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        assert "Hello World" in md

    def test_subtitle_extracted(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        assert "This is a subtitle" in md

    def test_metadata_populated(self):
        assert self.doc.metadata.title is not None
        assert "slide_count" in self.doc.metadata.extra
        assert self.doc.metadata.extra["slide_count"] == "9"


@skip_no_pptx_fixtures
class TestChartLibrary:
    """Test IEO2021_ChartLibrary_Industrial.pptx — 11 slides with 18 charts."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(KELVIN_PPTX_FIXTURES / "IEO2021_ChartLibrary_Industrial.pptx")

    def test_slide_count(self):
        assert len(self.doc.body) == 11

    def test_charts_as_tables(self):
        """Charts should be linearized as table blocks."""

        def count_tables(blocks):
            count = 0
            for b in blocks:
                if b.node_type == "table":
                    count += 1
                if hasattr(b, "children"):
                    count += count_tables(b.children)
            return count

        assert count_tables(self.doc.body) == 18

    def test_chart_data_extracted(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        # Chart categories should appear
        assert "OECD" in md or "non-OECD" in md

    def test_markdown_length(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        # 18 charts with data should produce significant content
        assert len(md) > 10000


@skip_no_pptx_fixtures
class TestTestimony:
    """Test Testimony-Mulvey-2013-03-22.pptx — 25 slides, SmartArt, images."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(KELVIN_PPTX_FIXTURES / "Testimony-Mulvey-2013-03-22.pptx")

    def test_slide_count(self):
        assert len(self.doc.body) == 25

    def test_title_extraction(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        assert "Surface Transportation Board" in md

    def test_bullet_lists(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        # Should have bullet content
        assert "Railroad" in md

    def test_images_extracted(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        # WMF and other images should appear
        assert "![" in md

    def test_markdown_substantial(self):
        from kaos_content.serializers.markdown import serialize_markdown

        md = serialize_markdown(self.doc)
        assert len(md) > 3000


@skip_without_external_fixture("pptx", "CIPLA_CLEVELAND_BAR_DEC_2023.pptx")
class TestLargePresentation:
    """Test CIPLA_CLEVELAND_BAR_DEC_2023.pptx — 59 slides, ~11 MB.

    Too large to vendor; opt in by pointing
    ``KAOS_OFFICE_EXTERNAL_FIXTURES_DIR`` at a directory containing
    ``pptx/CIPLA_CLEVELAND_BAR_DEC_2023.pptx``.
    """

    def test_parse_performance(self):
        """Large file should parse in under 5 seconds."""
        import time

        from kaos_office.pptx.reader import parse_pptx

        path = external_fixture("pptx", "CIPLA_CLEVELAND_BAR_DEC_2023.pptx")
        assert path is not None  # guarded by class-level skip
        t0 = time.perf_counter()
        doc = parse_pptx(path)
        elapsed = time.perf_counter() - t0

        assert elapsed < 5.0, f"Parsing took {elapsed:.1f}s, expected < 5s"
        assert len(doc.body) >= 55  # Most of 59 slides should have content

    def test_slide_listing(self):
        from kaos_office.pptx.reader import list_slides

        path = external_fixture("pptx", "CIPLA_CLEVELAND_BAR_DEC_2023.pptx")
        assert path is not None  # guarded by class-level skip
        slides = list_slides(path)
        assert len(slides) == 59


@skip_no_pptx_fixtures
class TestVariousPresentations:
    """Parametrized test across multiple fixtures."""

    @pytest.fixture(
        params=[
            "Hello-World.pptx",
            "Status report.pptx",
            "early-mobility-icu-slides.pptx",
            "testimony-poster-template.pptx",
        ]
    )
    def pptx_path(self, request):
        path = KELVIN_PPTX_FIXTURES / request.param
        if not path.exists():
            pytest.skip(f"{request.param} not available")
        return path

    def test_parses_without_error(self, pptx_path):
        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(pptx_path)
        assert doc is not None
        assert len(doc.body) > 0

    def test_produces_text(self, pptx_path):
        from kaos_content.serializers.text import serialize_text

        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(pptx_path)
        text = serialize_text(doc)
        assert len(text) > 10


class TestStressFixtures:
    """Test stress/edge case fixtures from Apache POI and python-pptx."""

    @pytest.fixture(
        params=[
            "bar-chart.pptx",
            "pie-chart.pptx",
            "table_test.pptx",
            "table_test2.pptx",
            "comment.pptx",
            "minimal.pptx",
            "no-slides.pptx",
            "no-core-props.pptx",
            "missing_rels_item.pptx",
            "SmartArt.pptx",
        ]
    )
    def stress_path(self, request):
        path = PPTX_STRESS_FIXTURES / request.param
        if not path.exists():
            pytest.skip(f"{request.param} not available")
        return path

    def test_parses_without_crash(self, stress_path):
        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(stress_path)
        assert doc is not None

    def test_produces_markdown(self, stress_path):
        from kaos_content.serializers.markdown import serialize_markdown

        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(stress_path)
        md = serialize_markdown(doc)
        assert isinstance(md, str)


class TestClusterFuzz:
    """Test that malformed files raise errors gracefully."""

    def test_clusterfuzz_minimal(self):
        path = PPTX_STRESS_FIXTURES / "clusterfuzz-minimal.pptx"
        if not path.exists():
            pytest.skip("clusterfuzz-minimal.pptx not available")

        # python-pptx raises PackageNotFoundError for malformed files
        from pptx.exc import PackageNotFoundError

        from kaos_office.pptx.reader import parse_pptx

        with pytest.raises((PackageNotFoundError, KeyError, ValueError, OSError)):
            parse_pptx(path)

    def test_grouping_issues(self):
        """grouping_issues.pptx may have problematic shape groups."""
        path = PPTX_STRESS_FIXTURES / "grouping_issues.pptx"
        if not path.exists():
            pytest.skip("grouping_issues.pptx not available")

        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(path)
        assert doc is not None
