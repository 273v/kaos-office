"""Battle tests for PPTX extraction quality.

Tests against generated fixtures that exercise every code path.
Run `uv run python tests/generate_battle_test_pptx.py` to regenerate fixtures.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from kaos_content.serializers.markdown import serialize_markdown
from kaos_content.serializers.text import serialize_text

BATTLE_DIR = Path(__file__).parent.parent / "fixtures" / "pptx" / "battle"


def _skip_if_missing(name: str) -> Path:
    path = BATTLE_DIR / name
    if not path.exists():
        pytest.skip(f"{name} not generated — run generate_battle_test_pptx.py")
    return path


class TestRichText:
    """Verify text formatting extraction quality."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("rich-text.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_bold(self):
        assert "**This is bold text**" in self.md

    def test_italic(self):
        assert "*This is italic text*" in self.md

    def test_bold_italic(self):
        assert "bold and italic" in self.md
        assert "**" in self.md

    def test_mixed_runs(self):
        assert "Normal, " in self.md
        assert "**bold**" in self.md
        assert "*italic*" in self.md
        assert "in one paragraph" in self.md

    def test_hyperlink_resolved(self):
        assert "[Click here for link](https://273ventures.com)" in self.md

    def test_unicode(self):
        text = serialize_text(self.doc)
        assert "café" in text
        assert "日本語" in text
        assert "العربية" in text

    def test_special_chars(self):
        text = serialize_text(self.doc)
        assert "©" in text
        assert "™" in text

    def test_empty_paragraph_handled(self):
        text = serialize_text(self.doc)
        assert "After empty paragraph" in text


class TestBullets:
    """Verify bullet list extraction quality."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("bullets.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_heading(self):
        assert "# Bullet Lists" in self.md

    def test_top_level_content(self):
        # Level 0 items without explicit lvl attribute are plain paragraphs
        assert "First bullet" in self.md
        assert "Second bullet" in self.md
        assert "Third bullet" in self.md

    def test_nested_bullets(self):
        # Level 1+ items with explicit lvl become bullets
        assert "Nested item one" in self.md
        assert "Nested item two" in self.md

    def test_numbered_list(self):
        assert "1. Step one" in self.md
        assert "2. Step two" in self.md
        assert "3. Step three" in self.md


class TestTables:
    """Verify table extraction quality."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("tables.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_simple_table_headers(self):
        assert "Name" in self.md
        assert "Role" in self.md
        assert "Department" in self.md

    def test_simple_table_data(self):
        assert "Alice" in self.md
        assert "Engineer" in self.md
        assert "Product" in self.md

    def test_merged_table(self):
        assert "Merged Header" in self.md
        assert "Single" in self.md

    def test_three_slides(self):
        assert len(self.doc.body) == 3

    def test_table_count(self):
        def count_tables(blocks):
            n = 0
            for b in blocks:
                if b.node_type == "table":
                    n += 1
                if hasattr(b, "children"):
                    n += count_tables(b.children)
            return n

        assert count_tables(self.doc.body) == 3


class TestCharts:
    """Verify chart linearization quality."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("charts.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_bar_chart_title(self):
        assert "Quarterly Performance" in self.md

    def test_bar_chart_categories(self):
        assert "Q1" in self.md
        assert "Q4" in self.md

    def test_bar_chart_series(self):
        assert "Revenue" in self.md
        assert "Costs" in self.md

    def test_bar_chart_values(self):
        assert "100.0" in self.md
        assert "300.0" in self.md

    def test_pie_chart(self):
        assert "Department Distribution" in self.md
        assert "Engineering" in self.md
        assert "45.0" in self.md

    def test_line_chart(self):
        assert "Jan" in self.md
        assert "May" in self.md

    def test_three_slides(self):
        assert len(self.doc.body) == 3


class TestImages:
    """Verify image extraction quality."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("images.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_image_present(self):
        assert "![" in self.md

    def test_alt_text(self):
        assert "red rectangle" in self.md

    def test_image_extension(self):
        assert ".png" in self.md
        assert ".jpg" in self.md


class TestNotes:
    """Verify speaker notes extraction."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("notes.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_notes_extracted(self):
        assert "speaker notes for slide 1" in self.md

    def test_multiline_notes(self):
        assert "multiple lines" in self.md

    def test_no_notes_slide_ok(self):
        assert "Slide Without Notes" in self.md

    def test_empty_notes_no_div(self):
        # Empty notes should not produce a speaker-notes div
        text = serialize_text(self.doc)
        lines = text.strip().split("\n")
        # "Slide With Empty Notes" should be present but no notes content after it
        assert any("Slide With Empty Notes" in line for line in lines)


class TestSpatialOrdering:
    """Verify shapes are sorted by reading position."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("spatial-ordering.pptx"))
        self.text = serialize_text(self.doc)

    def test_title_first(self):
        # Title text box is at y=0.3, should come first
        assert self.text.index("Spatial Ordering Test") < self.text.index("Shape")

    def test_shapes_ordered_by_y(self):
        # Shape 5 (y=2.0) should come before Shape 1 (y=4.0)
        pos_5 = self.text.index("Shape 5")
        pos_1 = self.text.index("Shape 1")
        assert pos_5 < pos_1


class TestPerformance:
    """Verify performance on large presentations."""

    def test_50_slides_under_1s(self):
        import time

        from kaos_office.pptx.reader import parse_pptx

        path = _skip_if_missing("50-slides.pptx")
        t0 = time.perf_counter()
        doc = parse_pptx(path)
        elapsed = time.perf_counter() - t0

        assert elapsed < 1.0, f"50-slide parse took {elapsed:.2f}s, expected < 1s"
        assert len(doc.body) == 50


class TestMixedContent:
    """Verify mixed content slide (text + bullets + table + chart)."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("mixed-content.pptx"))
        self.md = serialize_markdown(self.doc)

    def test_title(self):
        assert "Mixed Content Slide" in self.md

    def test_bullets(self):
        assert "Key finding one" in self.md

    def test_table(self):
        assert "Revenue" in self.md
        assert "$1.2M" in self.md

    def test_chart(self):
        assert "2023" in self.md
        assert "80.0" in self.md


class TestMetadata:
    """Verify metadata extraction."""

    def test_title_from_core(self):
        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(_skip_if_missing("metadata.pptx"))
        assert doc.metadata.title == "Battle Test Presentation"

    def test_author(self):
        from kaos_office.pptx.reader import parse_pptx

        doc = parse_pptx(_skip_if_missing("metadata.pptx"))
        assert doc.metadata.authors == ("Test Author",)


class TestEdgeCases:
    """Verify edge case handling."""

    @pytest.fixture(autouse=True)
    def setup(self):
        from kaos_office.pptx.reader import parse_pptx

        self.doc = parse_pptx(_skip_if_missing("edge-cases.pptx"))

    def test_four_slides(self):
        assert len(self.doc.body) == 4

    def test_long_text(self):
        text = serialize_text(self.doc)
        assert "very long paragraph" in text
        assert len(text) > 2000

    def test_single_char(self):
        text = serialize_text(self.doc)
        assert "X" in text

    def test_whitespace_only_skipped(self):
        # Slide 3 has only whitespace — should produce empty slide div
        div3 = self.doc.body[2]
        assert len(div3.children) == 0

    def test_overlapping_shapes(self):
        text = serialize_text(self.doc)
        assert "Overlapping shape 1" in text
        assert "Overlapping shape 3" in text
