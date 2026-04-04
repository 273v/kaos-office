#!/usr/bin/env python3
"""Head-to-head benchmark: kaos-office vs MarkItDown vs Mammoth.

Tests extraction quality (bold, italic, tables, merged cells, headings, lists)
and performance on real-world DOCX, PPTX, and XLSX fixtures from kelvin_office.

Run: cd kaos-office && uv run python tests/benchmark_competitors.py
"""

from __future__ import annotations

import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Fixture paths
# ---------------------------------------------------------------------------

FIXTURES = Path(__file__).parent / "fixtures"
DOCX_FILES = sorted(FIXTURES.glob("docx/*.docx"))
PPTX_FILES = [f for f in sorted(FIXTURES.glob("pptx/*.pptx")) if f.parent.name == "pptx"]
XLSX_FILES = sorted(FIXTURES.glob("xlsx/*.xlsx"))


# ---------------------------------------------------------------------------
# Quality checks
# ---------------------------------------------------------------------------


def check_quality(text: str, fmt: str) -> dict[str, bool | int]:
    """Analyze extraction quality features in output text."""
    return {
        "chars": len(text),
        "lines": text.count("\n") + 1,
        "has_bold": "**" in text or "<b>" in text.lower() or "<strong>" in text.lower(),
        "has_italic": ("*" in text and "**" not in text.replace("**", ""))
        or "<em>" in text.lower()
        or "<i>" in text.lower(),
        "has_headings": text.startswith("#") or "\n#" in text or "<h1" in text.lower(),
        "has_tables": "|" in text and "---" in text,
        "has_lists": "\n- " in text or "\n* " in text or "\n1. " in text,
    }


# ---------------------------------------------------------------------------
# Extractors
# ---------------------------------------------------------------------------


def extract_kaos(path: Path) -> str:
    """Extract with kaos-office."""
    from kaos_office import extract_to_markdown

    return extract_to_markdown(path)


def extract_markitdown(path: Path) -> str:
    """Extract with Microsoft MarkItDown."""
    from markitdown import MarkItDown

    md = MarkItDown()
    result = md.convert(str(path))
    return result.text_content


def extract_mammoth(path: Path) -> str:
    """Extract DOCX with mammoth (HTML output)."""
    import mammoth

    with path.open("rb") as f:
        result = mammoth.convert_to_markdown(f)
    return result.value


# ---------------------------------------------------------------------------
# Main benchmark
# ---------------------------------------------------------------------------


def run_benchmark() -> None:
    print("=" * 100)
    print("KAOS-OFFICE vs COMPETITORS — Quality + Performance Benchmark")
    print("=" * 100)

    # --- DOCX ---
    print("\n## DOCX Files\n")
    print(
        f"{'File':<45s} {'Library':<15s} {'Chars':>7s} {'Lines':>6s} {'Bold':>5s} {'Ital':>5s} {'Hdrs':>5s} {'Tbls':>5s} {'List':>5s} {'Time':>7s}"
    )
    print("-" * 100)

    for f in DOCX_FILES:
        extractors = [
            ("kaos", extract_kaos),
            ("markitdown", extract_markitdown),
            ("mammoth", extract_mammoth),
        ]
        for name, func in extractors:
            try:
                start = time.monotonic()
                text = func(f)
                elapsed = time.monotonic() - start
                q = check_quality(text, "docx")
                print(
                    f"{f.name:<45s} {name:<15s} {q['chars']:>7d} {q['lines']:>6d} "
                    f"{'Y' if q['has_bold'] else '-':>5s} {'Y' if q['has_italic'] else '-':>5s} "
                    f"{'Y' if q['has_headings'] else '-':>5s} {'Y' if q['has_tables'] else '-':>5s} "
                    f"{'Y' if q['has_lists'] else '-':>5s} {elapsed:>6.3f}s"
                )
            except Exception as exc:
                print(f"{f.name:<45s} {name:<15s} {'FAILED':>7s} — {exc!s:.50s}")
        print()

    # --- PPTX ---
    print("\n## PPTX Files\n")
    print(
        f"{'File':<45s} {'Library':<15s} {'Chars':>7s} {'Lines':>6s} {'Bold':>5s} {'Hdrs':>5s} {'Tbls':>5s} {'List':>5s} {'Time':>7s}"
    )
    print("-" * 100)

    for f in PPTX_FILES:
        extractors = [
            ("kaos", extract_kaos),
            ("markitdown", extract_markitdown),
        ]
        for name, func in extractors:
            try:
                start = time.monotonic()
                text = func(f)
                elapsed = time.monotonic() - start
                q = check_quality(text, "pptx")
                print(
                    f"{f.name:<45s} {name:<15s} {q['chars']:>7d} {q['lines']:>6d} "
                    f"{'Y' if q['has_bold'] else '-':>5s} "
                    f"{'Y' if q['has_headings'] else '-':>5s} {'Y' if q['has_tables'] else '-':>5s} "
                    f"{'Y' if q['has_lists'] else '-':>5s} {elapsed:>6.3f}s"
                )
            except Exception as exc:
                print(f"{f.name:<45s} {name:<15s} {'FAILED':>7s} — {exc!s:.50s}")
        print()

    # --- XLSX ---
    print("\n## XLSX Files\n")
    print(f"{'File':<55s} {'Library':<15s} {'Chars':>7s} {'Lines':>6s} {'Tbls':>5s} {'Time':>7s}")
    print("-" * 95)

    for f in XLSX_FILES:
        extractors = [
            ("kaos", extract_kaos),
            ("markitdown", extract_markitdown),
        ]
        for name, func in extractors:
            try:
                start = time.monotonic()
                text = func(f)
                elapsed = time.monotonic() - start
                q = check_quality(text, "xlsx")
                print(
                    f"{f.name:<55s} {name:<15s} {q['chars']:>7d} {q['lines']:>6d} "
                    f"{'Y' if q['has_tables'] else '-':>5s} {elapsed:>6.3f}s"
                )
            except Exception as exc:
                print(f"{f.name:<55s} {name:<15s} {'FAILED':>7s} — {exc!s:.50s}")
        print()


if __name__ == "__main__":
    run_benchmark()
