#!/usr/bin/env python3
"""Benchmark kaos-office vs other DOCX-to-markdown converters.

Compares extraction quality and performance against:
  - mammoth (HTML/markdown via python-docx-like approach)
  - markitdown (Microsoft's DOCX converter)
  - python-docx (raw text extraction only)

Run with: uv run python tests/benchmark_comparison.py
"""

from __future__ import annotations

import importlib
import time
from pathlib import Path

FIXTURES = Path("/home/mjbommar/projects/273v/kelvin-modules/kelvin_office/tests/resources/docx")

BENCHMARK_FILES = [
    ("MultiParagraphSample.docx", "Simple: paragraphs, formatting, lists, comments"),
    ("Footnote.docx", "Footnotes"),
    ("Toro 2022 Term Loan.docx", "Complex: 90-page legal doc, tables, numbering"),
    ("Toro 2022 Term Loan - Redline v1.docx", "Track changes (redline)"),
    ("Toro 2022 Term Loan - Comments.docx", "Comments"),
    ("CheeseSample.docx", "Mixed: images, nested lists, formatting"),
    ("bcfp_consumer-rights-summary_2018-09.docx", "Government: structured content"),
]


def bench_kaos_office(path: Path) -> dict:
    """Benchmark kaos-office extraction."""
    from kaos_content.serializers.markdown import serialize_markdown
    from kaos_content.serializers.text import serialize_text

    from kaos_office.docx.reader import parse_docx

    t0 = time.perf_counter()
    doc = parse_docx(path)
    parse_time = time.perf_counter() - t0

    t0 = time.perf_counter()
    md = serialize_markdown(doc)
    md_time = time.perf_counter() - t0

    text = serialize_text(doc)

    return {
        "tool": "kaos-office",
        "parse_ms": round(parse_time * 1000, 1),
        "serialize_ms": round(md_time * 1000, 1),
        "total_ms": round((parse_time + md_time) * 1000, 1),
        "blocks": len(doc.body),
        "md_chars": len(md),
        "text_chars": len(text),
        "footnotes": len(doc.footnotes),
        "annotations": len(doc.annotations),
        "headings": sum(1 for b in doc.body if b.node_type == "heading"),
        "tables": sum(1 for b in doc.body if b.node_type == "table"),
        "lists": sum(1 for b in doc.body if b.node_type in ("bullet_list", "ordered_list")),
        "has_bold": "**" in md,
        "has_italic": "*" in md and "**" not in md.replace("**", ""),
        "md_preview": md[:200],
    }


def bench_mammoth(path: Path) -> dict:
    """Benchmark mammoth extraction."""
    mammoth = importlib.import_module("mammoth")

    t0 = time.perf_counter()
    with path.open("rb") as f:
        result = mammoth.convert_to_markdown(f)
    total = time.perf_counter() - t0

    md = result.value
    messages = result.messages

    return {
        "tool": "mammoth",
        "total_ms": round(total * 1000, 1),
        "md_chars": len(md),
        "warnings": len(messages),
        "has_bold": "**" in md,
        "has_italic": "*" in md and "**" not in md.replace("**", ""),
        "md_preview": md[:200],
    }


def bench_markitdown(path: Path) -> dict:
    """Benchmark markitdown extraction."""
    markitdown = importlib.import_module("markitdown")
    converter_cls = markitdown.MarkItDown

    converter = converter_cls()
    t0 = time.perf_counter()
    result = converter.convert(str(path))
    total = time.perf_counter() - t0

    md = result.text_content

    return {
        "tool": "markitdown",
        "total_ms": round(total * 1000, 1),
        "md_chars": len(md),
        "has_bold": "**" in md,
        "has_italic": "*" in md and "**" not in md.replace("**", ""),
        "md_preview": md[:200],
    }


def bench_python_docx(path: Path) -> dict:
    """Benchmark python-docx text extraction."""
    docx = importlib.import_module("docx")

    t0 = time.perf_counter()
    document = docx.Document(str(path))
    text = "\n".join(p.text for p in document.paragraphs)
    total = time.perf_counter() - t0

    return {
        "tool": "python-docx",
        "total_ms": round(total * 1000, 1),
        "text_chars": len(text),
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "md_preview": text[:200],
    }


def main():
    print("=" * 80)
    print("DOCX Extraction Benchmark: kaos-office vs competitors")
    print("=" * 80)

    for filename, description in BENCHMARK_FILES:
        path = FIXTURES / filename
        if not path.exists():
            print(f"\n--- SKIP: {filename} (not found) ---")
            continue

        size_kb = path.stat().st_size / 1024
        print(f"\n{'─' * 80}")
        print(f"File: {filename} ({size_kb:.0f} KB)")
        print(f"Description: {description}")
        print(f"{'─' * 80}")

        # Run each benchmark
        results = []

        try:
            r = bench_kaos_office(path)
            results.append(r)
            print(
                f"\n  kaos-office:  {r['total_ms']:>7.1f}ms | "
                f"{r['md_chars']:>8,} chars | "
                f"blocks={r['blocks']} headings={r['headings']} "
                f"tables={r['tables']} lists={r['lists']} "
                f"footnotes={r['footnotes']} annotations={r['annotations']}"
            )
        except Exception as e:
            print(f"\n  kaos-office:  ERROR: {e}")

        try:
            r = bench_mammoth(path)
            results.append(r)
            print(
                f"  mammoth:      {r['total_ms']:>7.1f}ms | "
                f"{r['md_chars']:>8,} chars | warnings={r['warnings']}"
            )
        except Exception as e:
            print(f"  mammoth:      ERROR: {e}")

        try:
            r = bench_markitdown(path)
            results.append(r)
            print(f"  markitdown:   {r['total_ms']:>7.1f}ms | {r['md_chars']:>8,} chars")
        except Exception as e:
            print(f"  markitdown:   ERROR: {e}")

        try:
            r = bench_python_docx(path)
            results.append(r)
            print(
                f"  python-docx:  {r['total_ms']:>7.1f}ms | "
                f"{r['text_chars']:>8,} chars (text only) | "
                f"paragraphs={r['paragraphs']} tables={r['tables']}"
            )
        except Exception as e:
            print(f"  python-docx:  ERROR: {e}")

        # Quality comparison
        print("\n  Quality features:")
        for r in results:
            tool = r["tool"]
            features = []
            if r.get("has_bold"):
                features.append("bold")
            if r.get("has_italic"):
                features.append("italic")
            if r.get("footnotes", 0) > 0:
                features.append(f"footnotes({r['footnotes']})")
            if r.get("annotations", 0) > 0:
                features.append(f"comments({r['annotations']})")
            if r.get("headings", 0) > 0:
                features.append(f"headings({r['headings']})")
            if r.get("tables", 0) > 0:
                features.append(f"tables({r['tables']})")
            if r.get("lists", 0) > 0:
                features.append(f"lists({r['lists']})")
            print(f"    {tool:15s}: {', '.join(features) if features else '(none detected)'}")

    print(f"\n{'=' * 80}")
    print("Done.")


if __name__ == "__main__":
    main()
