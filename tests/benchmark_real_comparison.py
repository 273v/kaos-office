#!/usr/bin/env python3
"""Real head-to-head comparison: kaos-office vs competitors on real files.

Compares extraction quality AND speed on the full kelvin fixture corpus.
Measures: time, output length, structural features detected, actual content samples.

Run with: uv run python tests/benchmark_real_comparison.py
"""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

DOCX_DIR = Path("/home/mjbommar/projects/273v/kelvin-modules/kelvin_office/tests/resources/docx")
PPTX_DIR = Path("/home/mjbommar/projects/273v/kelvin-modules/kelvin_office/tests/resources/pptx")

# ═══════════════════════════════════════════════════════════════════════════════
# DOCX BENCHMARKS
# ═══════════════════════════════════════════════════════════════════════════════

DOCX_FILES = [
    "MultiParagraphSample.docx",
    "Footnote.docx",
    "Toro 2022 Term Loan.docx",
    "Toro 2022 Term Loan - Redline v1.docx",
    "Toro 2022 Term Loan - Comments.docx",
    "CheeseSample.docx",
    "bcfp_consumer-rights-summary_2018-09.docx",
    "MCSRedline10312022.docx",
]


def bench_kaos_docx(path: Path) -> dict:
    from kaos_content.serializers.markdown import serialize_markdown
    from kaos_content.serializers.text import serialize_text
    from kaos_office.docx.reader import parse_docx

    t0 = time.perf_counter()
    doc = parse_docx(path)
    parse_ms = (time.perf_counter() - t0) * 1000

    t0 = time.perf_counter()
    md = serialize_markdown(doc)
    md_ms = (time.perf_counter() - t0) * 1000

    text = serialize_text(doc)

    return {
        "tool": "kaos-office",
        "parse_ms": round(parse_ms, 1),
        "total_ms": round(parse_ms + md_ms, 1),
        "md_chars": len(md),
        "text_chars": len(text),
        "blocks": len(doc.body),
        "headings": sum(1 for b in doc.body if b.node_type == "heading"),
        "tables": sum(1 for b in doc.body if b.node_type == "table"),
        "lists": sum(1 for b in doc.body if b.node_type in ("bullet_list", "ordered_list")),
        "footnotes": len(doc.footnotes),
        "annotations": len(doc.annotations),
        "has_bold": "**" in md,
        "has_italic": ("*" in md and md.count("*") > md.count("**") * 2),
        "md_sample": md[:300],
    }


def bench_mammoth_docx(path: Path) -> dict:
    import mammoth

    t0 = time.perf_counter()
    with path.open("rb") as f:
        result = mammoth.convert_to_markdown(f)
    total_ms = (time.perf_counter() - t0) * 1000

    md = result.value
    return {
        "tool": "mammoth",
        "total_ms": round(total_ms, 1),
        "md_chars": len(md),
        "warnings": len(result.messages),
        "has_bold": "**" in md,
        "has_italic": ("*" in md and md.count("*") > md.count("**") * 2),
        "md_sample": md[:300],
    }


def bench_markitdown_docx(path: Path) -> dict:
    from markitdown import MarkItDown

    converter = MarkItDown()
    t0 = time.perf_counter()
    result = converter.convert(str(path))
    total_ms = (time.perf_counter() - t0) * 1000

    md = result.text_content
    return {
        "tool": "markitdown",
        "total_ms": round(total_ms, 1),
        "md_chars": len(md),
        "has_bold": "**" in md,
        "has_italic": ("*" in md and md.count("*") > md.count("**") * 2),
        "md_sample": md[:300],
    }


def bench_docling_docx(path: Path) -> dict:
    from docling.document_converter import DocumentConverter

    converter = DocumentConverter()
    t0 = time.perf_counter()
    result = converter.convert(str(path))
    total_ms = (time.perf_counter() - t0) * 1000

    md = result.document.export_to_markdown()

    return {
        "tool": "docling",
        "total_ms": round(total_ms, 1),
        "md_chars": len(md),
        "has_bold": "**" in md,
        "has_italic": ("*" in md and md.count("*") > md.count("**") * 2),
        "md_sample": md[:300],
    }


# ═══════════════════════════════════════════════════════════════════════════════
# PPTX BENCHMARKS
# ═══════════════════════════════════════════════════════════════════════════════

PPTX_FILES = [
    "Hello-World.pptx",
    "IEO2021_ChartLibrary_Industrial.pptx",
    "Testimony-Mulvey-2013-03-22.pptx",
    "CIPLA_CLEVELAND_BAR_DEC_2023.pptx",
    "Status report.pptx",
    "early-mobility-icu-slides.pptx",
    "redac-sas-201609-uas-concept-maturation.pptx",
]


def bench_kaos_pptx(path: Path) -> dict:
    from kaos_content.serializers.markdown import serialize_markdown
    from kaos_content.serializers.text import serialize_text
    from kaos_office.pptx.reader import parse_pptx

    t0 = time.perf_counter()
    doc = parse_pptx(path)
    parse_ms = (time.perf_counter() - t0) * 1000

    t0 = time.perf_counter()
    md = serialize_markdown(doc)
    md_ms = (time.perf_counter() - t0) * 1000

    text = serialize_text(doc)

    # Count nested features
    def count_type(blocks, ntype):
        n = 0
        for b in blocks:
            if b.node_type == ntype:
                n += 1
            if hasattr(b, "children"):
                n += count_type(b.children, ntype)
        return n

    return {
        "tool": "kaos-office",
        "parse_ms": round(parse_ms, 1),
        "total_ms": round(parse_ms + md_ms, 1),
        "md_chars": len(md),
        "text_chars": len(text),
        "slides": len(doc.body),
        "headings": count_type(doc.body, "heading"),
        "tables": count_type(doc.body, "table"),
        "images": count_type(doc.body, "figure"),
        "lists": count_type(doc.body, "bullet_list") + count_type(doc.body, "ordered_list"),
        "has_bold": "**" in md,
        "md_sample": md[:300],
    }


def bench_markitdown_pptx(path: Path) -> dict:
    from markitdown import MarkItDown

    converter = MarkItDown()
    t0 = time.perf_counter()
    result = converter.convert(str(path))
    total_ms = (time.perf_counter() - t0) * 1000

    md = result.text_content
    return {
        "tool": "markitdown",
        "total_ms": round(total_ms, 1),
        "md_chars": len(md),
        "has_bold": "**" in md,
        "tables": md.count("| ---"),
        "images": md.count("!["),
        "md_sample": md[:300],
    }


def bench_docling_pptx(path: Path) -> dict:
    from docling.document_converter import DocumentConverter

    converter = DocumentConverter()
    t0 = time.perf_counter()
    result = converter.convert(str(path))
    total_ms = (time.perf_counter() - t0) * 1000

    md = result.document.export_to_markdown()

    return {
        "tool": "docling",
        "total_ms": round(total_ms, 1),
        "md_chars": len(md),
        "has_bold": "**" in md,
        "md_sample": md[:300],
    }


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════


def main():
    print("=" * 90)
    print("REAL FILE COMPARISON: kaos-office vs competitors")
    print("=" * 90)

    # ── DOCX ──
    print("\n" + "═" * 90)
    print("DOCX EXTRACTION")
    print("═" * 90)

    docx_results = []
    for filename in DOCX_FILES:
        path = DOCX_DIR / filename
        if not path.exists():
            continue

        size_kb = path.stat().st_size / 1024
        print(f"\n{'─' * 90}")
        print(f"  {filename} ({size_kb:.0f} KB)")
        print(f"{'─' * 90}")

        row = {"file": filename, "size_kb": round(size_kb)}

        for bench_fn, name in [
            (bench_kaos_docx, "kaos-office"),
            (bench_mammoth_docx, "mammoth"),
            (bench_markitdown_docx, "markitdown"),
            (bench_docling_docx, "docling"),
        ]:
            try:
                r = bench_fn(path)
                ms = r["total_ms"]
                chars = r["md_chars"]
                feats = []
                if r.get("has_bold"):
                    feats.append("bold")
                if r.get("has_italic"):
                    feats.append("italic")
                for k in ["headings", "tables", "lists", "footnotes", "annotations"]:
                    if r.get(k, 0) > 0:
                        feats.append(f"{k}={r[k]}")
                feat_str = ", ".join(feats) if feats else "(none)"
                print(f"  {name:14s} {ms:>8.0f}ms  {chars:>8,} chars  {feat_str}")
                row[name] = r
            except Exception as e:
                print(f"  {name:14s}    ERROR: {type(e).__name__}: {str(e)[:80]}")
                row[name] = {"error": str(e)}

        docx_results.append(row)

    # ── PPTX ──
    print("\n" + "═" * 90)
    print("PPTX EXTRACTION")
    print("═" * 90)

    pptx_results = []
    for filename in PPTX_FILES:
        path = PPTX_DIR / filename
        if not path.exists():
            continue

        size_kb = path.stat().st_size / 1024
        print(f"\n{'─' * 90}")
        print(f"  {filename} ({size_kb:.0f} KB)")
        print(f"{'─' * 90}")

        row = {"file": filename, "size_kb": round(size_kb)}

        for bench_fn, name in [
            (bench_kaos_pptx, "kaos-office"),
            (bench_markitdown_pptx, "markitdown"),
            (bench_docling_pptx, "docling"),
        ]:
            try:
                r = bench_fn(path)
                ms = r["total_ms"]
                chars = r["md_chars"]
                feats = []
                if r.get("has_bold"):
                    feats.append("bold")
                for k in ["slides", "headings", "tables", "images", "lists"]:
                    if r.get(k, 0) > 0:
                        feats.append(f"{k}={r[k]}")
                feat_str = ", ".join(feats) if feats else "(none)"
                print(f"  {name:14s} {ms:>8.0f}ms  {chars:>8,} chars  {feat_str}")
                row[name] = r
            except Exception as e:
                print(f"  {name:14s}    ERROR: {type(e).__name__}: {str(e)[:80]}")
                row[name] = {"error": str(e)}

        pptx_results.append(row)

    # ── QUALITY SAMPLES ──
    print("\n" + "═" * 90)
    print("QUALITY SAMPLES: First 300 chars of markdown from each tool")
    print("═" * 90)

    # Pick 2 interesting DOCX files for sample comparison
    for filename in ["Toro 2022 Term Loan - Redline v1.docx", "Footnote.docx"]:
        path = DOCX_DIR / filename
        if not path.exists():
            continue
        print(f"\n{'─' * 90}")
        print(f"  {filename}")
        print(f"{'─' * 90}")
        for bench_fn, name in [
            (bench_kaos_docx, "kaos-office"),
            (bench_mammoth_docx, "mammoth"),
            (bench_markitdown_docx, "markitdown"),
        ]:
            try:
                r = bench_fn(path)
                sample = r["md_sample"].replace("\n", "\\n")
                print(f"\n  [{name}]")
                print(f"  {sample}")
            except Exception:
                pass

    # Pick 2 interesting PPTX files
    for filename in ["IEO2021_ChartLibrary_Industrial.pptx", "Testimony-Mulvey-2013-03-22.pptx"]:
        path = PPTX_DIR / filename
        if not path.exists():
            continue
        print(f"\n{'─' * 90}")
        print(f"  {filename}")
        print(f"{'─' * 90}")
        for bench_fn, name in [
            (bench_kaos_pptx, "kaos-office"),
            (bench_markitdown_pptx, "markitdown"),
        ]:
            try:
                r = bench_fn(path)
                sample = r["md_sample"].replace("\n", "\\n")
                print(f"\n  [{name}]")
                print(f"  {sample}")
            except Exception:
                pass

    print(f"\n{'═' * 90}")
    print("DONE")


if __name__ == "__main__":
    main()
