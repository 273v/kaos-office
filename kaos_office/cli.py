"""kaos-office CLI — extract, search, and inspect Office documents.

Usage:
    kaos-office extract FILE [--format markdown|text|json] [--output FILE] [--json]
    kaos-office search FILE QUERY [--top-k N] [--level paragraph|sentence] [--json]
    kaos-office metadata FILE [--json]
    kaos-office pptx-extract FILE [--format markdown|text|json] [--output FILE] [--json]
    kaos-office pptx-slides FILE [--json]
    kaos-office pptx-slide FILE SLIDE_NUMBER [--json]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="kaos-office",
        description="Office document extraction for KAOS",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- DOCX: extract ---
    p_extract = subparsers.add_parser("extract", help="Extract DOCX content")
    p_extract.add_argument("file", help="Path to DOCX file")
    p_extract.add_argument(
        "--format",
        choices=["markdown", "text", "json"],
        default="markdown",
        help="Output format (default: markdown)",
    )
    p_extract.add_argument("--output", "-o", help="Write to file instead of stdout")
    p_extract.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- DOCX: search ---
    p_search = subparsers.add_parser("search", help="Search within a DOCX document")
    p_search.add_argument("file", help="Path to DOCX file")
    p_search.add_argument("query", help="Search query")
    p_search.add_argument("--top-k", type=int, default=10, help="Max results (default: 10)")
    p_search.add_argument(
        "--level",
        choices=["paragraph", "sentence"],
        default="paragraph",
        help="Search granularity (default: paragraph)",
    )
    p_search.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- DOCX: metadata ---
    p_meta = subparsers.add_parser("metadata", help="Show DOCX metadata")
    p_meta.add_argument("file", help="Path to DOCX file")
    p_meta.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- PPTX: extract ---
    p_pptx = subparsers.add_parser("pptx-extract", help="Extract PPTX content")
    p_pptx.add_argument("file", help="Path to PPTX file")
    p_pptx.add_argument(
        "--format",
        choices=["markdown", "text", "json"],
        default="markdown",
        help="Output format (default: markdown)",
    )
    p_pptx.add_argument("--output", "-o", help="Write to file instead of stdout")
    p_pptx.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- PPTX: list slides ---
    p_slides = subparsers.add_parser("pptx-slides", help="List PPTX slides")
    p_slides.add_argument("file", help="Path to PPTX file")
    p_slides.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- PPTX: get slide ---
    p_slide = subparsers.add_parser("pptx-slide", help="Get text from a specific slide")
    p_slide.add_argument("file", help="Path to PPTX file")
    p_slide.add_argument("slide_number", type=int, help="Slide number (1-based)")
    p_slide.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- XLSX: extract ---
    p_xlsx = subparsers.add_parser("xlsx-extract", help="Extract XLSX content as tabular data")
    p_xlsx.add_argument("file", help="Path to XLSX file")
    p_xlsx.add_argument(
        "--format",
        choices=["csv", "tsv", "markdown", "json"],
        default="tsv",
        help="Output format (default: tsv)",
    )
    p_xlsx.add_argument("--sheet", help="Specific sheet name (default: all)")
    p_xlsx.add_argument(
        "--header-row", type=int, default=0, help="Header row, 0-based (default: 0)"
    )
    p_xlsx.add_argument("--max-rows", type=int, help="Max data rows per sheet")
    p_xlsx.add_argument("--output", "-o", help="Write to file instead of stdout")
    p_xlsx.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- XLSX: list sheets ---
    p_xlsx_sheets = subparsers.add_parser("xlsx-sheets", help="List sheets in XLSX file")
    p_xlsx_sheets.add_argument("file", help="Path to XLSX file")
    p_xlsx_sheets.add_argument(
        "--json", dest="json_output", action="store_true", help="JSON envelope"
    )

    # --- XLSX: get sheet ---
    p_xlsx_sheet = subparsers.add_parser("xlsx-sheet", help="Extract a single sheet")
    p_xlsx_sheet.add_argument("file", help="Path to XLSX file")
    p_xlsx_sheet.add_argument("sheet", help="Sheet name")
    p_xlsx_sheet.add_argument(
        "--format",
        choices=["csv", "tsv", "markdown", "json"],
        default="tsv",
        help="Output format (default: tsv)",
    )
    p_xlsx_sheet.add_argument("--max-rows", type=int, default=100, help="Max rows (default: 100)")
    p_xlsx_sheet.add_argument(
        "--json", dest="json_output", action="store_true", help="JSON envelope"
    )

    args = parser.parse_args(argv)

    handlers = {
        "extract": _cmd_extract,
        "search": _cmd_search,
        "metadata": _cmd_metadata,
        "pptx-extract": _cmd_pptx_extract,
        "pptx-slides": _cmd_pptx_slides,
        "pptx-slide": _cmd_pptx_slide,
        "xlsx-extract": _cmd_xlsx_extract,
        "xlsx-sheets": _cmd_xlsx_sheets,
        "xlsx-sheet": _cmd_xlsx_sheet,
    }
    try:
        handlers[args.command](args)
    except FileNotFoundError as exc:
        _error(str(exc))
    except Exception as exc:
        _error(f"Error: {exc}")


def _cmd_extract(args: argparse.Namespace) -> None:
    """Handle the DOCX extract subcommand."""
    path = _validate_file(args.file)

    from kaos_office.docx.reader import parse_docx

    doc = parse_docx(path)

    if args.format == "markdown":
        from kaos_content.serializers.markdown import serialize_markdown

        output = serialize_markdown(doc)
    elif args.format == "text":
        from kaos_content.serializers.text import serialize_text

        output = serialize_text(doc)
    elif args.format == "json":
        output = doc.model_dump_json(indent=2)
    else:
        output = ""

    if args.json_output:
        envelope = {
            "command": "extract",
            "file": args.file,
            "format": args.format,
            "blocks": len(doc.body),
            "content": output,
        }
        _json_out(envelope)
    elif args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output)


def _cmd_search(args: argparse.Namespace) -> None:
    """Handle the DOCX search subcommand."""
    path = _validate_file(args.file)

    from kaos_content.search import search_document

    from kaos_office.docx.reader import parse_docx

    doc = parse_docx(path)
    results = search_document(doc, args.query, top_k=args.top_k, level=args.level)

    if args.json_output:
        envelope = {
            "command": "search",
            "file": args.file,
            "query": results.query,
            "total_matches": results.total_matches,
            "has_more": results.has_more,
            "results": [
                {
                    "text": r.text,
                    "score": round(r.score, 4),
                    "block_ref": r.block_ref,
                    "section_title": r.section_title,
                }
                for r in results.results
            ],
        }
        _json_out(envelope)
    else:
        print(f"Search: {results.query}")
        print(f"Found: {results.total_matches} matches")
        if results.has_more:
            print(f"(showing top {len(results.results)})")
        print()
        for i, r in enumerate(results.results, 1):
            score = f"{r.score:.4f}"
            text = r.text[:200].replace("\n", " ")
            print(f"  {i}. [{score}] {text}")
            if r.section_title:
                print(f"     Section: {r.section_title}")
            print()


def _cmd_metadata(args: argparse.Namespace) -> None:
    """Handle the DOCX metadata subcommand."""
    path = _validate_file(args.file)

    from kaos_office.docx.metadata import DocxMetadata
    from kaos_office.opc.package import OPCPackage

    with OPCPackage.open(path) as pkg:
        core_xml = pkg.read_part("docProps/core.xml") if pkg.has_part("docProps/core.xml") else None
        app_xml = pkg.read_part("docProps/app.xml") if pkg.has_part("docProps/app.xml") else None
        meta = DocxMetadata.from_xml(core_xml, app_xml)

    if args.json_output:
        envelope = {"command": "metadata", "file": args.file, **meta.to_dict()}
        _json_out(envelope)
    else:
        d = meta.to_dict()
        if not d:
            print("No metadata found.")
            return
        max_key = max(len(k) for k in d)
        for k, v in d.items():
            print(f"  {k:<{max_key + 2}}{v}")


def _cmd_pptx_extract(args: argparse.Namespace) -> None:
    """Handle the PPTX extract subcommand."""
    path = _validate_file(args.file)

    from kaos_office.pptx.reader import parse_pptx

    doc = parse_pptx(path)

    if args.format == "markdown":
        from kaos_content.serializers.markdown import serialize_markdown

        output = serialize_markdown(doc)
    elif args.format == "text":
        from kaos_content.serializers.text import serialize_text

        output = serialize_text(doc)
    elif args.format == "json":
        output = doc.model_dump_json(indent=2)
    else:
        output = ""

    if args.json_output:
        envelope = {
            "command": "pptx-extract",
            "file": args.file,
            "format": args.format,
            "slides": len(doc.body),
            "content": output,
        }
        _json_out(envelope)
    elif args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output)


def _cmd_pptx_slides(args: argparse.Namespace) -> None:
    """Handle the PPTX list-slides subcommand."""
    path = _validate_file(args.file)

    from kaos_office.pptx.reader import list_slides

    slides = list_slides(path)

    if args.json_output:
        envelope = {
            "command": "pptx-slides",
            "file": args.file,
            "slide_count": len(slides),
            "slides": slides,
        }
        _json_out(envelope)
    else:
        print(f"Slides: {len(slides)}")
        print()
        for s in slides:
            title = s["title"] or "(untitled)"
            notes = " [notes]" if s["has_notes"] else ""
            print(f"  {s['slide_number']:>3}. {title}  ({s['shape_count']} shapes){notes}")


def _cmd_pptx_slide(args: argparse.Namespace) -> None:
    """Handle the PPTX get-slide subcommand."""
    path = _validate_file(args.file)

    from kaos_office.pptx.reader import get_slide_text

    text = get_slide_text(path, args.slide_number)

    if args.json_output:
        envelope = {
            "command": "pptx-slide",
            "file": args.file,
            "slide_number": args.slide_number,
            "content": text,
        }
        _json_out(envelope)
    else:
        print(text)


# --- Helpers ---


def _validate_file(path_str: str) -> Path:
    """Validate that a file exists."""
    p = Path(path_str)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path_str}. Verify the path is correct.")
    return p


def _json_out(data: dict) -> None:
    """Write JSON to stdout."""
    print(json.dumps(data, indent=2, ensure_ascii=False, default=str))


def _error(msg: str) -> None:
    """Print error to stderr and exit."""
    print(msg, file=sys.stderr)
    sys.exit(1)


# --- XLSX command handlers ---


def _cmd_xlsx_extract(args: argparse.Namespace) -> None:
    path = _validate_file(args.file)
    from kaos_content.serializers.tabular import (
        serialize_csv,
        serialize_json_records,
        serialize_markdown_table,
        serialize_tsv,
    )

    from kaos_office.xlsx.reader import parse_xlsx

    sheets = [args.sheet] if args.sheet else None
    doc = parse_xlsx(path, sheets=sheets, header_row=args.header_row, max_rows=args.max_rows)

    if args.json_output:
        _json_out(
            {
                "command": "xlsx-extract",
                "file": Path(args.file).name,
                "table_count": len(doc.tables),
                "total_rows": sum(t.row_count for t in doc.tables),
                "tables": [
                    {
                        "name": t.name,
                        "row_count": t.row_count,
                        "columns": [
                            {"name": c.name, "type": c.column_type.value} for c in t.columns
                        ],
                    }
                    for t in doc.tables
                ],
            }
        )
        return

    formatters = {
        "tsv": serialize_tsv,
        "csv": serialize_csv,
        "json": serialize_json_records,
        "markdown": serialize_markdown_table,
    }
    parts = []
    for table in doc.tables:
        if len(doc.tables) > 1:
            parts.append(f"--- {table.name} ({table.row_count} rows) ---\n")
        parts.append(formatters[args.format](table))
    output = "\n".join(parts)
    if args.output:
        Path(args.output).write_text(output, encoding="utf-8")
        print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output, end="")


def _cmd_xlsx_sheets(args: argparse.Namespace) -> None:
    path = _validate_file(args.file)
    from kaos_office.xlsx.reader import list_sheets

    sheets = list_sheets(path)
    if args.json_output:
        _json_out(
            {
                "command": "xlsx-sheets",
                "file": Path(args.file).name,
                "sheets": sheets,
                "count": len(sheets),
            }
        )
        return
    for s in sheets:
        vis = "" if s["visible"] else " (hidden)"
        print(f"  {s['name']}: {s['rows']} rows x {s['columns']} cols{vis}")


def _cmd_xlsx_sheet(args: argparse.Namespace) -> None:
    path = _validate_file(args.file)
    from kaos_content.serializers.tabular import (
        serialize_csv,
        serialize_json_records,
        serialize_markdown_table,
        serialize_tsv,
    )

    from kaos_office.xlsx.reader import parse_xlsx

    doc = parse_xlsx(path, sheets=[args.sheet], max_rows=args.max_rows)
    if not doc.tables:
        _error(f"Sheet '{args.sheet}' not found. Use xlsx-sheets to list available sheets.")

    table = doc.tables[0]
    if args.json_output:
        _json_out(
            {
                "command": "xlsx-sheet",
                "file": Path(args.file).name,
                "sheet": table.name,
                "row_count": table.row_count,
                "columns": [{"name": c.name, "type": c.column_type.value} for c in table.columns],
            }
        )
        return

    formatters = {
        "tsv": serialize_tsv,
        "csv": serialize_csv,
        "json": serialize_json_records,
        "markdown": serialize_markdown_table,
    }
    print(formatters[args.format](table), end="")
