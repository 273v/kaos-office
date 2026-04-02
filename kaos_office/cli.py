"""kaos-office CLI — extract, search, and inspect Office documents.

Usage:
    kaos-office extract FILE [--format markdown|text|json] [--output FILE] [--json]
    kaos-office search FILE QUERY [--top-k N] [--level paragraph|sentence] [--json]
    kaos-office metadata FILE [--json]
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

    # --- extract ---
    p_extract = subparsers.add_parser("extract", help="Extract document content")
    p_extract.add_argument("file", help="Path to DOCX file")
    p_extract.add_argument(
        "--format",
        choices=["markdown", "text", "json"],
        default="markdown",
        help="Output format (default: markdown)",
    )
    p_extract.add_argument("--output", "-o", help="Write to file instead of stdout")
    p_extract.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    # --- search ---
    p_search = subparsers.add_parser("search", help="Search within a document")
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

    # --- metadata ---
    p_meta = subparsers.add_parser("metadata", help="Show document metadata")
    p_meta.add_argument("file", help="Path to DOCX file")
    p_meta.add_argument("--json", dest="json_output", action="store_true", help="JSON envelope")

    args = parser.parse_args(argv)

    handlers = {
        "extract": _cmd_extract,
        "search": _cmd_search,
        "metadata": _cmd_metadata,
    }
    try:
        handlers[args.command](args)
    except FileNotFoundError as exc:
        _error(str(exc))
    except Exception as exc:
        _error(f"Error: {exc}")


def _cmd_extract(args: argparse.Namespace) -> None:
    """Handle the extract subcommand."""
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
    """Handle the search subcommand."""
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
    """Handle the metadata subcommand."""
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


# --- Helpers ---


def _validate_file(path_str: str) -> Path:
    """Validate that a file exists."""
    p = Path(path_str)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {path_str}. Verify the path is correct.")
    return p


def _json_out(data: dict) -> None:
    """Write JSON to stdout."""
    print(json.dumps(data, indent=2, ensure_ascii=False))


def _error(msg: str) -> None:
    """Print error to stderr and exit."""
    print(msg, file=sys.stderr)
    sys.exit(1)
