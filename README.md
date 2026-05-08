# kaos-office

> **Part of [Kelvin Agentic OS](https://kelvin.legal) (KAOS)** — open agentic
> infrastructure for legal work, built by
> [273 Ventures](https://273ventures.com).
> See the [full KAOS package map](https://github.com/273v) for the rest of the stack.

[![PyPI - Version](https://img.shields.io/pypi/v/kaos-office)](https://pypi.org/project/kaos-office/)
[![Python](https://img.shields.io/pypi/pyversions/kaos-office)](https://pypi.org/project/kaos-office/)
[![License](https://img.shields.io/pypi/l/kaos-office)](https://github.com/273v/kaos-office/blob/main/LICENSE)
[![CI](https://github.com/273v/kaos-office/actions/workflows/ci.yml/badge.svg)](https://github.com/273v/kaos-office/actions/workflows/ci.yml)

`kaos-office` is the Office-document layer of KAOS — it turns Microsoft
Office files (`.docx`, `.pptx`, `.xlsx`) into typed
[`kaos-content`](https://github.com/273v/kaos-content) AST models with
provenance, and turns those models back into round-trip-fidelity Office
files. DOCX and PPTX produce `ContentDocument` (Block / Inline flow
content with headings, paragraphs, lists, tables, footnotes,
annotations, tracked changes, and per-section page setup); XLSX produces
`TabularDocument` (typed columns over a 13-type `ColumnType` system, one
`Table` per sheet, formulas and merged ranges preserved as metadata).
The package also ships 17 read / write MCP tools and a 12-subcommand
admin CLI for agentic workflows.

The base install is intentionally small — three runtime dependencies
(`kaos-content[markdown]`, `kaos-core`, `lxml`) and no compiled native
code beyond the `lxml` wheel. Everything OOXML-shaped is parsed and
written with `lxml` directly so the read / write paths stay symmetric;
the optional extras only kick in when you want a different engine.
`[pptx]` adds `python-pptx` (MIT) for the PPTX writer; `[xlsx]`
aggregates `python-calamine` (MIT, Rust — 7-28× faster XLSX read) and
`openpyxl` (MIT, formula extraction) — pick `[xlsx-calamine]` or
`[xlsx-formulas]` individually if you want only one. We do not and
will not depend on AGPL or GPL libraries.

## Install

```bash
uv add kaos-office
# or
pip install kaos-office

# PPTX writer (python-pptx)
uv add 'kaos-office[pptx]'

# Calamine XLSX fast-path + openpyxl formula extraction
uv add 'kaos-office[xlsx]'

# BM25 sentence-level search via kaos-nlp-core
uv add 'kaos-office[nlp]'
```

`kaos-office` requires Python **3.13** or newer (3.14 is supported).
The package is pure Python — the only native code is the `lxml` wheel,
which has prebuilt wheels for Linux, macOS, and Windows on x86_64 and
arm64.

## Quick start

Read a DOCX into the AST, render it as markdown, and search it; then
read an XLSX as a typed tabular document:

```python
from kaos_office import (
    extract_to_markdown,
    parse_docx,
    parse_pptx,
    search_document,
)
from kaos_office.xlsx import list_sheets, parse_xlsx

# DOCX → ContentDocument with Block/Inline + provenance on every node
doc = parse_docx("contract.docx")
print(len(doc.body), "top-level blocks")
print(doc.metadata.title, doc.metadata.source.uri)

# Same shape for PPTX (each slide becomes a Div(classes="slide"))
deck = parse_pptx("brief.pptx")
print(deck.metadata.extra.get("slide_count"), "slides")

# AST-grounded search — paragraph-level by default
hits = search_document(doc, "indemnification", top_k=5)
for hit in hits.results:
    print(f"score={hit.score:.2f} :: {hit.text[:80]}")

# XLSX → TabularDocument (one Table per sheet, typed columns)
tab = parse_xlsx("report.xlsx")
for table in tab.tables:
    print(f"{table.name}: {table.row_count} rows × {len(table.columns)} cols")
print(list_sheets("report.xlsx"))  # cheap workbook metadata, no parse

# Format-agnostic shortcut: any of the three → markdown
print(extract_to_markdown("contract.docx")[:200])
```

Every node in the returned `ContentDocument` carries a `Provenance`
(source URI, page or slide number, char span, extractor name) so
downstream consumers — citation verifiers, redaction tooling, labelers
— can ground answers back to the original file.

## Concepts

The package is a thin, typed surface over the OOXML wire format. The
most important entries:

| Concept | What it is |
|---|---|
| **`parse_docx(path, *, track_changes=False, image_src_builder=...)`** | DOCX reader. Returns a `ContentDocument` with paragraphs, headings, lists, tables, footnotes, comments (as annotations), hyperlinks, embedded images, and per-section page setup. `track_changes=True` preserves `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` as `Span` / `Div` with `rev-*` classes plus `TRACKED_CHANGE` annotations. |
| **`parse_pptx(path)`** | PPTX reader. Each slide → `Div(classes="slide", slide_number=N)`. Uses `python-pptx` for shape traversal and falls back to OPC/lxml for SmartArt text — the only Python tool that does. Charts linearize to `Table`s with category + series columns. Speaker notes land as `Div(classes="speaker-notes")`. |
| **`parse_xlsx(path, *, sheets=None, max_rows=None, header_row=0, include_formulas=False, engine="native")`** | XLSX reader. Returns a `TabularDocument`. Default `engine="native"` is pure lxml; `engine="calamine"` switches to the Rust fast-path (`[xlsx-calamine]`). `include_formulas=True` extracts cell formulas via openpyxl (`[xlsx-formulas]`). |
| **`write_docx(doc, path)` / `write_docx_bytes(doc)`** | DOCX writer (lxml). Round-trips the DOCX feature surface: paragraphs / headings, bullet + ordered lists with proper numbering.xml, tables with grid spans, hyperlinks (with proper rels), footnotes, endnotes, comments, headers, footers, page setup, multi-section documents, embedded images (data: / file:// URIs), and SDT / content-control wrappers. |
| **`write_pptx(doc, path, *, template=None, overflow="warn"`/`"autofit"`/`"extend")`** | PPTX writer (`python-pptx`, lazy-imported with `[pptx]` install hint at call time). Auto-segments at `Heading(depth=1)`. `overflow` controls how text that may not fit a shape is handled — `"warn"` (default) emits a logger warning, `"autofit"` shrinks the font, `"extend"` grows the shape. |
| **`write_xlsx(doc, path, *, bold_headers=True, auto_width=True, freeze_header=True)` / `write_xlsx_bytes(doc)`** | XLSX writer (lxml — no extras needed). Native SpreadsheetML output with proper date formats, money formats, percentage / float / integer formats per `ColumnType`, auto-sized columns, bold header row, and frozen panes. |
| **`search_document(doc, query, *, top_k=10, level="paragraph")`** | Re-exported from `kaos-content`. AST-grounded ranked search returning `SearchResults` with `total_matches` / `has_more` for pagination. `level="sentence"` requires the `[nlp]` extra. |
| **`extract_to_markdown(path, **kwargs)`** | Format-agnostic convenience wrapper. Dispatches by extension to `parse_docx` + `serialize_markdown`, `parse_pptx` + `serialize_markdown`, or `parse_xlsx` + `serialize_tabular_markdown`. |
| **17 MCP tools** | `ParseDocxTool`, `GetDocxTextTool`, `GetDocxMarkdownTool`, `DocxMetadataTool`, `SearchDocxTool` (5 DOCX) · `ParsePptxTool`, `ListSlidesTool`, `GetSlideTool`, `GetSlideNotesTool`, `SearchPptxTool` (5 PPTX) · `ParseXlsxTool`, `ListSheetsXlsxTool`, `GetSheetXlsxTool`, `XlsxMetadataTool` (4 XLSX) · `WriteDocxTool`, `WritePptxTool`, `WriteXlsxTool` (3 writers). All readers are `readOnly` + `idempotent` + non-destructive + non-open-world; writers refuse silent overwrites unless `force=true`. Register with `register_office_tools(runtime)`. |
| **Errors (`KaosOfficeError`, `DocxExtractionError`, `PptxExtractionError`, `XlsxExtractionError`)** | Dedicated exception hierarchy. MCP tools translate these into `ToolResult.create_error()` with the documented three-part recovery hint (what / how to fix / alternative tool). |

## CLI

`kaos-office` ships two entry-point scripts. Every structured command
on the admin CLI supports `--json` for machine-readable output piped to
other agents:

```bash
kaos-office --help                                  # admin CLI
kaos-office-serve --help                            # MCP server

# DOCX
kaos-office extract contract.docx -f markdown       # AST → markdown / text / json / html
kaos-office search contract.docx "indemnification"  # AST-grounded ranked search
kaos-office metadata contract.docx --json           # title, author, page setup, sections

# PPTX
kaos-office pptx-extract brief.pptx -f markdown
kaos-office pptx-slides brief.pptx --json           # slide inventory (number, title, layout)
kaos-office pptx-slide brief.pptx 3                 # text from a single slide (1-based)

# XLSX
kaos-office xlsx-extract report.xlsx -f markdown    # tabular markdown
kaos-office xlsx-sheets report.xlsx --json          # sheet names + dimensions
kaos-office xlsx-sheet report.xlsx Revenue          # one sheet as TSV

# Writers (JSON file or '-' for stdin)
kaos-office write-docx body.json out.docx --force
kaos-office write-pptx body.json out.pptx --template brand.pptx
kaos-office write-xlsx tabular.json out.xlsx

kaos-office-serve                                   # stdio (Claude Code / Desktop)
kaos-office-serve --http --port 8000                # streamable HTTP
```

The admin CLI uses 1-based slide / page numbers (consistent with how
the file opens in any viewer) and translates internally to the
0-based indices the Python API uses. `kaos-office-serve` exposes the
17 MCP tools listed in **Concepts** above.

## Compatibility & status

| Aspect | |
|---|---|
| **Python** | 3.13, 3.14 |
| **OS** | Linux, macOS, Windows (pure-Python wheel; the only native code is the `lxml` wheel) |
| **Maturity** | Alpha (`Development Status :: 3 - Alpha`). The public API is documented in `kaos_office.__all__`. |
| **Stability policy** | Pre-1.0: minor bumps may change behaviour. Every change is documented in [`CHANGELOG.md`](CHANGELOG.md). The MCP tool surface (`kaos-office-*` names) and the `KAOS_OFFICE_*` environment-variable namespace are public API and follow the same policy. |
| **Test coverage** | 492 unit tests plus a 144-test integration tier covering DOCX / PPTX / XLSX round-trip fidelity against real-world fixtures. Bounded unit gate (`pytest tests/unit -q --no-cov`) finishes in ~30s. |
| **Type checker** | Validated with [`ty`](https://docs.astral.sh/ty/), Astral's Python type checker. |

## Companion packages

`kaos-office` is one of the packages in the
[Kelvin Agentic OS](https://kelvin.legal). The broader stack:

| Package | Layer | What it does |
|---|---|---|
| [`kaos-core`](https://github.com/273v/kaos-core) | Core | Foundational runtime, MCP-native types, registries, execution engine, VFS |
| [`kaos-content`](https://github.com/273v/kaos-content) | Core | Typed document AST: Block/Inline, provenance, views |
| [`kaos-mcp`](https://github.com/273v/kaos-mcp) | Bridge | FastMCP server, `kaos` management CLI, MCP resource templates |
| [`kaos-pdf`](https://github.com/273v/kaos-pdf) | Extraction | PDF → AST with provenance |
| [`kaos-web`](https://github.com/273v/kaos-web) | Extraction | Web extraction, browser automation, search, domain intelligence |
| [`kaos-office`](https://github.com/273v/kaos-office) | Extraction | DOCX / PPTX / XLSX readers + writers to AST |
| [`kaos-tabular`](https://github.com/273v/kaos-tabular) | Extraction | DuckDB-powered SQL analytics |
| [`kaos-source`](https://github.com/273v/kaos-source) | Data | Government + financial data connectors (Federal Register, eCFR, EDGAR, GovInfo, PACER, GLEIF) |
| [`kaos-llm-client`](https://github.com/273v/kaos-llm-client) | LLM | Multi-provider LLM transport |
| [`kaos-llm-core`](https://github.com/273v/kaos-llm-core) | LLM | Typed LLM programming (Signatures, Programs, Optimizers) |
| [`kaos-nlp-core`](https://github.com/273v/kaos-nlp-core) | Primitives (Rust) | High-performance NLP primitives |
| [`kaos-nlp-transformers`](https://github.com/273v/kaos-nlp-transformers) | ML | Dense embeddings + retrieval |
| [`kaos-graph`](https://github.com/273v/kaos-graph) | Primitives (Rust) | Graph algorithms + RDF/SPARQL |
| [`kaos-ml-core`](https://github.com/273v/kaos-ml-core) | Primitives (Rust) | Classical ML on the document AST |
| [`kaos-citations`](https://github.com/273v/kaos-citations) | Legal | Legal citation extraction, resolution, verification |
| [`kaos-agents`](https://github.com/273v/kaos-agents) | Agentic | Agent runtime, memory, recipes |
| [`kaos-reference`](https://github.com/273v/kaos-reference) | Sample | Reference module for module authors |

Packages depend on `kaos-core`; everything else is opt-in. Mix and match the
ones you need.

## Development

```bash
git clone https://github.com/273v/kaos-office
cd kaos-office
uv sync --group dev
```

Install pre-commit hooks (recommended — they run the same checks as CI on
every commit, scoped to staged files):

```bash
uvx pre-commit install
uvx pre-commit run --all-files     # one-time full sweep
```

Manual QA commands (the same set CI runs):

```bash
uv run ruff format --check kaos_office tests
uv run ruff check kaos_office tests
uv run ty check kaos_office tests
uv run pytest tests/unit -q --no-cov
```

## Build from source

```bash
uv build
uv pip install dist/*.whl
python -c "import kaos_office; print(kaos_office.__version__)"  # smoke import
```

## Contributing

Issues and pull requests are welcome. By contributing you certify the
[Developer Certificate of Origin v1.1](https://developercertificate.org/) —
sign every commit with `git commit -s`. Please open an issue before starting
on a non-trivial change so we can align on scope.

## Security

For security issues, **please do not file a public issue**. Report privately
via [GitHub Private Vulnerability Reporting](https://github.com/273v/kaos-office/security/advisories/new)
or email **security@273ventures.com**. See [SECURITY.md](SECURITY.md) for the
full disclosure policy.

## License

Apache License 2.0 — see [LICENSE](LICENSE) and [NOTICE](NOTICE).

Copyright 2026 [273 Ventures LLC](https://273ventures.com).
Built for [kelvin.legal](https://kelvin.legal).
