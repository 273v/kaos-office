# kaos-office Development Notes

## DOCX Extraction (Phase 1 — complete)
- Uses **lxml** for XML parsing — no python-docx dependency. Direct OOXML XML access gives full control for legal document fidelity.
- DOCX parser uses two-pass architecture: metadata load (styles, numbering, rels) then body walk with tag dispatch → `DocumentBuilder`.
- Style resolution walks inheritance chains with cycle detection. Heading detection checks outline level, style name pattern, then parent chain.
- Numbering resolution handles three-level indirection: numId → abstractNumId → level definitions → numFmt.
- List state tracking maintains a stack of open lists for proper begin/end nesting.
- Track changes: accept insertions (include `w:ins` content), skip deletions (`w:del`), ignore formatting changes.

## PPTX Extraction (Phase 3 — complete)
- Uses **python-pptx** (MIT, v1.0.2) for high-level shape traversal + OPC/lxml fallback for SmartArt.
- Each slide → `Div(classes="slide", slide_number=N)`. Shapes sorted by position (top, left).
- Title/center-title placeholders → Heading depth 1. Subtitle → Heading depth 2. Date/footer/slide-number → skip.
- Bullet detection via `a:buChar` (unordered) and `a:buAutoNum` (ordered) with nesting via `lvl` attribute.
- Charts linearized as Table blocks with Category column + Series columns.
- **SmartArt text extraction via OPC fallback** — parse `diagrams/data1.xml` directly. Only Python tool that does this.
- Tables handle `gridSpan`, `rowSpan`, `hMerge`, `vMerge` continuation cells.
- Speaker notes extracted as `Div(classes="speaker-notes")`.

## XLSX Extraction (Phase 2C — complete)
- Uses **python-calamine** (MIT, Rust, 7-28x faster than openpyxl) for data extraction.
- Produces `TabularDocument` (from kaos-content), not `ContentDocument` — tabular data, not flow content.
- Each worksheet → `Table` with typed columns via `ColumnType` (13 types).
- Calamine returns Python-native types (int, float, str, bool, date, datetime, time, timedelta).
- Formula extraction via openpyxl (optional `[xlsx-formulas]` extra) — stores in `table.metadata["formulas"]`.
- Merged cell ranges preserved in `table.metadata["merged_ranges"]`.
- Header row configurable (default: row 0). Rows above header are skipped.

## Shared Infrastructure
- DOCX/PPTX produce `ContentDocument`; XLSX produces `TabularDocument`. Both are kaos-content model types.
- OPC layer (`opc/`) is format-agnostic — shared by DOCX, XLSX, PPTX. Built as proper classes to support L1 (read), L2 (write), L3 (round-trip).
- Security: ZIP bomb detection, path traversal prevention, XML bomb protection via `lxml.etree.XMLParser(resolve_entities=False)`.
- Follow the KAOS Python QA process: `ruff format`, `ruff check --fix`, `ty check`, `pytest`.
- 12 MCP tools total (5 DOCX + 3 PPTX + 4 XLSX). Register with `register_office_tools(runtime)`.
- All MCP tools use shared `_OFFICE_ANNOTATIONS` (readOnly, idempotent, !destructive, !openWorld).
- Tool error messages must include recovery guidance (what went wrong + how to fix it + alternative tool).
- `search_document()` is imported from kaos-content (canonical, shared with kaos-pdf and kaos-web).
- CLI: 9 subcommands (3 DOCX + 3 PPTX + 3 XLSX).
- **Never add AGPL/GPL dependencies.** This is a proprietary codebase.
