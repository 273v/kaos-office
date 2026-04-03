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

## Shared Infrastructure
- All extraction produces kaos-content `ContentDocument` AST with provenance (extractor tag, source URI).
- OPC layer (`opc/`) is format-agnostic — shared by DOCX, XLSX, PPTX. Built as proper classes to support L1 (read), L2 (write), L3 (round-trip).
- Security: ZIP bomb detection, path traversal prevention, XML bomb protection via `lxml.etree.XMLParser(resolve_entities=False)`.
- Follow the KAOS Python QA process: `ruff format`, `ruff check --fix`, `ty check`, `pytest`.
- 8 MCP tools total (5 DOCX + 3 PPTX). Register with `register_office_tools(runtime)`.
- All MCP tools use shared `_OFFICE_ANNOTATIONS` (readOnly, idempotent, !destructive, !openWorld).
- Tool error messages must include recovery guidance (what went wrong + how to fix it + alternative tool).
- `search_document()` is imported from kaos-content (canonical, shared with kaos-pdf and kaos-web).
- CLI: 6 subcommands (extract, search, metadata for DOCX; pptx-extract, pptx-slides, pptx-slide for PPTX).
- **Never add AGPL/GPL dependencies.** This is a proprietary codebase.
