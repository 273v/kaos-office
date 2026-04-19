# kaos-office Development Notes

## Required Checklists

Apply these checklist sources to every change in this module.

Python:
- `../docs/python/checklists/index.md`
- `../docs/python/checklists/01-research.md`
- `../docs/python/checklists/02-design.md`
- `../docs/python/checklists/03-implement.md`
- `../docs/python/checklists/04-test.md`
- `../docs/python/checklists/05-quality.md`
- `../docs/python/checklists/06-review.md`
- `../docs/python/checklists/07-commit.md`
- `../docs/python/checklists/08-debug.md`
- `../docs/python/checklists/09-optimize.md`
- `../docs/python/checklists/10-document.md`
- `../docs/python/checklists/11-retrieval-and-evaluation.md`
- `../docs/python/checklists/12-benchmarking.md`
- `../docs/python/checklists/13-kaos-agent-retrieval.md`

Rust-adjacent:
- `../kaos-nlp-core/docs/FUZZY_HASHING_PLAN.md` (`QA Checklist`) for Rust, PyO3, native bindings, and performance-critical boundary work
- `../kaos-nlp-core/docs/todo/API_IMPROVEMENTS_TODO.md` for Rust-adjacent backlog and API-shape guidance

## DOCX Extraction (Phase 1 — complete)
- Uses **lxml** for XML parsing — no python-docx dependency. Direct OOXML XML access gives full control for legal document fidelity.
- DOCX parser uses two-pass architecture: metadata load (styles, numbering, rels) then body walk with tag dispatch → `DocumentBuilder`.
- Style resolution walks inheritance chains with cycle detection. Heading detection checks outline level, style name pattern, then parent chain.
- Numbering resolution handles three-level indirection: numId → abstractNumId → level definitions → numFmt.
- List state tracking maintains a stack of open lists for proper begin/end nesting.
- Track changes: accept insertions (include `w:ins` content), skip deletions (`w:del`), ignore formatting changes.

## DOCX Generation (Phase 2 + 3 — complete, Phase 4-5 pending)
- Uses **lxml** for XML serialization — no python-docx dependency. Consistent with the read path.
- Entry points: `write_docx(doc, path)` and `write_docx_bytes(doc)` in `kaos_office.docx.writer`.
- AST walker serializes all kaos-content block types: Paragraph, Heading (1-6 with outlineLvl), BulletList/OrderedList (with numbering.xml), Table (header rows, col_span, grid), CodeBlock (Consolas monospace), BlockQuote, ThematicBreak, PageBreak.
- Inline serialization: Text, Strong (bold), Emphasis (italic), Strikethrough, Code (Consolas), Link (blue+underline), LineBreak, SoftBreak.
- Uses `OPCPackageWriter` (OPC L2 layer) to assemble the ZIP archive with proper content types, relationships, and parts.
- Generates styles.xml (Heading 1-6, Normal, Code, TableGrid), numbering.xml (bullet + decimal definitions, 9 levels), docProps/core.xml (Dublin Core metadata).
- kaos-content model uses `.value` for Text/Code/CodeBlock, `.content` for Cell (not `.children`). The writer handles both via getattr fallback.
- `xml:space="preserve"` uses the XML namespace (`http://www.w3.org/XML/1998/namespace`), NOT the `R` namespace. `xsi:type` on dcterms elements uses `http://www.w3.org/2001/XMLSchema-instance`.
- **Phase 3 complete**: proper hyperlink relationships (`w:hyperlink` + rId with deduped URLs), footnote/endnote write-back (`word/footnotes.xml` / `word/endnotes.xml` with required separator+continuation IDs, `w:footnoteReference` / `w:endnoteReference` runs), comment write-back (`word/comments.xml` with author/date/text, content type and rel entries added automatically).
- **Phase 4 pending**: headers, footers, sections, page setup, images (writer — reader captures dimensions).
- DOCX reader populates `Image.width` / `Image.height` from `wp:extent cx`/`cy` (EMU → points via `kaos_office.ooxml.namespace.emu_to_pt`). PPTX reader/writer round-trip dimensions through `shape.width` / `shape.height` for picture shapes; writer falls back to `Inches(8.0)` width when unset. Convention: points (1/72 in).
- **Phase 5 pending**: MCP tools, CLI integration.
- Round-trip coverage (`tests/unit/test_docx_writer.py`): 18 DOCX fixtures; `TestRoundTrip` (identity), `TestModificationRoundTrip` (parse → edit → write → re-parse → verify edit + untouched-content preservation), `TestStyleRoundTrip` (heading style + outlineLvl survive), `TestNumberingRoundTrip` (bullet + ordered list preservation).

## XLSX Generation (complete)
- Uses **lxml** for XML serialization — native SpreadsheetML, no xlsxwriter dependency for production.
- Entry points: `write_xlsx(doc, path)` and `write_xlsx_bytes(doc)` in `kaos_office.xlsx.writer`.
- Produces valid XLSX from `TabularDocument`: workbook.xml, sheetN.xml, sharedStrings.xml, styles.xml, OPC packaging.
- ColumnType → Excel format mapping (DATE→serial+numFmtId14, MONEY→$#,##0.00, PERCENTAGE→0.00%, INTEGER→#,##0, FLOAT→#,##0.00).
- Auto-sized column widths, bold header row, frozen panes.
- ISO date string coercion to serial numbers via `date_to_serial()`.
- Lists serialized as semicolon-separated strings (not JSON).

## PPTX Generation (Phase 1 + 2 — complete)
- Uses **python-pptx** (MIT, already `[pptx]` extra) — handles OPC packaging, themes, layouts.
- Entry points: `write_pptx(doc, path, template=None)` and `write_pptx_bytes(doc)` in `kaos_office.pptx.writer`.
- Auto-segmentation: each `Heading(depth=1)` starts a new slide; tables get their own slide.
- Layout selection: Title Slide (H1+H2), Title+Content (H1+body), Blank (tables, no heading).
- Block types: Heading (title/subtitle placeholders), Paragraph, BulletList/OrderedList (XML bullet/numbering via `a:buChar`/`a:buAutoNum`), Table (`add_table` shape), CodeBlock (Consolas monospace), BlockQuote, speaker notes.
- Inline formatting: Strong (bold), Emphasis (italic), Code (Consolas), Link (hyperlink.address).
- Template support: optional custom `.pptx` template for branded output.
- **Bug note**: `SlidePlaceholders.__contains__` returns False for valid indices; use try/except not `in`.
- **Phase 2 complete**: image blocks (Figure → `add_picture` with alt text), table cell merging (`col_span` → `gridSpan`/`hMerge`, `row_span` → `rowSpan`/`vMerge`), speaker notes with structured content (not just flat text). Speaker-notes / explicit-slide Div detection reads classes from `block.attr.classes`, not the block directly (fixed a pre-existing bug).

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
- **New tests should use `kaos_content.shortcuts`** (`paragraph`, `heading`, `bold`, `italic`, `link`, `bullet_list`, `ordered_list`, `table_from_rows`) instead of nested `Paragraph(children=(Text(value=...),))` constructors. Existing tests stay verbose until they're touched for another reason — no blanket rewrite.
- 17 MCP tools total (5 DOCX read + 5 PPTX read + 4 XLSX read + 3 writers). Register with `register_office_tools(runtime)`.
- Read tools use shared `_OFFICE_ANNOTATIONS` (readOnly, idempotent, !destructive, !openWorld).
- Write tools (`WriteDocxTool`, `WritePptxTool`, `WriteXlsxTool`) use `_OFFICE_WRITE_ANNOTATIONS` (!readOnly, !idempotent, !destructive, openWorld). They accept either `document_json` (inline JSON) or `document_id` (artifact id loaded via `runtime.artifacts`) and refuse to overwrite `output_path` unless `force=true`. When a `KaosContext` with a runtime is supplied, the produced file is also copied into the VFS (`office_output/{name}`) and registered as an artifact; the result includes `artifact_id`, `body_uri`, and `manifest_uri` so downstream MCP tools can chain via the artifact registry without filesystem round-trips.
- Tool error messages must include recovery guidance (what went wrong + how to fix it + alternative tool).
- `search_document()` is imported from kaos-content (canonical, shared with kaos-pdf and kaos-web).
- CLI: 12 subcommands (3 DOCX read + 3 PPTX read + 3 XLSX read + `write-docx` + `write-pptx` + `write-xlsx`). Write commands accept a JSON file or `-` (stdin), write to a positional output path, and support `--force` / `--json` (and `--template` for `write-pptx`).
- **Never add AGPL/GPL dependencies.** This is a proprietary codebase.
