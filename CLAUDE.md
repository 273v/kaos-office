# kaos-office Development Notes

- Uses **lxml** for XML parsing — no python-docx dependency. Direct OOXML XML access gives full control for legal document fidelity.
- All extraction produces kaos-content `ContentDocument` AST with provenance (extractor tag, source URI).
- OPC layer (`opc/`) is format-agnostic — shared by DOCX, XLSX, PPTX. Built as proper classes to support L1 (read), L2 (write), L3 (round-trip).
- DOCX parser uses two-pass architecture: metadata load (styles, numbering, rels) then body walk with tag dispatch → `DocumentBuilder`.
- Style resolution walks inheritance chains with cycle detection. Heading detection checks outline level, style name pattern, then parent chain.
- Numbering resolution handles three-level indirection: numId → abstractNumId → level definitions → numFmt.
- List state tracking maintains a stack of open lists for proper begin/end nesting.
- Track changes: accept insertions (include `w:ins` content), skip deletions (`w:del`), ignore formatting changes.
- Security: ZIP bomb detection, path traversal prevention, XML bomb protection via `lxml.etree.XMLParser(resolve_entities=False)`.
- Follow the KAOS Python QA process: `ruff format`, `ruff check --fix`, `ty check`, `pytest`.
- Tool classes follow kaos-core `KaosTool` ABC. Register with `register_office_tools(runtime)`.
- All MCP tools use shared `_OFFICE_ANNOTATIONS` (readOnly, idempotent, !destructive, !openWorld).
- Tool error messages must include recovery guidance (what went wrong + how to fix it + alternative tool).
- `search_document()` is imported from kaos-content (canonical, shared with kaos-pdf and kaos-web).
- **Never add AGPL/GPL dependencies.** This is a proprietary codebase.
