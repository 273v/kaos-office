# kaos-office

Office document extraction for [KAOS](https://273ventures.com/) — DOCX, XLSX, PPTX to structured AST with provenance.

## Features

- **DOCX extraction**: Paragraphs, headings, tables, lists, images, footnotes, comments, track changes
- **Structured output**: kaos-content `ContentDocument` AST with provenance
- **Security**: ZIP bomb detection, path traversal prevention, XML bomb protection
- **MCP tools**: 5 tools for agent integration (parse, text, markdown, metadata, search)
- **CLI**: Human-readable and JSON output modes

## CLI Usage

```bash
# Extract as markdown (default)
kaos-office extract document.docx

# Extract as plain text
kaos-office extract document.docx --format text

# Extract as JSON AST
kaos-office extract document.docx --format json

# Search within a document
kaos-office search document.docx "force majeure" --top-k 5

# Get metadata
kaos-office metadata document.docx --json
```

## Python API

```python
from kaos_office import parse_docx

doc = parse_docx("document.docx")
print(doc.metadata.title)
print(len(doc.body))  # Number of top-level blocks

# Serialize to markdown
from kaos_content import serialize_markdown
print(serialize_markdown(doc))

# Search
from kaos_content.search import search_document
results = search_document(doc, "indemnification", top_k=5)
```

## MCP Integration

```python
from kaos_core import KaosRuntime
from kaos_office.tools import register_office_tools

runtime = KaosRuntime.default()
count = register_office_tools(runtime)  # Registers 5 tools
```
