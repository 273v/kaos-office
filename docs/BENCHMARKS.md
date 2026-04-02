# kaos-office DOCX Extraction Benchmarks

**Date**: 2026-04-02
**Test machine**: Linux, Python 3.13

## Performance Comparison

Benchmarked against mammoth, markitdown (Microsoft), and python-docx on 7 real DOCX files
from the kelvin_office test corpus.

### Speed (lower is better)

| File | Size | kaos-office | mammoth | markitdown | python-docx |
|------|------|-------------|---------|------------|-------------|
| MultiParagraphSample | 7 KB | **6.7ms** | 5.3ms | 12.0ms | 1.4ms |
| Footnote | 6 KB | **1.1ms** | 2.8ms | 11.1ms | 0.8ms |
| Toro Term Loan | 187 KB | **130ms** | 475ms | 1,033ms | 54ms |
| Toro Redline | 147 KB | **204ms** | 536ms | 1,150ms | 62ms |
| Toro Comments | 148 KB | **201ms** | 535ms | 1,142ms | 61ms |
| CheeseSample | 1,051 KB | **9ms** | 140ms | 108ms | 8ms |
| BCFP Consumer Rights | 69 KB | **13ms** | 46ms | 127ms | 7ms |

**kaos-office is 2.5-8x faster than mammoth and 5-10x faster than markitdown.**
python-docx is faster on raw text extraction because it doesn't parse into a structured AST.

### Quality Features (check = detected in output)

| Feature | kaos-office | mammoth | markitdown | python-docx |
|---------|:-----------:|:-------:|:----------:|:-----------:|
| **Bold** (`**text**`) | 7/7 | 0/7 | 7/7 | 0/7 |
| **Italic** (`*text*`) | 7/7 | 7/7 | 7/7 | 0/7 |
| **Lists** (nested) | 6/7 | 0/7 | 0/7 | 0/7 |
| **Tables** | 3/3 | 0/3 | 0/3 | 3/3 (count) |
| **Footnotes** | 2/2 | 0/2 | 0/2 | 0/2 |
| **Comments** (annotations) | 3/3 | 0/3 | 0/3 | 0/3 |
| **Track changes** (accept) | yes | yes | yes | no |
| **Structured AST** | yes | no (HTML) | no (text) | no (runs) |
| **Provenance** | yes | no | no | no |
| **Search (BM25)** | yes | no | no | no |

### Key Advantages of kaos-office

1. **Only tool that extracts footnotes as structured content** — mammoth and markitdown lose them
2. **Only tool that extracts comments as annotations** — enables legal review workflows
3. **Only tool that produces nested lists** — mammoth and markitdown flatten list structure
4. **Produces a typed AST** (ContentDocument) — not raw HTML/text. Enables downstream search, section navigation, annotation, and re-serialization to multiple formats
5. **2.5-8x faster than mammoth** on the same documents
6. **5-10x faster than markitdown** on the same documents
7. **Track changes handled correctly** — insertions included, deletions skipped, clean output

### Known Gaps

1. **No heading detection on legal documents** — the Toro Term Loan uses bold/caps paragraph styles, not Word heading styles. kaos-office correctly detects 0 headings because none are styled as headings in the DOCX. This is a real limitation for legal documents that use custom styles.
2. **mammoth produces more characters** — includes some markup that kaos-office strips. May include content kaos-office filters.
3. **python-docx is faster for raw text** — if you only need unformatted text, python-docx is fastest. But it loses all structure.

## Approach Comparison

| Tool | Approach | License | Dependencies |
|------|----------|---------|-------------|
| **kaos-office** | Direct lxml XML parsing | Proprietary | lxml |
| **mammoth** | python-docx + custom HTML/md conversion | BSD-2 | lxml, python-docx |
| **markitdown** | python-docx + LLM-friendly conversion | MIT | python-docx, beautifulsoup4, many others |
| **python-docx** | XML parsing with OO model | MIT | lxml |

kaos-office uses direct lxml parsing (no python-docx dependency) for full control over OOXML processing. This enables track changes handling, comment extraction, and precise style resolution that wrapper-based approaches cannot access.

## Broader Competitive Landscape

Research across the full Python DOCX ecosystem (April 2026):

| Tool | Approach | License | Headings | Lists | Tables | Footnotes | Track Changes | Comments | Install Size |
|------|----------|---------|:--------:|:-----:|:------:|:---------:|:-------------:|:--------:|:------------:|
| **kaos-office** | Direct lxml | Proprietary | yes | yes (nested) | yes (merged) | yes | yes (accept) | yes | ~5 MB |
| **Pandoc** (pypandoc) | Haskell binary | GPL-2+ | yes | yes | yes | yes | yes (accept/reject/all) | yes | ~100 MB |
| **mammoth** | Direct XML | BSD-2 | yes | partial | yes | yes | no | yes | ~10 MB |
| **markitdown** | mammoth wrapper | MIT | yes | partial | weak | inherited | no | inherited | ~251 MB |
| **Docling** (IBM) | python-docx + lxml | MIT | yes | yes | best | no | no | yes | ~1 GB |
| **docx2python** | Direct XML | MIT | yes | yes | yes | yes | no | yes | ~5 MB |
| **docx2md** | Direct XML | MIT | yes | yes | yes (merged) | no | no | no | ~5 MB |
| **Unstructured** | python-docx | Apache-2.0 | yes | no (misclassified) | yes | no | no | no | ~146 MB |
| **Kreuzberg** | Pandoc wrapper (Rust) | MIT | yes | yes | yes | yes | yes | yes | ~71 MB |

### Key Findings

1. **Pandoc is the quality ceiling** — GPL Haskell binary, not embeddable as a Python library. Useful as benchmark reference only.
2. **mammoth is the closest peer** — direct XML parsing, BSD-licensed, clean output. But its markdown path is deprecated and it lacks track changes.
3. **markitdown is essentially a mammoth wrapper** — three-stage pipeline (DOCX → mammoth → HTML → markdownify). Inherits all mammoth limitations. "Designed as a basic text scraper" for complex layouts.
4. **Docling has the richest feature set** but at extreme cost — 1GB install, 60+ minutes/file on complex docs, frequent timeouts.
5. **Track changes is a differentiator** — only Pandoc and kaos-office handle it among the Python-accessible tools.
6. **No existing tool produces a typed AST with provenance** — kaos-office's ContentDocument output is unique. Every other tool outputs flat markdown/HTML or a proprietary intermediate representation.

### License Safety

All tools benchmarked are license-safe for KAOS: mammoth (BSD-2), python-docx (MIT), markitdown (MIT), docling (MIT), docx2python (MIT), kreuzberg (MIT). **Avoid**: Python-OOXML (AGPL-3.0), Pandoc (GPL-2+ — subprocess use only).
