# Tracked Changes / Redline Support — Design Options

**Created**: 2026-04-18
**Status**: Option 1 selected. Core read-path + typed API + transforms implemented; write-path pending on DOCX Phase 3.
**Scope**: kaos-content AST model + kaos-office DOCX reader/writer

## Implementation Status (2026-04-18, complete)

Option 1 (Span/Div containers) is fully shipped end-to-end. All four
use cases are working and validated on real legal fixtures.

**kaos-content**:
- `AnnotationType.TRACKED_CHANGE`
- `parse_docx(path, track_changes=True)` wraps revisions in Span/Div
- `serialize_text/markdown/html(doc, view="final"|"original"|"markup")`
- `kaos_content.revision`:
  - `Revision` / `Revisions` typed wrapper API
  - `accept / reject / accept_all / reject_all`
  - `accept_by_author / reject_by_author`
  - `at_time(doc, t)` — time machine
  - `make_inline_insertion / make_inline_deletion / make_block_insertion / make_block_deletion` — node constructors
  - `append_block_insertion / insert_block_after / delete_block_at` — doc-level authoring helpers

**kaos-office (DOCX writer)**:
- `word/comments.xml` emission from `AnnotationType.COMMENT` annotations
- `word/footnotes.xml` / `word/endnotes.xml` with required separator/continuation IDs
- `FootnoteRef` → `w:footnoteReference` / `w:endnoteReference` runs
- Proper `w:hyperlink` elements with relationship entries (dedup on URL)
- `rev-*` Span/Div → `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` with `w:delText` for deletions

**End-to-end validated**:
- Toro 2022 Term Loan - Redline v1 (12 revisions): round-trip with 0 metadata mismatches, 100% markup-view word overlap
- Toro 2022 Term Loan - Comments (5 comments): round-trip with all author/date/text preserved
- Footnote.docx: round-trip with separator+continuation markers emitted
- UC4 demo: agent authors redlines on clean CFPB summary → write → re-parse (2 revisions recovered) → accept_all → clean final doc

**Use case coverage**: UC1 ✓ UC2 ✓ UC3 ✓ UC4 ✓

---

## Problem

A redlined DOCX contains **two versions of every changed passage** — the original and the modified — plus metadata about who changed what and when. Our DOCX reader currently flattens tracked changes: insertions are silently accepted, deletions are silently dropped, all revision metadata (author, date, revision ID) is destroyed.

For a legal document platform, this is a critical gap. Redlines are the core of contract negotiation, deal review, and regulatory comment.

## What Users Need

### Three View Modes

1. **Original** — the document before any changes (accept all deletions, reject all insertions)
2. **Final** — the document after all changes accepted (what we produce today)
3. **Markup** — both versions visible, insertions underlined, deletions struck through, color-coded by author

### Core Operations

- Accept/reject individual changes
- Accept/reject all changes, or by author
- Navigate between changes ("next change" / "previous change")
- Change summaries ("12 insertions by Author A, 5 deletions by Author B")

### Two Distinct Problems

1. **Embedded tracked changes** — revision markup stored in the DOCX XML itself (`w:ins`, `w:del`). Word records these in real-time as the user edits.
2. **Computed comparison** — given two separate documents, compute a diff and produce a third document with tracked-change annotations. This is what Litera Compare, Draftable, and Word's Compare feature do.

This design covers problem 1 (faithful parsing of embedded tracked changes). Problem 2 (computed diff) is a separate, future project.

---

## Current State

### OOXML Tracked Changes Primitives

| Element | Purpose | Metadata |
|---------|---------|----------|
| `w:ins` | Inserted content (contains `w:r` with `w:t`) | `w:id`, `w:author`, `w:date` |
| `w:del` | Deleted content (contains `w:r` with `w:delText`) | `w:id`, `w:author`, `w:date` |
| `w:moveFrom` / `w:moveTo` | Moved content (paired by `w:name` on range markers) | `w:id`, `w:author`, `w:date` |
| `w:rPrChange` | Run property change (bold, italic, etc.) — stores "before" state | `w:id`, `w:author`, `w:date` |
| `w:pPrChange` | Paragraph property change (alignment, spacing) — stores "before" state | `w:id`, `w:author`, `w:date` |
| `w:tblPrChange` | Table property change | `w:id`, `w:author`, `w:date` |
| `w:sectPrChange` | Section property change | `w:id`, `w:author`, `w:date` |

### What kaos-office Does Today (reader.py)

```python
# Body-level: accept insertions, skip deletions
elif tag == W_INS:
    for child in el:
        _process_body_child(child, ctx)  # include content
elif tag == W_DEL:
    pass  # skip entirely — content lost forever
```

**Lost**: all revision metadata (author, date, ID), all deleted content, all formatting change tracking, move pair linkage, ability to render the original version.

### What kelvin-office Had

- `RevisionInfo` model: id, type, author, date, element_ids, original_content, new_content, accepted/rejected flags
- `RevisionManager`: 3 indexes (by ID, by element, by author), accept/reject operations, query by type/author
- `RevisionParser`: XML element → RevisionInfo
- `RevisionType` enum: 9 types (INSERTION, DELETION, MOVE_FROM, MOVE_TO, FORMAT_CHANGE, etc.)

Limitations: stored content as flat strings (not AST), monkey-patched `.revisions` onto objects, only handled 4 of ~28 OOXML revision types.

### What kaos-content Has

- `Annotation` system: standoff annotations with typed `AnnotationType`, multi-target `AnnotationTarget` (node_ref + char offsets), freeform `body: dict`
- `AnnotationType` enum: 15 values including COMMENT, REDACTION, AMENDMENT — but **no revision/tracked-change types**
- `Span` (inline) and `Div` (block): generic containers with `Attr(id, classes, kv)` — the Pandoc extension mechanism for domain semantics
- `REDACTION` annotation pattern: serializers check annotations before rendering, replacing content with `[REDACTED]`. This is the established pattern for annotations that modify rendering.

---

## Option 1: Span/Div Containers (use the designed extension point)

Both versions of content live in the AST as proper nodes, wrapped in `Span` (inline) or `Div` (block) with revision metadata in `Attr.kv`.

### AST Representation

```python
# "old text" was deleted, "new text" was inserted by Jane on 2026-04-15
Paragraph(children=(
    Span(
        attr=Attr(classes=("rev-del",), kv={
            "rev:id": "2", "rev:author": "Jane Smith", "rev:date": "2026-04-15T11:00:00Z"
        }),
        children=(Text(value="old text"),),
    ),
    Span(
        attr=Attr(classes=("rev-ins",), kv={
            "rev:id": "1", "rev:author": "Jane Smith", "rev:date": "2026-04-15T10:30:00Z"
        }),
        children=(Text(value="new text"),),
    ),
))
```

Block-level (e.g., an entire paragraph was inserted):
```python
Div(
    attr=Attr(classes=("rev-ins",), kv={"rev:id": "5", "rev:author": "John", ...}),
    children=(
        Paragraph(children=(Text(value="This entire paragraph was added."),)),
    ),
)
```

### Serializer Behavior

Generalizes the existing REDACTION pattern:

```python
def serialize_text(doc, *, view: str = "final") -> str:
    # Pre-compute rev-del and rev-ins node refs
    # view="final":   skip rev-del, render rev-ins normally (default, backward-compatible)
    # view="original": skip rev-ins, render rev-del normally
    # view="markup":  render both with formatting markers
```

### Metadata Queries via Annotations

Add `AnnotationType.TRACKED_CHANGE` for cross-cutting queries:

```python
# Reader emits one annotation per revision, pointing at the Span/Div
Annotation(
    type=AnnotationType.TRACKED_CHANGE,
    targets=(AnnotationTarget(node_ref="/body/0/children/1"),),
    body={"change_type": "insertion", "author": "Jane Smith", "date": "2026-04-15T10:30:00Z"},
)
```

This enables: "list all revisions", "filter by author", "count changes by type" — without walking the tree.

### Reader Changes

`parse_docx(path, *, track_changes: bool = False)`:
- `track_changes=False` (default): today's behavior, full backward compatibility
- `track_changes=True`: wrap `w:ins` content in `Span(classes=("rev-ins",))`, include `w:del` content in `Span(classes=("rev-del",))`, emit `TRACKED_CHANGE` annotations

### Writer Changes

`write_docx(doc, path)`:
- Detect `rev-ins` / `rev-del` Span/Div nodes
- Emit `w:ins` / `w:del` XML with metadata from `Attr.kv`
- Content without revision wrappers → clean output (today's behavior)

### DX Example

```python
from kaos_office.docx import parse_docx, write_docx
from kaos_content.serializers.text import serialize_text
from kaos_content.model.annotation import AnnotationType

# Parse with tracked changes preserved
doc = parse_docx("contract_redline.docx", track_changes=True)

# Three views from one parse
final = serialize_text(doc, view="final")       # accepted version
original = serialize_text(doc, view="original") # before changes
markup = serialize_markdown(doc, view="markup")  # both, formatted

# Query revisions
revisions = [a for a in doc.annotations if a.type == AnnotationType.TRACKED_CHANGE]
authors = {a.body["author"] for a in revisions}
insertions = [a for a in revisions if a.body["change_type"] == "insertion"]
deletions = [a for a in revisions if a.body["change_type"] == "deletion"]

# Accept all → produces a clean document (no rev-* spans)
from kaos_content.transforms.revisions import accept_all
clean_doc = accept_all(doc)
write_docx(clean_doc, "contract_clean.docx")
```

### Pros

- **Zero model changes** — Span/Div + Attr is the Pandoc extension mechanism; this is exactly what it's for
- **Both versions are proper AST** — walkable, searchable, serializable with formatting
- **Existing serializer pattern** — REDACTION generalization
- **Backward-compatible** — `track_changes=False` is the default; nothing changes for existing users
- **Writer round-trip** — Span/Div with `rev:*` kv maps directly to `w:ins`/`w:del` XML

### Cons

- **Stringly-typed** — `"rev-del"` not enforced by type checker; typo → silent bug
- **Every consumer must be aware** — tree walkers need to handle rev-* spans or they'll double-count content
- **kv values are all strings** — dates and IDs need parsing
- **Larger AST** — every tracked change adds wrapper nodes

---

## Option 2: New First-Class AST Types

Add `TrackedInsertion` and `TrackedDeletion` as new inline and block types with typed metadata fields.

### AST Representation

```python
# New in kaos_content/model/inlines.py
class TrackedInsertion(BaseInline):
    node_type: Literal["tracked_insertion"] = "tracked_insertion"
    revision_id: str
    author: str
    date: str | None = None
    children: tuple[Inline, ...]

class TrackedDeletion(BaseInline):
    node_type: Literal["tracked_deletion"] = "tracked_deletion"
    revision_id: str
    author: str
    date: str | None = None
    children: tuple[Inline, ...]

# Parallel in kaos_content/model/blocks.py
class TrackedInsertionBlock(BaseBlock):
    node_type: Literal["tracked_insertion_block"] = "tracked_insertion_block"
    revision_id: str
    author: str
    date: str | None = None
    children: tuple[Block, ...]

class TrackedDeletionBlock(BaseBlock):
    node_type: Literal["tracked_deletion_block"] = "tracked_deletion_block"
    revision_id: str
    author: str
    date: str | None = None
    children: tuple[Block, ...]
```

### DX Example

```python
from kaos_content.model.inlines import TrackedInsertion, TrackedDeletion

# Type-safe filtering
for node in walk(doc):
    if isinstance(node, TrackedDeletion):
        print(f"{node.author} deleted: {extract_text(node)}")
```

### Pros

- **Type-safe** — `isinstance()` checks, IDE autocomplete, type checker catches errors
- **Explicit in the grammar** — self-documenting, no magic strings
- **Typed metadata fields** — `author: str`, not `kv["rev:author"]`
- **Clean Union signatures** — `Inline = Text | Strong | ... | TrackedInsertion | TrackedDeletion`

### Cons

- **4 new types** — 2 inline + 2 block
- **Every serializer updated** — markdown, HTML, text serializers must handle new types
- **Every tree walker updated** — NodeIndex, visitor, extract_text
- **Breaks Pandoc philosophy** — Div/Span is the extension mechanism; new types are the nuclear option
- **Combinatorial explosion risk** — what about `TrackedMove`? `TrackedFormatChange`? Each becomes 2 more types

---

## Option 3: Annotations Only (standoff)

Keep the AST as "final" only. Store revision information entirely in standoff annotations. Deleted content lives in annotation bodies as serialized strings or embedded AST fragments.

### AST Representation

```python
# AST is the "final" version (identical to today)
# Annotations carry revision metadata:
Annotation(
    type=AnnotationType.TRACKED_CHANGE,
    targets=(AnnotationTarget(node_ref="/body/3/children/0", start_offset=0, end_offset=8),),
    body={
        "change_type": "insertion",
        "author": "Jane Smith",
        "date": "2026-04-15T10:30:00Z",
    },
)
Annotation(
    type=AnnotationType.TRACKED_CHANGE,
    targets=(AnnotationTarget(node_ref="/body/3", start_offset=0, end_offset=0),),
    body={
        "change_type": "deletion",
        "author": "Jane Smith",
        "date": "2026-04-15T11:00:00Z",
        "deleted_text": "old text that was removed",
        # OR: "deleted_ast": {...serialized AST fragment...}
    },
)
```

### Pros

- **Zero tree changes** — AST stays clean and small
- **Clean separation** — content vs. metadata fully decoupled
- **Simple for consumers that only want "final"** — they see today's AST, ignore annotations
- **Annotations already work** — existing infrastructure, serialization, querying

### Cons

- **Deleted content is not AST** — it's a string (or serialized fragment) in annotation body. Can't walk it, search it, or render it with formatting
- **"Original" view requires reconstruction** — must splice deleted text back into the tree from strings, losing AST structure
- **Markup view is hard** — interleaving strings from annotations with AST nodes at render time is fragile
- **Formatting of deleted text is lost** — bold/italic in deleted runs becomes plain text in `deleted_text`

---

## Comparison Matrix

| Criterion | Option 1 (Span/Div) | Option 2 (New Types) | Option 3 (Annotations) |
|-----------|---------------------|---------------------|----------------------|
| Model changes | None | 4 new types | None |
| Serializer changes | Add view mode param | Handle 4 new types | Add view mode (harder) |
| Deleted content is AST | Yes | Yes | No (strings) |
| Type safety | Stringly-typed | Full | N/A |
| Backward compatible | Yes (flag-gated) | Requires consumer updates | Yes |
| Markup view fidelity | Full (both versions in tree) | Full | Lossy (no formatting in deleted text) |
| Writer round-trip | Direct mapping | Direct mapping | Must reconstruct from strings |
| Follows Pandoc pattern | Yes (designed for this) | No (new grammar) | Partially |
| Complexity to implement | Medium | High | Low (but limited) |

## Prior Art

- **kelvin-office**: RevisionInfo + RevisionManager with flat string content (closest to Option 3)
- **Pandoc**: Uses Span/Div with classes for track-changes mode (`--track-changes=all` produces `Span` with `insertion`/`deletion` classes) — **this is Option 1**
- **Docling**: No tracked changes support
- **python-docx**: Exposes raw XML; no abstraction layer for revisions
- **Litera Compare**: Computed diff → OOXML tracked changes XML (problem 2, not covered here)

## Recommendation

Option 1 (Span/Div). Pandoc itself uses this exact approach for `--track-changes=all`. The extension mechanism was designed for this. Both versions are proper AST. Zero model changes. Backward-compatible via flag.

## Open Questions

1. Should `track_changes=True` be the default, or opt-in? (Recommend: opt-in, to avoid breaking existing consumers)
2. How to handle `w:rPrChange` (formatting changes)? The "before" and "after" formatting are both on the same run. Could emit nested Spans with `rev-fmt-before` / `rev-fmt-after` classes.
3. How to handle `w:moveFrom` / `w:moveTo` pairs? Need to link them by move name. Could use matching `rev:move-name` kv values.
4. Should we add a `kaos_content.transforms.revisions` module with `accept_all()`, `reject_all()`, `accept_by_author()` tree transforms?
5. Scope of computed comparison (problem 2): should this live in kaos-content (AST diff) or kaos-office (DOCX-specific)?
