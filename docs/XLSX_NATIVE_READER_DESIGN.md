# XLSX Native Reader Design: lxml + OPC

**Date**: 2026-04-03
**Status**: Design — ready for implementation
**Depends on**: OPC infrastructure (complete), namespace.py (needs SpreadsheetML additions)

## Problem

The current XLSX reader uses python-calamine (Rust). It works and is fast, but:
1. **Bypasses our OPC security layer** — no ZIP bomb detection, no path traversal prevention
2. **Black-box type inference** — we can't control edge cases
3. **External Rust dependency** where we already have the infrastructure (lxml + OPC)
4. **Inconsistent** with DOCX and PPTX which use our own lxml parsers

DOCX and PPTX readers are 5-10x faster than alternatives because they use direct
lxml XML parsing through our OPC layer. XLSX should follow the same pattern.

## Decision

- **Default**: Native lxml reader using OPC package (no external dependencies)
- **Optional**: python-calamine reader as `[xlsx-calamine]` extra for users who
  want maximum speed on very large files (100K+ rows)
- Both produce identical `TabularDocument` output

## Architecture

```
kaos_office/xlsx/
├── __init__.py
├── reader.py              # parse_xlsx() entry point — dispatches to native or calamine
├── native.py              # Native lxml + OPC reader (NEW — default)
├── calamine_reader.py     # Calamine reader (RENAMED from current reader.py)
├── shared_strings.py      # Shared string table parser (NEW)
├── styles.py              # Style/date format detection (NEW)
└── cell_ref.py            # Cell reference parsing utilities (NEW)
```

### Entry point: `reader.py`

```python
def parse_xlsx(
    path: str | Path,
    *,
    sheets: list[str] | None = None,
    max_rows: int | None = None,
    header_row: int = 0,
    include_formulas: bool = False,
    engine: Literal["native", "calamine"] = "native",
) -> TabularDocument:
    """Parse XLSX → TabularDocument.

    Args:
        engine: "native" (default, lxml+OPC) or "calamine" (Rust, faster for 100K+ rows).
            Calamine requires the [xlsx-calamine] extra.
    """
    if engine == "calamine":
        from kaos_office.xlsx.calamine_reader import parse_xlsx_calamine
        return parse_xlsx_calamine(path, sheets=sheets, max_rows=max_rows,
                                   header_row=header_row, include_formulas=include_formulas)
    from kaos_office.xlsx.native import parse_xlsx_native
    return parse_xlsx_native(path, sheets=sheets, max_rows=max_rows,
                             header_row=header_row, include_formulas=include_formulas)
```

### Native reader: `native.py`

Two-pass architecture (same as DOCX reader):

**Pass 1: Metadata load**
1. Open OPC package (security checks: ZIP bomb, path traversal, XML bomb)
2. Parse `xl/workbook.xml` → sheet names, rIds, visibility
3. Resolve sheet paths via `xl/_rels/workbook.xml.rels`
4. Parse `xl/sharedStrings.xml` → build string lookup table
5. Parse `xl/styles.xml` → build numFmtId → is_date lookup table

**Pass 2: Sheet data walk**
For each selected sheet:
1. Parse `xl/worksheets/sheetN.xml`
2. Walk `<sheetData>` → `<row>` → `<c>` elements
3. For each cell:
   - Check `t` attribute (s=shared string, b=boolean, e=error, n/absent=number)
   - For shared strings: look up index in string table
   - For numbers: check style index → numFmtId → is_date?
   - For formulas: extract `<f>` text if `include_formulas=True`
4. Extract merged cell ranges from `<mergeCells>`
5. Build `Table` with `infer_column_type()` from kaos-content

### Shared strings: `shared_strings.py`

```python
class SharedStringTable:
    """Parsed shared string table from xl/sharedStrings.xml."""

    def __init__(self, xml: etree._Element) -> None:
        self._strings: list[str] = []
        for si in xml.iterchildren(SML_SI):
            # Plain string: <si><t>text</t></si>
            t_elem = si.find(SML_T)
            if t_elem is not None and t_elem.text:
                self._strings.append(t_elem.text)
                continue
            # Rich text: <si><r><t>part1</t></r><r><t>part2</t></r></si>
            parts = [t.text or "" for r in si.iterchildren(SML_R) for t in r.iterchildren(SML_T)]
            self._strings.append("".join(parts))

    def get(self, index: int) -> str:
        return self._strings[index]

    def __len__(self) -> int:
        return len(self._strings)
```

### Styles / date detection: `styles.py`

```python
# Built-in date numFmtIds
_DATE_FMT_IDS = frozenset({14, 15, 16, 17, 22})
_TIME_FMT_IDS = frozenset({18, 19, 20, 21, 45, 46, 47})
_DATETIME_FMT_IDS = _DATE_FMT_IDS | _TIME_FMT_IDS

# Heuristic for custom formats: contains date/time tokens
_DATE_TOKEN_RE = re.compile(r'[yYdDhHsS]|(?<![hH])m{1,5}(?![sS])')

class StyleTable:
    """Parsed style info from xl/styles.xml for date detection."""

    def __init__(self, xml: etree._Element) -> None:
        # Parse custom numFmts (id >= 164)
        self._custom_date_fmts: set[int] = set()
        # Parse cellXfs to map style index → numFmtId
        self._xf_to_numfmt: list[int] = []

    def is_date(self, style_index: int) -> bool:
        """Check if a style index represents a date format."""
        numfmt_id = self._xf_to_numfmt[style_index]
        if numfmt_id in _DATETIME_FMT_IDS:
            return True
        return numfmt_id in self._custom_date_fmts

    def is_time_only(self, style_index: int) -> bool:
        """Check if this is a time-only format (no date component)."""
        numfmt_id = self._xf_to_numfmt[style_index]
        return numfmt_id in _TIME_FMT_IDS
```

### Cell reference parsing: `cell_ref.py`

```python
_CELL_RE = re.compile(r'^([A-Z]+)(\d+)$')

def parse_cell_ref(ref: str) -> tuple[int, int]:
    """Parse 'A1' → (row=0, col=0), 'B2' → (row=1, col=1). 0-based."""

def col_letter_to_index(letters: str) -> int:
    """'A' → 0, 'Z' → 25, 'AA' → 26, 'XFD' → 16383."""
```

### Excel date serial conversion

```python
_EPOCH_1900 = datetime.date(1899, 12, 30)  # Excel epoch (off by one due to Lotus bug)
_EPOCH_1904 = datetime.date(1904, 1, 1)

def serial_to_date(serial: float, *, date1904: bool = False) -> datetime.date | datetime.datetime:
    """Convert Excel serial number to Python date/datetime.

    If fractional part is non-zero, returns datetime. Otherwise returns date.
    Handles the Excel 1900 leap year bug (serial 60 = Feb 29, 1900 doesn't exist).
    """
```

## Namespace Additions

Add to `kaos_office/ooxml/namespace.py`:

```python
# SpreadsheetML
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Pre-computed Clark notation tags
SML_WORKBOOK = qn(SML, "workbook")
SML_SHEETS = qn(SML, "sheets")
SML_SHEET = qn(SML, "sheet")
SML_WORKSHEET = qn(SML, "worksheet")
SML_SHEET_DATA = qn(SML, "sheetData")
SML_ROW = qn(SML, "row")
SML_CELL = qn(SML, "c")
SML_VALUE = qn(SML, "v")
SML_FORMULA = qn(SML, "f")
SML_SST = qn(SML, "sst")
SML_SI = qn(SML, "si")
SML_T = qn(SML, "t")
SML_R = qn(SML, "r")
SML_MERGE_CELLS = qn(SML, "mergeCells")
SML_MERGE_CELL = qn(SML, "mergeCell")
SML_NUM_FMTS = qn(SML, "numFmts")
SML_NUM_FMT = qn(SML, "numFmt")
SML_CELL_XFS = qn(SML, "cellXfs")
SML_XF = qn(SML, "xf")
SML_COLS = qn(SML, "cols")
SML_COL = qn(SML, "col")
SML_DIMENSION = qn(SML, "dimension")
SML_INLINE_STR = qn(SML, "is")
SML_WORKBOOK_PR = qn(SML, "workbookPr")
```

## pyproject.toml Changes

```toml
[project.optional-dependencies]
xlsx = []                    # Native reader needs no extra deps (lxml is already a dep)
xlsx-calamine = [
    "python-calamine>=0.6",
]
xlsx-formulas = [
    "openpyxl>=3.1",         # Only for formula text extraction
]
```

The native reader uses only lxml (already a core dependency) and the OPC layer (already
built). **Zero new dependencies for the default path.**

## Implementation Order

### Step 1: Add SpreadsheetML namespace constants (~20 lines)
File: `ooxml/namespace.py`

### Step 2: Implement cell_ref.py (~30 lines)
Cell reference parsing, column letter conversion.

### Step 3: Implement shared_strings.py (~40 lines)
Parse `xl/sharedStrings.xml` into indexed lookup table.
Handle both plain `<t>` and rich text `<r><t>` forms.

### Step 4: Implement styles.py (~80 lines)
Parse `xl/styles.xml` for numFmtId → is_date mapping.
Built-in date format IDs (14-22, 45-47).
Custom format detection via regex heuristic.
Excel serial number → Python date/datetime conversion.

### Step 5: Implement native.py (~250 lines)
The core reader:
- Open OPCPackage → parse workbook.xml → resolve sheet paths
- Load SharedStringTable and StyleTable
- Walk sheet XML: rows → cells → typed values
- Build Table with column type inference
- Extract merged cell ranges and formulas

### Step 6: Rename current reader.py → calamine_reader.py
Move calamine-specific code. Update imports.

### Step 7: Update reader.py as dispatcher
`parse_xlsx(engine="native"|"calamine")` entry point.
Default to "native". Calamine requires `[xlsx-calamine]` extra.
`list_sheets()` uses native reader (no calamine needed for metadata).

### Step 8: Update pyproject.toml
Change `xlsx` extra to empty (native needs nothing).
Add `xlsx-calamine` extra.

### Step 9: Tests
- Run existing 26 XLSX tests against both engines
- Parametrize: `@pytest.mark.parametrize("engine", ["native", "calamine"])`
- Add native-specific tests: date detection, shared strings, formulas, merged cells
- Verify identical output from both engines on all 5 XLSX fixtures

### Step 10: QA pass
ruff format, ruff check, ty check, pytest

## Estimated Size

| File | Lines |
|------|-------|
| cell_ref.py | ~30 |
| shared_strings.py | ~40 |
| styles.py | ~80 |
| native.py | ~250 |
| reader.py (dispatcher) | ~40 |
| namespace.py additions | ~25 |
| **Total new** | **~465** |
| calamine_reader.py (renamed) | ~280 (existing, moved) |

## Key Design Decisions

**1. Two-pass architecture** — same as DOCX. Load metadata first (shared strings,
styles), then walk data. Avoids re-parsing XML.

**2. No iterparse for MVP** — parse full sheet DOM. iterparse optimization can come
later if we need to handle million-row sheets. Our XLSX fixtures max out at 14K rows;
full DOM is fine for that.

**3. Date detection via styles.xml** — the only reliable way. Numbers are dates if and
only if their numFmtId is a date format. Built-in IDs 14-22/45-47 plus custom format
regex heuristic.

**4. Shared strings are mandatory** — virtually all XLSX files use them. Parse once,
lookup by index.

**5. Formula extraction is opt-in** — `include_formulas=True` stores formula text in
`table.metadata["formulas"]`. The `<v>` cached value is always used for the cell value.

**6. Both engines produce identical TabularDocument** — same types, same column names,
same row values. Tests will verify this with parametrized fixtures.

## Verification

```bash
cd kaos-office
uv run ruff format kaos_office/ tests/
uv run ruff check --fix kaos_office/ tests/
uv run ty check kaos_office/ tests/
uv run pytest tests/unit/test_xlsx_reader.py -v   # All tests pass with both engines
```

Then verify cross-engine consistency:
```python
from kaos_office.xlsx.reader import parse_xlsx

native = parse_xlsx("file.xlsx", engine="native")
calamine = parse_xlsx("file.xlsx", engine="calamine")
assert native.tables[0].row_count == calamine.tables[0].row_count
assert native.tables[0].column_names() == calamine.tables[0].column_names()
```
