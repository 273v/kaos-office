# Changelog

All notable changes to `kaos-office` are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0a7] — 2026-05-18

### Added

- **`kaos_office.docx.numbering` is now a package**, replacing the
  single-file module of the same name. New public surface:
  - `NumberingDefinitions` — the parsed schema. Resolves
    `(num_id, ilvl)` to a `LevelDefinition`, honoring
    `<w:lvlOverride>` / `<w:startOverride>`.
  - `NumberingState` — running counter machine. Emits the rendered
    visible label (`"11."`, `"(a)"`, `"11(a)(i)"`) for each numbered
    paragraph as the reader streams the document.
  - `parse_numbering_xml(numbering_xml_bytes) -> NumberingDefinitions`
    — replaces ad-hoc XML parsing scattered across the resolver.
  - `format_number(value, num_fmt)` and `format_lower_letter`,
    `format_lower_roman`, `format_upper_letter`, `format_upper_roman`,
    `format_ordinal`, `format_decimal_zero` — converters for visible
    numerals. Excel-style letter wraparound (`z → aa`) and Roman
    boundary cases are explicitly tested.
  - `is_ordered_format`, `BULLET_CHAR` re-exported for convenience.
- **`NumberingResolver`** preserves its 0.1.0a6 public API
  (`from_xml`, `is_ordered`, `get_format`, `has_numbering`) as a thin
  shim over `NumberingDefinitions` so existing callers continue to
  work unchanged. New code should prefer `NumberingDefinitions` +
  `NumberingState` because they expose the rendered label, not just a
  list-type boolean.
- **OOXML namespace constants** for the newly-parsed elements:
  `W_LVL_TEXT`, `W_START`, `W_LVL_RESTART`, `W_LVL_OVERRIDE`,
  `W_START_OVERRIDE`, `W_IS_LGL`, `W_SUFF`, `W_LVL_JC`.
- **`<w:pStyle>`-linked numbering** — paragraphs that inherit
  numbering through their paragraph style (no inline `<w:numPr>`)
  now pick up the rendered label. Firm templates routinely link
  "Heading 1" to a numbering definition so document authors get
  "Article 1.", "Article 2." automatically without manual numPr
  maintenance; before this change those headings parsed without
  labels. `NumberingDefinitions.resolve_pstyle(style_id)` does the
  lookup; `_handle_paragraph` consults it when no inline numPr is
  present.
- **International numbering formats** added to the converter table:
  `hebrew1`, `arabicAlpha`, `chineseCounting` (also
  `chineseCountingThousand`), `aiueo`, and `iroha`. Each is
  exercised by the formatter unit tests and the `format_number`
  dispatch test. Remaining Word formats (`hindi*`, `korean*`,
  `thai*`, `vietnamese*`, `ordinalText`, etc.) fall through to the
  decimal fallback with a structured warning, ready for incremental
  addition when fixtures surface.
- **`kaos-office-search` and `kaos-office-search-pptx` result dicts**
  now include `path: list[str]` per hit — the structural breadcrumb
  (root-first, INCLUDING the immediate section) for the matched
  paragraph or slide. Empty list is the explicit "no enclosing
  heading" contract; downstream agents MUST NOT invent section
  identifiers for hits with empty `path`. See
  `kaos-modules/docs/plans/persona-matrix-followups.md` §4.

### Changed

- **kaos-content floor raised to `>=0.1.0a12`** to pick up the
  structural-breadcrumb contract on `SearchResult.path` /
  `DocumentView.block_path()` AND the new `numbering_label` field on
  `Paragraph` / `Heading` / `ListItem`. The DOCX reader populates
  `numbering_label` with the rendered visible numeral from
  `word/numbering.xml` (e.g. `"Section 11."`, `"(a)"`, `"11(a)(i)"`),
  the writer round-trips them as plain-text run prefixes, and the
  serializers in kaos-content emit them verbatim.
- **DOCX reader populates `numbering_label`** on `Heading`,
  `Paragraph`, and `ListItem` AST nodes. `ParseContext` now carries
  a `NumberingState` alongside the existing `NumberingResolver`; for
  each paragraph carrying `<w:numPr>` with a non-zero `numId`, the
  rendered visible label is resolved in document order and attached
  to the AST node. Headings that inherit numbering through Word's
  auto-numbering machinery (the common legal pattern: `Section 11.
  GOVERNING LAW`) now carry the attorney-citable label that the
  previous reader silently dropped. List items receive the same
  treatment, replacing the silent fall-through to position-based
  recomputation downstream.
- **DOCX writer preserves `numbering_label` round-trip.** When a
  `Heading`, `Paragraph`, or `ListItem` carries a rendered numbering
  label (set by the reader from `numbering.xml`), the writer bakes
  the label as a plain-text run prefix on the paragraph so the
  round-tripped DOCX renders the same attorney-citable token. The
  pragmatic trade-off (versus reconstructing a full `<w:abstractNum>`
  per pattern) is that Word's edit-time renumbering no longer
  applies to imported sections — acceptable for review / redline
  workflows where the document is the source of truth, not a
  regenerable template. Round-trip tests under
  `tests/unit/test_docx_numbering_roundtrip.py` exercise all three
  fixture shapes (decimal / NDA / legal-outline).

Stages 1c (search-path) + 2-7 (docx-numbering-resolution) of the
respective plans under `kaos-modules/docs/plans/`.

## [0.1.0a6] — 2026-05-17

### Changed

- **kaos-core floor raised to `>=0.1.0a10`** to pick up the URI
  contract redesign (bare names route through
  `context.default_vfs_namespace`; `file://` and `vfs://` schemes).
  See `kaos-modules/docs/plans/uri-contract-redesign.md`. The 15
  file-input tools route through `resolve_input_path` as
  pass-throughs; no synthetic bare names internally.
- **Tests migrated to the new URI contract.** Test fixtures and
  nonexistent-path literals in `tests/unit/test_tools.py`,
  `tests/unit/test_pptx_tools.py`,
  `tests/integration/test_mcp_office_pipeline.py`, and
  `tests/integration/test_mcp_xlsx_pipeline.py` now supply
  `file:///abs/path` URIs (via `Path.as_uri()`) instead of bare
  absolute strings, mirroring how MCP clients will pass
  trusted-source filesystem paths under the new contract. No
  production code change in `kaos-office` itself — `tools.py` was
  already a pure pass-through to `resolve_input_path`.

## [0.1.0a5] — 2026-05-17

### Changed

- **All 15 file-input MCP tools now route their `path` parameter
  through `kaos_core.path_resolver.resolve_input_path()`** via a new
  internal adapter `kaos_office._path_resolver.resolve_office_input()`.
  Previously every tool ran `Path(path_str).exists()` against the
  process CWD, which made files uploaded into `KaosRuntime.vfs` by a
  host UI (e.g. `kaos-ui`'s single-user-chat SPA) invisible — agents
  saw an unbroken sequence of "File not found" errors and were at risk
  of hallucinating answers from zero successful reads. Affected tools:
  `kaos-office-parse-docx`, `kaos-office-get-text`,
  `kaos-office-get-markdown`, `kaos-office-metadata`,
  `kaos-office-search` (DOCX); `kaos-office-parse-pptx`,
  `kaos-office-list-slides`, `kaos-office-get-slide`,
  `kaos-office-search-pptx`, `kaos-office-get-slide-notes` (PPTX);
  `kaos-office-parse-xlsx`, `kaos-office-list-sheets-xlsx`,
  `kaos-office-get-sheet-xlsx`, `kaos-office-xlsx-metadata` (XLSX);
  and `kaos-office-write-pptx`'s optional `template_path`. Each tool's
  `path` schema description now advertises that
  `kaos://artifacts/<id>` URIs and session-VFS paths are accepted in
  addition to filesystem paths. Parse-* tools that materialise a new
  derived artifact now thread `source_artifact_id` / `source_body_uri`
  into their `structuredContent` when the input came from the
  artifact store, so the SPA's ArtifactCard renders the original
  upload's id rather than a fresh derived one. Stage 1 of
  `kaos-modules/docs/plans/vfs-blind-tools-audit-and-fix-plan.md` —
  upstream fix for the production hallucination incident where every
  SPA-uploaded `.docx` was invisible to the entire office tool set.
  Behavior for absolute filesystem paths is unchanged (the resolver's
  filesystem branch is a passthrough). New unit tests:
  `tests/unit/test_vfs_path_resolution.py` (15 cases — one per
  affected read tool plus the WritePptx template path).
- **Pinned `kaos-core>=0.1.0a9,<0.2`** (was `>=0.1.0a1`). The 0.1.0a9
  release ships `kaos_core.path_resolver`, which the new adapter
  depends on.

## [0.1.0a4] — 2026-05-15

### Added — documents + authoring registration entry points (PRD PR 1)

- **`register_office_documents_tools(runtime)`** — registers the 14
  read-only Office tools (DOCX / PPTX / XLSX parsers, listers,
  getters, metadata inspectors, BM25 searchers). Pins the
  SessionToolSet `documents` group entry point.
- **`register_office_authoring_tools(runtime)`** — registers the 3
  Office writers (`kaos-office-write-docx` /
  `kaos-office-write-pptx` / `kaos-office-write-xlsx`). Pins the
  SessionToolSet `authoring` group entry point: denied by default
  at the ceiling and opted into per-session for drafting workflows.
- **`register_office_tools(runtime)`** is now a backward-compatible
  union — every existing caller continues to see the same 17 tools
  with the same names and schemas.

Motivated by `kaos-modules/docs/internal/dynamic-tool-planning-prd.md`
§4 ("PR 1 — catalog expansion"). Purely additive: no tool name,
schema, or behavior changes.

## [0.1.0a3] — 2026-05-15

### Fixed

- **`kaos-office-parse-xlsx`'s `sheets` parameter now declares its
  element type.** Previously the schema was `type=array` with no
  `items`, which OpenAI's strict JSON Schema validator rejected
  with HTTP 400 `invalid_function_parameters`. The whole tool
  catalog for the turn was lost. Now `items: {type: "string"}` so
  the LLM gets a precise contract for sheet names. kaos-core
  0.1.0a7's defensive `items: {}` floor is belt + suspenders.

### Added

- **`[mcp]` extra restored.** kaos-office's pyproject originally
  declared the `[mcp]` extra absent at 0.1.0a1 because `kaos-mcp`
  wasn't on PyPI yet and `uv lock` refuses to resolve unresolvable
  declared extras (F009 #4). `kaos-mcp` shipped (now at 0.1.0a3), so
  the extra is back: `pip install kaos-office[mcp]` (or
  `uv add kaos-office[mcp]`) now pulls in
  `kaos-mcp>=0.1.0a3,<0.2` for the FastMCP-backed
  `kaos-office-serve` runner and the MCP integration tests.

### Fixed

- **CI: nightly integration tests now install the `mcp` extra.** The
  scheduled-only `integration tests` job in `security.yml` failed
  collection on `test_mcp_office_pipeline.py` /
  `test_mcp_xlsx_pipeline.py` with
  `ModuleNotFoundError: No module named 'kaos_mcp'` because the
  job ran `uv sync --group dev` without the (then-missing) extra.
  Now syncs with `--group dev --extra mcp` so the MCP-bridge tests
  can import `kaos_mcp`. Push/PR runs are unchanged (the integration
  job is `if: github.event_name == 'schedule' || workflow_dispatch`).


### Fixed

- **Tests: Windows-x64 leg failed on hardcoded POSIX tempfile path.**
  ``tests/unit/test_reader.py::_parse_from_body`` wrote its synthetic
  DOCX to ``Path("/tmp/test_reader.docx")``. On Windows this resolved
  to ``\\tmp\\test_reader.docx`` (a non-existent drive-root path) so
  every DOCX reader test failed with ``FileNotFoundError``. Switched
  to ``tempfile.mkstemp(suffix=".docx")`` so the path lives under
  ``%TEMP%`` on Windows and ``/tmp`` on POSIX. No production code
  change. Files: ``tests/unit/test_reader.py``.
### Security

- **vulture (dead-code scan) now runs in pre-commit + CI alongside
  the existing bandit job.** New `vulture` hook in
  ``.pre-commit-config.yaml`` mirrored by a new ``vulture (dead-code
  scan)`` job in ``security.yml``. `--min-confidence 100` with the
  shared `--ignore-names` list for names vulture can't infer from
  the import graph (framework callbacks, OAuth/OIDC field names,
  signal handlers, MCP `_meta` keys). Also lands the existing
  bandit hook in pre-commit (it was only in CI before). Both pass
  clean. Mirrors the rollout from kaos-core.
### Changed

- **uv.lock is now tracked in git.** Previously gitignored at v0.1.0a1
  because the ``[mcp]`` optional extra (and the ``kaos-mcp`` dev
  dependency) referenced a sibling not yet on PyPI; ``uv lock``
  couldn't resolve them. ``kaos-mcp`` shipped (0.1.0a2), so the
  original gating reason no longer applies. Tracking the lockfile
  gives reproducible local dev environments, lets Dependabot surface
  sibling-version bumps as PRs, and makes the supply-chain pin set
  publicly auditable. Mirrors the org-wide convention being adopted
  across all 16 kaos-* repos.

## [0.1.0a2] — 2026-05-08

CI supply-chain hardening (audit-02 F7) and SECURITY.md polish (audit-02
F8). No source code or public API changes.

### Security

- **F7: CI supply-chain hardening.** `.github/workflows/security.yml`
  pins the gitleaks Docker image to `v8.21.2` (no longer tracking
  `:latest`), adds a Bandit static-analysis job (medium severity /
  medium confidence; `B101,B404,B603,B607` skipped), and runs the
  integration suite on `schedule` and `workflow_dispatch` so
  cross-package regressions surface against `main` even though the
  unit gate stays the PR fast path. SHA-pinning of GitHub Actions
  themselves remains a follow-up; the existing
  `.github/dependabot.yml` `github-actions` ecosystem PRs continue to
  keep tag-pinned actions current.

### Changed

- **F8: `SECURITY.md` scope polished.** Added a one-paragraph preamble
  describing what `kaos-office` does, listed the actual Tool boundary
  (`ParseDocxTool`, `WriteDocxTool`, `ParseXlsxTool`, …), called out
  that MCP transport security lives in `kaos-mcp` rather than here.
  Existing OPC / OOXML / writer-path scope kept verbatim — it was
  already accurate for this module.

## [0.1.0a1] — 2026-05-07

First public alpha. Office document extraction (DOCX, PPTX, XLSX) into
the kaos-content typed AST, with native lxml round-trip writers and 17
MCP tools. Closes every finding in `docs/audit-01/kaos-office.md`
(KO-001..KO-008).

### Added

- **`LICENSE`, `NOTICE`, `CHANGELOG.md`** seeded for the public
  release. License flips from `LicenseRef-Proprietary` to Apache-2.0
  via PEP 639 (`license = "Apache-2.0"`, `license-files =
  ["LICENSE", "NOTICE"]`). PEP-639-superseded `License ::` classifier
  removed.
- **Aggregate `[xlsx]` extra** — `python-calamine` + `openpyxl`. The
  default native lxml reader needs no extras; this aggregate is for
  callers who want the calamine fast-path *and* openpyxl-backed
  formula extraction in one install. Closes audit-01 KO-004.
- **`kaos_office.pptx.OverflowMode` re-export** plus `parse_pptx`,
  `get_slide_count`, `get_slide_text`, `get_slide_notes`, and
  `list_slides` at the `kaos_office.pptx` package level.
  `kaos_office.xlsx` now also re-exports `parse_xlsx` + `list_sheets`
  next to the writer entry points. Closes audit-01 KO-007.
- **External fixtures opt-in mechanism** — `tests/conftest.py`
  exposes `external_fixture()` + `skip_without_external_fixture()`
  driven by the `KAOS_OFFICE_EXTERNAL_FIXTURES_DIR` env var, for
  real-world documents too large to vendor in-repo
  (e.g. multi-MB legal decks). Closes audit-01 KO-006.
- **`IEO2021_ChartLibrary_Industrial.pptx`** vendored under
  `tests/fixtures/pptx/` (447 KB; previously read from an absolute
  `kelvin-modules` path on the maintainer's workstation). Same
  reproducibility fix.

### Changed

- **`parse_docx()` resolves the input path before `Path.as_uri()`.**
  Pre-fix, any relative DOCX path crashed the reader with
  `ValueError: relative path can't be expressed as a file URI` —
  ordinary CLI / MCP usage that handed a filename in the current
  working directory was broken before the file was even opened.
  Mirrors the existing PPTX/XLSX behavior. Closes audit-01 KO-001.
  Regression coverage: `tests/unit/test_reader.py::TestRelativePaths`.
- **`kaos_office.pptx` writer entry points are typed lazy wrappers.**
  The previous `try/except ImportError: write_pptx = None` pattern
  failed `ty check` (`invalid-assignment` — `None` is not assignable
  to a callable returning `Path` / `bytes`) and exposed `None`
  callables when `python-pptx` was absent. Replaced with thin
  wrappers that defer the writer import until call time and raise
  `ImportError` with the `[pptx]` install hint instead. Public
  signatures and behavior are unchanged when `python-pptx` is
  installed. Closes audit-01 KO-002. Regression coverage:
  `tests/unit/test_pptx_writer.py::TestPptxLazyWrappers` (including
  a monkeypatched missing-dep path).
- **DOCX vMerge continuation cells use a typed marker class.**
  `_handle_table` previously emitted `Cell(content=(), row_span=0,
  col_span=col_span)` for `<w:vMerge/>` continuation cells. The
  `row_span=0` sentinel was rejected after kaos-content 0.1.0a1
  tightened `Cell` validation to require span ≥ 1, so real-world
  documents with vertical merges (e.g. the audit's Toro 2022 Term
  Loan and MCS Redline fixtures) failed to parse. Continuation cells
  now carry `row_span=1` and `attr=Attr(classes=("vmerge-continue",))`,
  preserving the grid geometry while letting downstream consumers
  detect the merge without a magic-value sentinel. Closes audit-01
  KO-003.
- **XLSX tool errors follow the KAOS three-part contract.** Every
  `ToolResult.create_error()` site in the four XLSX tools
  (`ParseXlsxTool`, `ListSheetsXlsxTool`, `GetSheetXlsxTool`,
  `XlsxMetadataTool`) now states what went wrong, how to fix it, and
  which sibling tool to try next. Shared `_XLSX_IMPORT_ERROR` and
  `_xlsx_file_not_found()` helpers keep the messages consistent and
  unit-testable. Errors no longer falsely claim that "XLSX requires
  python-calamine" — the default native lxml reader has no extras.
  Closes audit-01 KO-008.
- **`xlsxwriter`** is no longer a published `[xlsx-write]` extra; it
  remains in `[dependency-groups].dev` where it belongs (it is the
  cross-validator the test suite uses to verify the native lxml
  writer's output, not a production code path). The reference
  implementation in `kaos_office.xlsx._xlsxwriter_reference` (private
  per the leading underscore) updates its docstring + ImportError to
  point at the dev group. Closes audit-01 KO-005.
- **`tests/conftest.py`** rewritten to use in-repo vendored fixture
  roots as primary; the legacy `KELVIN_FIXTURES` /
  `KELVIN_PPTX_FIXTURES` symbols are preserved as aliases pointing at
  the local copies so existing test imports keep working. The two
  benchmark scripts under `tests/` are env-overridable
  (`KAOS_OFFICE_BENCHMARK_{DOCX,PPTX}_DIR`) and default to the
  vendored corpora. No `/home/<user>/...` paths remain in tracked
  files. Closes audit-01 KO-006.

### Project metadata

- `[project.urls]` adds `Issues` + `Changelog`; `Repository` switches
  to `https://github.com/273v/kaos-office`.
- `keywords` populated; `Operating System :: OS Independent`
  classifier added.
- `[tool.hatch.build.targets.sdist]` includes `LICENSE`, `NOTICE`,
  and `CHANGELOG.md`.
- `[mcp]` extra (which depended on `kaos-mcp`, not yet on PyPI) is
  stripped from the per-module repo's `pyproject.toml` per
  `docs/oss/checklists/per-package-release.md` Phase B5; it will be
  re-added in `0.1.0a2` once `kaos-mcp` ships.

[Unreleased]: https://github.com/273v/kaos-office/compare/v0.1.0a1...HEAD
[0.1.0a1]: https://github.com/273v/kaos-office/releases/tag/v0.1.0a1
