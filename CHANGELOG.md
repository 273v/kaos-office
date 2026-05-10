# Changelog

All notable changes to `kaos-office` are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
