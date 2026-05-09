# Agent Instructions

## Scope

This file is the canonical coding-agent guidance for this repository. It applies to all repository work unless a maintainer gives more specific instructions in an issue, pull request, or task prompt.

Keep changes focused, public-repo-safe, and consistent with [CONTRIBUTING.md](CONTRIBUTING.md) and the standards under [docs/standards/](docs/standards/):

- [Python design and architecture](docs/standards/python-design-and-architecture.md)
- [Code quality standards](docs/standards/code-quality-standards.md)
- [Engineering process](docs/standards/engineering-process.md)
- [Tests, fixtures, and CI](docs/standards/tests-fixtures-ci.md)

## Project Identity

`kaos-office` is the distribution name. It publishes the `kaos_office` import package and the `kaos-office` and `kaos-office-serve` console scripts.

The package reads and writes Microsoft Office documents:

- DOCX and PPTX map to `kaos-content` `ContentDocument` AST models with provenance.
- XLSX maps to `kaos-content` `TabularDocument` AST models with provenance and metadata.
- CLI, MCP tool, JSON, schema, public API, and environment-variable contracts are public behavior once released.

## Setup

Use Python 3.13 or newer and `uv`:

```bash
uv sync --group dev
```

Install pre-commit hooks when working interactively:

```bash
uvx pre-commit install
```

Use `ruff` for formatting and linting, `ty` for type checking, and `pytest` for tests. Do not substitute mypy for `ty`.

## Local Checks

Run the checks that match the change. For normal code changes:

```bash
uv run ruff format --check kaos_office tests
uv run ruff check kaos_office tests
uv run ty check kaos_office tests
uv run pytest tests/unit -q --no-cov
```

For broader verification, use the tiered test guidance in [docs/standards/tests-fixtures-ci.md](docs/standards/tests-fixtures-ci.md):

```bash
uv run pytest -m "not live and not network and not slow" --no-cov
```

When packaging, release behavior, metadata, or README rendering changes, also run:

```bash
uv build
uvx --from twine twine check --strict dist/*
```

For docs-only changes, at minimum run `git diff --check` and a practical Markdown/link sanity check.

## Architecture Rules

Follow [docs/standards/python-design-and-architecture.md](docs/standards/python-design-and-architecture.md). In particular:

- Keep the base dependency set small and put optional integrations behind extras.
- Keep optional dependencies behind lazy imports.
- Keep import-time work minimal.
- Keep public API re-exports explicit in `kaos_office.__all__`.
- Preserve `py.typed` and typed public boundaries.
- Use package-specific exceptions for user-facing failures.
- Keep CLI and MCP errors concise, actionable, and stable.

Office-document implementation rules:

- DOCX, PPTX, and XLSX readers and writers should map Office structures to and from the `kaos-content` AST without discarding provenance that downstream tools need for citation, review, or search.
- Preserve package, OPC, relationship, and XML safety. Treat Office files as untrusted archives and XML inputs.
- Avoid unsafe archive and path behavior: no path traversal, arbitrary file reads, unsafe symlink following, unbounded extraction, or silent overwrite behavior.
- Bound parsing work for untrusted inputs where practical, including archive size, XML complexity, row counts, recursion, and memory-heavy paths.
- Keep read/write paths symmetric where feasible so round-trip behavior remains understandable and testable.
- Do not introduce GPL, AGPL, unknown-license, non-commercial, or no-derivatives dependencies.
- Keep CLI, MCP tool, JSON, schema, and environment-variable contracts stable unless the change is intentional, documented, tested, and reflected in the changelog when user-visible.

Source layout:

- `kaos_office/docx/`: DOCX reader, writer, metadata, styles, and numbering support.
- `kaos_office/pptx/`: PPTX reader, writer, and SmartArt handling.
- `kaos_office/xlsx/`: XLSX readers, writer, cell references, shared strings, styles, and optional engine adapters.
- `kaos_office/opc/` and `kaos_office/ooxml/`: Open Packaging Convention, relationship, content type, namespace, and safety helpers.
- `kaos_office/cli.py`, `kaos_office/serve.py`, and `kaos_office/tools.py`: CLI and MCP-facing surfaces.

## Testing

Add or update tests when behavior changes. Bug fixes need regression tests through the real public entry point where practical.

Use realistic redistributable fixtures and golden files for Office behavior. Fixtures must be free of secrets, privileged content, customer data, and incompatible licensing, and their provenance should be documented as described in [docs/standards/tests-fixtures-ci.md](docs/standards/tests-fixtures-ci.md).

Security-sensitive parser, archive, XML, path, and writer behavior should test both accepted and rejected cases. Mocked-only tests are not enough for those paths.

## Security

Never commit secrets, credentials, private keys, `.env` files, customer documents, or private data. Do not include internal paths, credentials, provider payloads, or stack traces in user-facing errors.

Handle suspected vulnerabilities through [SECURITY.md](SECURITY.md), not public issues. Preserve existing input-safety checks around archives, XML, relationships, paths, URLs, subprocesses, and generated artifacts.

## Commits, PRs, And Releases

Follow [CONTRIBUTING.md](CONTRIBUTING.md) and [docs/standards/engineering-process.md](docs/standards/engineering-process.md).

- Use conventional commit style.
- Sign commits with `git commit -s`.
- Keep each PR to one logical change.
- Rebase on `main` before review or push when needed.
- Document testing performed.
- Update `CHANGELOG.md` for user-visible public API, CLI, schema, package metadata, security behavior, or deprecation changes.
- Do not move public tags. Do not force-push unless a maintainer explicitly instructs you to do so for the branch at hand.
