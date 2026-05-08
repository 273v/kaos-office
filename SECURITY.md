# Security policy

## Reporting a vulnerability

We take security seriously. If you believe you have found a security
vulnerability in `kaos-office`, please report it privately so we can address it
before public disclosure.

**Please do not file a public GitHub issue for security reports.**

### How to report

Use [GitHub Private Vulnerability Reporting](https://github.com/273v/kaos-office/security/advisories/new)
to send a report. Alternatively, email **security@273ventures.com**.

Include as much of the following as you can:

- A description of the vulnerability and its impact
- Steps to reproduce, including affected versions
- Any proof-of-concept code, if available
- Suggested mitigations, if you have any

### What to expect

- **Acknowledgement** — within 3 business days of your report.
- **Initial triage** — within 7 business days, including a severity assessment.
- **Fix and disclosure** — coordinated with you. Our target window is 90 days
  from acknowledgement to public disclosure, faster for high-severity issues.
- **Credit** — we credit reporters in the release notes and security advisory
  unless you prefer to remain anonymous.

## Supported versions

`kaos-office` follows Semantic Versioning. While the project is pre-1.0, only
the latest minor release receives security fixes. After 1.0, the latest two
minor releases will be supported.

| Version | Supported |
|---------|-----------|
| 0.1.x   | Yes       |
| < 0.1   | No        |

## Scope

In-scope:

- The `kaos-office` Python package as published on PyPI
- The `273v/kaos-office` GitHub repository (CI, release, supply chain)
- OPC (Open Packaging Conventions) layer security — ZIP-bomb detection,
  path-traversal prevention, file-size caps, XML-bomb / external-entity
  protection (`lxml.etree.XMLParser(resolve_entities=False)`)
- DOCX, PPTX, XLSX reader input handling — malformed OOXML, deeply
  nested structures, oversized parts, malicious content types
- Writer output paths — refusal to silently overwrite existing files
  (`force=true` opt-in gate), parent-directory creation behavior
- MCP server (`kaos-office-serve`) — request validation, tool
  annotations (`readOnlyHint` / `destructiveHint` / `idempotentHint` /
  `openWorldHint`), response size caps via `kaos-core` artifact tiers

Out of scope:

- Third-party dependencies (report to the upstream project — `lxml`,
  `python-pptx`, `python-calamine`, `openpyxl`, `pydantic`, `kaos-core`,
  `kaos-content`)
- The `xlsxwriter`-based `kaos_office.xlsx._xlsxwriter_reference` module
  (private; dev-only test cross-validator; not part of the public surface)
- Issues caused by user-supplied configuration that explicitly disables
  safety features (e.g., calling writer tools with `force=true` against
  untrusted output paths)
