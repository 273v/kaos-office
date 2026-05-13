# Fixture provenance — docx

Real-world DOCX corpus for the `kaos-office` DOCX reader / writer round-trip
suite. The collection mixes (a) hand-crafted minimal DOCXs that
exercise specific OOXML features (footnotes, comments, multi-paragraph
flow) and (b) public-domain or publicly-distributed government /
publisher templates that exercise real layout, redlines, comments,
forms, and styles. All files were vendored from the prior
`kelvin_office/tests/resources/docx/` corpus (see commit `e03b9b5`,
`Add real-world DOCX/PPTX/XLSX/SQLite fixtures…`, 2026-04-03 and
follow-up audit `KO-006` in commit `3b71460`, 2026-05-07) which
itself sourced them per the notes below.

Per `docs/oss/50-data-and-fixtures/provenance-policy.md`: every row
points at a public upstream or marks the file as hand-crafted, and no
file in this directory is a customer / privileged / pseudonymized
document.

| File | Source | License | Retrieved | SHA-256 | Notes |
|------|--------|---------|-----------|---------|-------|
| 1444711772592.docx | The Hartford insurance form / claim document (creator `C. Peter Hitson`, Company `The Hartford`, Category `Company Confidential`). Public-record / publicly-distributed insurance template surfaced by web crawl; the `#C0nf1d3nti@l#` keyword is the upstream document's own marking, **not** ours — the file is intentionally retained to exercise reader handling of confidentiality keywords + revision history. | `<unknown — needs verification>` (publisher-distributed template; license not asserted in the file metadata) | 2026-04-18 | ffe972b51714f93cace0dddfbe7bbc829042d5c988d7657f66624fc528d141c9 | Exercises Microsoft Office Word styles + Hartford-template numbering; ported from kelvin-office. |
| Burnout_Intervention_Planning_Guide_Fillable_Form_1.docx | CDC / NIOSH publication "Burnout Intervention Planning Guide" (fillable form). Creator `Novicki, Emily (CDC/NIOSH/OD/ODDM)`. Distributed by CDC NIOSH. | Public domain — 17 USC §105 (US Government work) | 2026-04-18 | 77e5aad03856e27f350b742e6c2ca413d7208fb3fcbac93ad383a318ff87c74b | Exercises fillable form fields (`w:sdt` content controls) for the SDT Phase 6.2 round-trip path. |
| CheeseSample.docx | National FFA Organization "Dairy Foods Career Development Event" handbook ("DAIRY FOODS / CAREER DEVELOPMENT EVENT / Revised 2023"). Public FFA student handbook material. Modified locally (Apr 2025) for OOXML edge-case testing — original retains FFA structure + content but headings/list nesting were re-saved through LibreOffice. | FFA publication — publicly distributed student handbook; treat as fair-use educational fixture | 2026-04-18 | b2bc68a6488852a01f95fb9cd4ea6bd5a4f339c8be0f66b6de39c8a982f1fa78 | Heading depth + ordered-list nesting + Unicode characters (🧀, ƒ∞d§) — used in `TestModificationRoundTrip` for emoji / non-ASCII fidelity. |
| Footnote.docx | Hand-crafted DOCX created in LibreOffice 25.2 for footnote-extraction regression. Minimal body + one `w:footnote`. | Hand-crafted, 273V (Apache-2.0, this repo) | 2026-04-03 | 7661c4d3b3c3d0b7b8526106379b023d3f91e4b0a6790cd30f67c60ef8079f38 | Smallest fixture (~6 KB); pins reader handling of `footnotes.xml` references. |
| Footnote-Edit.docx | Hand-crafted DOCX — variant of `Footnote.docx` with footnote text edited; used to verify edit-round-trip preserves untouched content. | Hand-crafted, 273V (Apache-2.0, this repo) | 2026-04-03 | b24e7e5c2c1231fdbf89518253525aa3de9d7de789882f93d305f2c65727bbfc | Paired with `Footnote.docx` for diff-style modification tests. |
| Footnote_with_comment.docx | Hand-crafted DOCX — `Footnote.docx` plus a single `w:comment` referencing the footnote. | Hand-crafted, 273V (Apache-2.0, this repo) | 2026-04-03 | a6d3542ab308d8d731cc2c0e49aabbd24e997cf3f2fad4cf00e79b002c28c5b9 | Exercises `comments.xml` write-back. |
| Letter of Commitment for Packaged Furniture Program.docx | Vendor-supplied commitment-letter template (`DanaRosa`, 2024-02-22). Public template distributed by a packaged-furniture vendor. | `<unknown — needs verification>` (vendor-distributed letter template; license not asserted) | 2026-04-18 | f0e2b2e8924e74b047bb0f49c08c85fa2e1bde788b339272aa06396d1e9da4d0 | Exercises body-only flow + simple paragraph styles. |
| MCSRedline10312022.docx | Master Concession Services (MCS) redline of an unspecified agreement (2022-10-31). Public redline distribution; precise upstream URL not recorded at vendoring time. The file contains track-changes (`w:ins` / `w:del`) used by `TestRoundTrip` to verify the reader's track-changes policy (accept insertions, skip deletions). Audit `KO-003` (2026-05-07) fixed a `row_span=0` regression triggered by this fixture's `vMerge` cells. | `<unknown — needs verification>` (publicly distributed redline; license not asserted) | 2026-04-18 | 9191fa2bbfe62dea3748e45daf6dcc19e147a5af0c1e6af1b3fd6922dab0d97f | Largest DOCX (~1.7 MB); heavy table + vMerge + redline coverage. |
| MultiParagraphSample.docx | Hand-crafted DOCX with several paragraphs of plain prose; created via LibreOffice 25.2. | Hand-crafted, 273V (Apache-2.0, this repo) | 2026-04-03 | 74cd6ed0033b7b804b44b5f8b7e160deeb042a6c239eedc9d3f9447bc40d0470 | Smoke test for the paragraph/Heading/Bullet block dispatch. |
| PolicyProcedureTemplate_PhysicalFacility_Final.docx | Administration for Community Living (ACL) "Physical Facility Policy Template". ACL is a US HHS operating division. | Public domain — 17 USC §105 (US Government work) | 2026-04-18 | 55d767c928ca6a5a65d5257a82e42a6536549240d8bb636a090238eb5ab86128 | Exercises policy-style headings + numbered lists. |
| Toro 2022 Term Loan.docx | SEC EDGAR exhibit — The Toro Company 2022 term-loan agreement (`Exhibit 10.1`, CUSIP 891091AR1). Filed publicly with the SEC. Original creator `Windows User`; last modified by Michael Bommarito (2023-10-10) when vendored as a Kelvin fixture. | Public domain — US federal filing (SEC EDGAR public record) | 2026-04-18 | 46ae287460ead02dd9004731dbffc722b09472385cdad419fb3521ba612c1a33 | Largest legal-contract fixture; used as the gold round-trip target. |
| Toro 2022 Term Loan - Comments.docx | Derived from `Toro 2022 Term Loan.docx` with reviewer comments added (multiple `w:comment` entries). Generated locally (LibreOffice 25.2, 2025-05-27). | Derived work over a public SEC filing (273V annotations Apache-2.0) | 2026-04-18 | 9bdd5fd67493d8e7a6ce7932e5486b0dc7d4d79a9894496bff55ac645a3ca705 | Exercises `comments.xml` reader; paired with `Toro 2022 Term Loan.docx`. |
| Toro 2022 Term Loan - Redline v1.docx | Derived from `Toro 2022 Term Loan.docx` with track-changes (`w:ins` / `w:del`) added. Generated locally (LibreOffice 25.2, 2025-05-30). | Derived work over a public SEC filing (273V annotations Apache-2.0) | 2026-04-18 | 5d184244c7ce54898b92d2e361bf150da5a4db4f516f61adf8e9d23153ba7c2a | Exercises track-changes policy on a non-trivial contract. |
| bcfp_consumer-rights-summary_2018-09.docx | Consumer Financial Protection Bureau "Summary of Consumer Rights" model form (March 2023 revision, originally 2018-09 model form). Creator `Consumer Financial Protection Bureau`, Company `The U.S. Department of the Treasury`. Distributed by CFPB. | Public domain — 17 USC §105 (US Government work) | 2026-04-03 | ad82a76d5298a2e81e50123130468bf260c1839567d273e71cee8c783780d29b | CFPB-issued consumer-disclosure model form. |
| cms-10704-hra-model-notice.docx | Centers for Medicare & Medicaid Services (CMS) / CCIIO "Individual Coverage HRA Model Notice" (CMS form 10704). Creator `CMS/CCIIO`, Company `Center For Medicaid Services`. | Public domain — 17 USC §105 (US Government work) | 2026-04-18 | db4c36ed981cb4c237ded552b54db2c0a365458ab222046ba3717ac47fdd4f10 | CMS-distributed model notice. |
| mutual-to-stock-application-for-conversion.docx | OTS / OCC Form 1680 ("Application for Conversion from Mutual to Stock Form", OMB 1550-0014). Creator `shackletterh`, Company `OCC`. | Public domain — 17 USC §105 (US Government work) | 2026-04-03 | bf16afce13fb762424ac57aed64de83834ea41cea717dad21c3eef9e7a9c4c51 | Office of the Comptroller of the Currency regulatory form. |
| p2021-203386.docx | Australian Government — Department of the Treasury (Australia) — research paper "The Treasury Macroeconometric Model of Australia — Modelling approach" (Treasury working paper, Jared Bullen, 2021). Re-saved locally via LibreOffice 25.2 during vendoring. | Commonwealth of Australia — Treasury Working Paper; distributed publicly. Treasury content is typically released under CC-BY-4.0 (Australia Government open access licence). License **not** asserted in the file metadata; flag for review if redistributed. | 2026-04-03 | 2e6e6291f6b208e28097cf0a59fed82b1a0577e1e2b39a0d3dd975ca90caf145 | Only non-US-government fixture; included for non-en-US locale + econometric-equation typography. |
| right-to-use-leases-with-operating-budget-treatment-with-cancellation-clause.docx | Bureau of the Fiscal Service (Treasury — `BFS`) accounting policy memo on right-to-use leases. Creator `Regina D. Epperly`, Company `BFS`. | Public domain — 17 USC §105 (US Government work) | 2026-04-18 | 285fda7a0299a3870edb5d24548708870ff1bf10bda8797ec83b7e5b519e08be | Multi-page government policy memo. |

## TODO (human verification)

The rows below have provenance recovered from file metadata but no
recorded public source URL or explicit license assertion at the time of
vendoring. Flag for follow-up: confirm the upstream URL and license,
or replace with an equivalent hand-crafted fixture.

- `1444711772592.docx` (The Hartford insurance template — confirm public
  distribution + license terms; the `#C0nf1d3nti@l#` keyword is the
  upstream document's own marking but warrants a second look that
  redistribution as a test fixture is appropriate).
- `Letter of Commitment for Packaged Furniture Program.docx`
  (vendor-distributed letter template — confirm upstream URL).
- `MCSRedline10312022.docx` (Master Concession Services redline —
  confirm publication source + license).
- `p2021-203386.docx` (Australian Treasury working paper — confirm
  CC-BY-4.0 applicability and add explicit attribution row if so).

## Regenerating

```bash
cd kaos-office
sha256sum tests/fixtures/docx/*.docx
# update SHA-256 column above + the matching `Retrieved` date
# (git log -1 --format=%cI -- tests/fixtures/docx/<file>)
```
