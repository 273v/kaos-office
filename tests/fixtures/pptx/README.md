# Fixture provenance — pptx (top level)

Real-world PPTX fixtures for the `kaos-office` PPTX reader regression
and benchmark suites. These are the *production-scale* decks (194 KB –
~550 KB) that exercise typical legal/government slide content: title
slides, bullets, charts, SmartArt, tables, speaker notes, and the
Hartford-style insurance / Treasury / CDC layouts. Smaller hand-crafted
fixtures live in `battle/` (python-pptx-generated, ~30 KB each) and
upstream-Apache-POI corpora live in `stress/` — see those directories'
own READMEs.

Vendored from the prior `kelvin_office/tests/resources/pptx/` corpus
(commit `e03b9b5`, 2026-04-03) plus the audit-01 KO-006 vendoring
(commit `3b71460`, 2026-05-07) which moved the IEO2021 deck out of a
hard-coded user-home path and into the in-repo fixture tree. Multi-MB
decks (e.g. CIPLA_CLEVELAND_BAR_DEC_2023.pptx) are kept out-of-repo
behind `KAOS_OFFICE_EXTERNAL_FIXTURES_DIR`.

| File | Source | License | Retrieved | SHA-256 | Notes |
|------|--------|---------|-----------|---------|-------|
| Hello-World.pptx | LibreOffice Impress "Grey Elegant" template by Ahmad Bayhaqi Saputra (icons from material.io / icons8, illustrations from Pixeltrue, photos from Unsplash). Distributed via the LibreOffice / Indonesian LibreOffice Community templates collection. | CC-BY-SA-4.0 (declared in the file's `dc:description`) | 2026-04-03 | 86bed0ff219bb8d973d3d9f4590b48aaf8737ea715d507cd41f7adfae9e5185c | Used as the canonical PPTX smoke-test fixture. The CC-BY-SA-4.0 share-alike obligation propagates to derivatives **of the fixture itself**; the kaos-office wheel does not ship the fixture, so the obligation does not propagate to consumers — but if we ever re-publish a transformed copy, that copy is CC-BY-SA. |
| IEO2021_ChartLibrary_Industrial.pptx | US Energy Information Administration (EIA) — "International Energy Outlook 2021 — Industrial Chart Library". Creator `Kahan, Ari`, Company `EIA`. Distributed publicly by the EIA. | Public domain — 17 USC §105 (US Government work) | 2026-05-07 | 23912f05e640a89e64c1be10786a870c98260a3e40c341352c70d38c6cdba050 | Heavy chart fixture used by the chart-linearization benchmarks; vendored in audit-01 KO-006 to remove a hard-coded `/home/$USER/...` path. |
| Status report.pptx | Microsoft Office "Status Report" template — the missing `docProps/core.xml` is consistent with Office's stock-template export path, which strips authoring metadata. Surfaced by a public web search of "status report PowerPoint template" during the kelvin-office vendoring pass. Ported from `kelvin_office/tests/resources/pptx/Status report.pptx`. | Microsoft Office Template Service Agreement — Microsoft publishes status-report templates as public downloads. Used here under fair use for parser-regression testing (the deck is empty / placeholder content). | 2026-04-03 | 27331690733357b8eddbc88b5860699663a4e646bb85bcf59af069797112ba7e | Exercises the reader's missing-core-props path on a real-sized deck. |
| Testimony-Mulvey-2013-03-22.pptx | US Surface Transportation Board (STB) — Congressional testimony deck: "Shippers Taking Charge in a Capacity Constrained Environment: The Role of the STB in Railroad Regulation" (2013-03-22). Creator `bezoldk`, Company `Surface Transportation Board`, lastModifiedBy `Government of the United States`. | Public domain — 17 USC §105 (US Government work) | 2026-04-03 | 9a1cb477ec70741cdb4b0c1d0b5fb24b9488a5be53cae986d525526f0b623aa4 | Real Congressional testimony deck; exercises SmartArt + speaker notes + heavy slide-master content. |

## Regenerating

```bash
cd kaos-office
sha256sum tests/fixtures/pptx/*.pptx
# update SHA-256 column above + the matching `Retrieved` date
# (git log -1 --format=%cI -- tests/fixtures/pptx/<file>)
```
