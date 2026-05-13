# Fixture provenance — xlsx

5 XLSX fixtures for the `kaos-office` XLSX reader + writer + the
`kaos-tabular` integration tests. They mix small smoke fixtures
(`states`, `ledes98b` — both regenerated locally from the kelvin
tabular CSV/JSON canonical copies) with three real-world XLSX files
that exercise multi-sheet workbooks, merged cells, and large row
counts.

Vendored from the prior `kelvin_office/tests/resources/xlsx/` and
`kelvin_tabular` corpora (commit `e03b9b5`, 2026-04-03).

| File | Source | License | Retrieved | SHA-256 | Notes |
|------|--------|---------|-----------|---------|-------|
| CBS-BNSF-2015-Q4.xlsx | BNSF Railway — STB Common-Sense Benchmark (CBS) data, Q4 2015 (creator `Stanton Parker`, Company `BNSF Railway`, lastPrinted 2016-01-25). Distributed publicly by BNSF to the US Surface Transportation Board (STB) as part of the regulatory record. | Public regulatory filing (STB record); treat as fair-use research / regulatory disclosure | 2026-04-03 | d47db08e0c9bae25771b2193f7ad10fffc930da87cc8760fa6b4ccfef8f6591e | 2 sheets; exercises multi-sheet `list_sheets` and merged-cell handling. |
| ledes98b.xlsx | Hand-crafted minimal XLSX matching the LEDES 98B legal e-billing schema (same data as the kelvin_tabular `ledes98b.csv` / `.json` / `.sqlite` siblings). Created by `xlsxwriter` / `openpyxl` for the 4-way cross-format JOIN integration test (`engine.register_file()` smoke). | Hand-crafted, 273V (Apache-2.0, this repo) — LEDES 98B is an open billing schema, no fee data is real | 2026-04-03 | 700e4f8029d74a526a8b849cd070ba4e08c7c7e00b9994c50319e344199f6770 | Smallest fixture (~6 KB); pairs with `states.xlsx` for 4-way JOIN tests. |
| payment-report-07-01-20-thru-07-15-20.xlsx | Bureau of the Fiscal Service (Treasury — Company `BFS`) — biweekly payment-report distribution (creator `L Mills`, lastPrinted 2020-07-24). 342 rows. | Public domain — 17 USC §105 (US Government work) | 2026-04-03 | d00215c5a36d1e0f198908f4a854ea58f024dde7dfb8437fcece2b2ae770eef7 | Real-world federal payment report; exercises typed-column inference (DATE, MONEY, INTEGER). |
| ppplf-transaction-specific-disclosures-07-13-21.xlsx | Federal Reserve — Paycheck Protection Program Liquidity Facility (PPPLF) — transaction-specific disclosures (2021-07-13). 14,429 rows. Distributed publicly via the Federal Reserve's PPPLF disclosure portal. | Public domain — 17 USC §105 (US Government work) | 2026-04-03 | ae3e03c8fafaa4ff49cd69398433e3ece3a96efe30d2a9826a7ce704a3a36c8c | Largest XLSX (~959 KB / 14k rows); performance budget fixture (must parse in < 3 s). |
| states.xlsx | Hand-crafted XLSX with the US states (name, abbreviation, capital, population) — same data as the `kelvin_tabular` `states.csv` / `.json` / `.sqlite` siblings. Used by the 4-way cross-format JOIN integration. | Hand-crafted, 273V (Apache-2.0, this repo) — facts (state names, capitals, populations) are public-domain | 2026-04-03 | 3a4748fb80ebf1e381cfbfe979368132025ddf4823a3f68a73b9485c1f6edcbc | Pairs with `ledes98b.xlsx` for the 4-way JOIN smoke test. |

No file in this directory is a customer / privileged / pseudonymized
document.

## Regenerating

```bash
cd kaos-office
sha256sum tests/fixtures/xlsx/*.xlsx
# update SHA-256 column above + the matching `Retrieved` date
# (git log -1 --format=%cI -- tests/fixtures/xlsx/<file>)
```
