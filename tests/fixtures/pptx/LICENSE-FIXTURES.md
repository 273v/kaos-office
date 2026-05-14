# Third-party fixture licenses — share-alike obligations

This file enumerates fixtures in this directory that are **derived from
third-party works under a copyleft / share-alike license**, and records
the attribution + share-alike notices that license requires us to
propagate downstream.

The per-file provenance manifest lives in [`README.md`](README.md); this
file is the dedicated CC-BY-SA-4.0 attribution block that the share-alike
clause asks for. Where both files apply, this file is the source of
truth for licence terms; the README row is the short summary that points
back here.

## CC-BY-SA-4.0 fixtures

### Affected files

| File | SHA-256 |
|------|---------|
| `Hello-World.pptx` | `86bed0ff219bb8d973d3d9f4590b48aaf8737ea715d507cd41f7adfae9e5185c` |

### Upstream evidence

The fixture itself declares its license inside the OOXML package, in
`docProps/core.xml` under `dc:description`:

> This work is licensed under the Creative Commons Attribution-ShareAlike 4.0
> International License by Ahmad Bayhaqi Saputra <bayhaqisptr04@gmail.com>.
> To view a copy of this license, visit
> https://creativecommons.org/licenses/by-sa/4.0/legalcode
> or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.

That string is verifiable directly against the bundled fixture via:

```bash
unzip -p tests/fixtures/pptx/Hello-World.pptx docProps/core.xml
```

### Attribution

- **Author** — Ahmad Bayhaqi Saputra (<bayhaqisptr04@gmail.com>).
- **Title** — *Grey Elegant — LibreOffice Impress Template*.
- **Source / distribution** — LibreOffice Impress templates collection
  (contributed via the Indonesian LibreOffice Community + Gimpscape ID
  call-to-arms for the 10th anniversary of LibreOffice).
- **License** — `CC-BY-SA-4.0` (SPDX identifier). Full text at
  <https://creativecommons.org/licenses/by-sa/4.0/legalcode>. The
  declaration is bundled in the fixture's own `docProps/core.xml`
  (see "Upstream evidence" above).
- **Embedded media credits** (re-stated from the deck's own
  `dc:description`):
  - Icons sourced from <https://material.io> and
    <https://icons8.com/icons>.
  - Illustrations sourced from <https://icons8.com/illustrations>
    (master-slide "Table of content" illustration by Pixeltrue,
    <https://icons8.com/illustrations/author/5ec7b0e101d0360016f3d1b3>).
  - Photos sourced from <https://unsplash.com> (slide 9 photo by
    Dave Hoefler, <https://unsplash.com/fr/@iamthedave>, distributed
    under the Unsplash license at
    <https://unsplash.com/fr/licence>).
- **Retrieved on** — Vendored into kaos-office on 2026-04-03 (from
  the prior `kelvin_office/tests/resources/pptx/` corpus, commit
  `e03b9b5`).

### Indication of changes

273V has **not** modified `Hello-World.pptx` after vendoring. The file
on disk is byte-identical to the upstream copy (the SHA-256 above
matches across the kelvin-office → kaos-office hop). If future commits
modify the deck, this section must be updated to enumerate the
changes per CC-BY-SA-4.0 §3(a)(1)(B).

### Share-alike notice

Any redistribution of this fixture — modified or unmodified, in whole
or in part — must:

1. Carry attribution to Ahmad Bayhaqi Saputra as the original author
   of the *Grey Elegant — LibreOffice Impress Template* (per the
   contact email and source URLs above).
2. Indicate any further modifications relative to the upstream
   template.
3. Be licensed under `CC-BY-SA-4.0` (or a later compatible version
   per CC-BY-SA-4.0 §4(b)).

The fixture is bundled only with the kaos-office test suite and is
**not** included in the published `kaos-office` wheel; the share-alike
obligation therefore does not propagate to library consumers. If a
future release packages the deck or a transformed copy of it inside
a published artifact, that artifact must reproduce this attribution
block alongside the file.

## Non-CC-BY-SA-4.0 fixtures

All other fixtures in this directory (`IEO2021_ChartLibrary_Industrial.pptx`,
`Status report.pptx`, `Testimony-Mulvey-2013-03-22.pptx`) carry no
share-alike obligation. Their license terms are documented per-row in
[`README.md`](README.md).
