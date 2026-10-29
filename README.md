# res-uitvraag

Tooling for the half-yearly data request ("uitvraag") to the municipalities of RES West-Overijssel: generating the request files from the RES monitor, and bringing the returned answers back into the monitor in a reviewed, traceable way.

**Status:** planned — this README describes the intended design; no code yet.

## Why

Each round (voorjaar / najaar) every municipality receives a *Projectoverzicht* — an Excel overview of its large-scale solar and wind projects — to check and update. Until now that file was produced by hand from the monitor (`RES West monitor_data_26_Q4.xlsx`, sheet `Data Veld+Wind`), and returned changes were copied back by hand. Both steps are slow and error-prone: the najaar 2026 template, for example, already lagged the monitor on 4 project names and 25 corrected coordinates by the time it was compared.

The template turns out to be a purely mechanical derivation of the monitor sheet, so it can be generated — and the returned files can be diffed against what was sent instead of relying on red markings.

## Workflow

```
 monitor (.xlsx) ──1. generate──▶ 11 × Projectoverzicht <gemeente>.xlsx ──▶ municipalities
                                                                                │
 monitor (.xlsx) ◀──4. apply── reviewed change list ◀──3. review── 2. collect ◀─┘
```

1. **Generate** — read the monitor and write one file per municipality, containing:
   - an *explanation* tab: what we ask, how to mark changes, a column legend, contact details;
   - a *data* tab with only that municipality's active projects.
2. **Collect** — read all returned files and compare every cell with what was sent (a snapshot of the sent values is kept alongside the generated files). Output: a change list per project and field (old → new), plus anomalies (unparseable values, new rows, deleted rows).
3. **Review** — the change list is checked by a person. Nothing reaches the monitor without that review.
4. **Apply** — approved changes are written into the monitor, cells marked so changes are traceable.

## Design decisions

- **One file per municipality, not one file with a selector.** A formula-based filter (`FILTER`) produces read-only output, an auto-filter driven by a dropdown needs a macro (`.xlsm`, often blocked by municipal IT), and either way every municipality would receive every other municipality's data.
- **Change detection by diff, not by red marking.** Red marking stays in the instructions as a courtesy, but the collect step compares values, so unmarked changes are still caught.
- **The monitor is only read by the generator, never saved by Python.** The monitor contains slicers, pivots and data-validation extensions that `openpyxl` drops on save. The apply step must therefore write values without round-tripping the workbook through `openpyxl` (approach to be decided).
- **The layout is defined in code, not in a seed file.** `COLUMNS` / `GROUPS` in `layout.py` define every column (header, monitor source, width, number format, units, hidden). They were taken once from `Projectoverzicht template najaar 2026.xlsx`; that file is no longer used. Deviations from it: `RES regio` removed; `Gemeente` and the coordinates hidden; `Kolom1` replaced by the monitor's `Netcongestie`; dates pinned to Dutch month names (`[$-413]mmm-yy`); `Verwachte realisatie` formatted as a year; empty monitor cells stay empty instead of `0`.
- **Derived from the monitor, with these transformations** (matching the existing template):
  - stopped projects (`0 Project gestopt`) are skipped entirely;
  - the column set is exactly that of the current template (najaar 2026) — municipalities see all of those columns, including derived values; monitor-only columns (e.g. `Vollasturen`) are not exported;
  - formulas are flattened to values;
  - `STATUS RES` is split into `Fase_ID` and `Fase`;
  - the monitor's column A (internal notes such as "geen reactie") is **never** exported;
  - empty cells stay empty instead of the monitor's `0` placeholder.

## Decided

- Stopped projects are skipped, as in previous rounds.
- Municipalities see all columns currently in the template.

## Open points

- Explanation text for the explanation tab — to be supplied; until then the tab holds a placeholder.

## Data and privacy

The monitor and all generated / returned files contain project data and stay **outside git**: `input/`, `output/` and `returned/` are git-ignored, as are all `*.xlsx` / `*.csv` files anywhere in the tree. Never commit them, and check `git status` before every commit.

## Related

- [`npres-pipeline`](../npres-pipeline) — syncs the same monitor into the NP RES GeoPackage; its README documents the monitor's column semantics and known data quirks.
