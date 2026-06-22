# Advay June – Area Cleaning & Tower Identification

Cleans the Index II export by extracting the genuine RERA carpet areas and the
tower/wing from the Marathi property descriptions (the reliable ground truth;
the English column is machine-translated and often garbled).

## Files
- `clean_areas.py` – extraction script (fills the workbook in place).
- `build_master.py` – builds the Master Data sheet and backfills blanks that have a consistent internal match.
- `Advay_June_All_projects.xlsx` – cleaned output (sheets: `Index II`, `Master Data`).

## Filled columns
| Column | Meaning |
|---|---|
| `Flat Area (Sqm)` | flat RERA carpet, sq.m (converted from sq.ft ÷10.764 where the deed quoted feet) |
| `Balcony area (Sqm)` | balcony, sq.m |
| `Additional Area (Sqm)` | other attached carpet (adjacent *लगत* / utility *उपयोगिता* / EBVT / purchased) |
| `Actual Areas (Sqft)` | `max(deed's stated total, Flat + Balcony + Additional) × 10.764` |
| `Tower Name` | tower / wing from the description |
| `Area and Price Remarks` | flag for rows with no stated carpet area |

## Results
- 2,579 / 2,831 rows got an Actual Area; 2,595 got a tower.
- 99.8% of computed areas reconcile (±4 sq.ft) with the existing transacted-area columns.
- 252 blanks are deeds that state no carpet area (only built-up or just parking) — each flagged in remarks.

## Master Data sheet
A catalog keyed on **Project + Tower + Flat Line** (the unit/stack from the flat
number, e.g. Sky City `5606` & `5906` → line `06`), so a row with no area in its
description can be looked up by its flat number.

| Source | Meaning |
|---|---|
| `Extracted from descriptions` | area derived from deeds (with #transactions and area spread) |
| `NEEDS EXTERNAL DATA` | no deed states an area for this line — fill from the external competition master |

Lines still needing external data (223 Index II rows): **Skyline Icon**, **Lodha
Altus A&B** (the few lines with no described flat), plus 1 Shraddha + 2 Winter Green.
29 internally-matchable blanks were auto-backfilled (marked in `Area and Price Remarks`).

## Run
```
python3 clean_areas.py     # 1. extract areas/towers into Index II
python3 build_master.py    # 2. build Master Data sheet + backfill blanks
```
`clean_areas.py` reads `SRC` (the original upload) and writes `OUT`.
