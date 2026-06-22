# Advay June – Area Cleaning & Tower Identification

Cleans the Index II export by extracting the genuine RERA carpet areas and the
tower/wing from the Marathi property descriptions (the reliable ground truth;
the English column is machine-translated and often garbled).

## Files
- `clean_areas.py` – extraction script (fills the workbook in place).
- `Advay_June_All_projects.xlsx` – cleaned output.

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

## Run
```
python3 clean_areas.py
```
Reads `SRC` (the original upload), writes the cleaned workbook to `OUT`.
