# Example: ProductCode + Region keys, Q1/Q2 transferred, Q3 untouched

Regenerate the files and rerun the transfer with:

```bash
dotnet run --project tools/ExampleGenerator
```

Files:

| File | Role |
|------|------|
| `source.xlsx` | Sheet `Data 2026`, ordinary range with header row 1: `ProductCode`, `Region`, `2026 Q1`, `2026 Q2`, `Comment`. Rows are deliberately shuffled, one key differs only by case and spaces (`p-002 ` / `south`), one Q1 value is a formula (`=200+25`), one row (`P-005`/`East`) has no destination match. |
| `destination.xlsx` | Sheet `Plan` with Excel Table `Products` (header row 3, totals row, `Total` formula column, number formats, a coloured `Q3 Actual` column), a note below the table and a second sheet `Notes` with a formula referencing the table. |
| `destination - updated.xlsx` | The result produced by the application. |

Selections used: source = sheet `Data 2026`, header row 1; destination = table `Products`;
keys = `ProductCode (A) -> ProductCode (A)` and `Region (B) -> Region (B)`;
periods = `2026 Q1 (C) -> Q1 Actual (D)` and `2026 Q2 (D) -> Q2 Actual (E)`.

Result (the generator prints the same table):

| Row | ProductCode | Region | Q1 Actual before | Q2 Actual before | Q1 Actual after | Q2 Actual after | Q3 Actual (unselected) |
|----:|-------------|--------|-----------------:|-----------------:|----------------:|----------------:|-----------------------:|
| 4 | P-001 | North | 1.00 | 2.00 | **110.00** | **111.00** | 3.00 (unchanged) |
| 5 | P-001 | South | 4.00 | 5.00 | **115.00** | **116.00** | 6.00 (unchanged) |
| 6 | P-002 | North | 7.00 | 8.00 | **220.00** | *(blank: source cell is empty)* | 9.00 (unchanged) |
| 7 | P-002 | South | 7.50 | 8.50 | **225.00** (stored formula result) | **226.00** | 9.50 (unchanged) |
| 8 | P-003 | South | 10.00 | 11.00 | **330.00** | **331.50** | 12.00 (unchanged) |
| 9 | P-004 | East | 13.00 | 14.00 | *(cleared)* | *(cleared)* | 15.00 (unchanged) |

What the validation reported for this run:

* Warning: 1 source record has no matching destination record and is skipped (`P-005 | East`).
* Note: 1 destination record has no matching source record; its transferred columns are blank (`P-004 | East`).
* Note: 1 value comes from a formula in the source; the stored result is copied.
* Columns to clear: `Q1 Actual (D)`, `Q2 Actual (E)`: 12 cells, 12 currently non-blank. Values to write: 9.

Everything else is unchanged: keys, `Description`, `Q3 Actual`, the `Total` formulas, the totals row, the table
definition, formatting, the note below the table and the `Notes` sheet (byte-identical part). The output is flagged
so that Excel recalculates all formulas when it is opened.
