# Test fixtures

Real Excel-produced files used to test the binary (.xlsb) and macro-enabled (.xlsm) engines.
They are test data only; nothing in the application depends on them.

| File | Origin | License |
|------|--------|---------|
| `xlsb/calamine-date.xlsb`, `xlsb/calamine-date_1904.xlsb`, `xlsb/calamine-issues.xlsb` | [tafia/calamine](https://github.com/tafia/calamine) `tests/` | MIT |
| `xlsb/poi-62815.xlsb`, `xlsb/poi-sample.xlsb`, `xlsb/poi-testVarious.xlsb`, `xlsb/poi-Simple.xlsb` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet/` | Apache License 2.0 |
| `xlsm/poi-SimpleMacro.xlsm` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet/` | Apache License 2.0 |

`xlsb/expected/*.cells.txt` were generated with SheetJS Community Edition 0.18.5 (an independent parser)
and serve as the reference values the application's own `.xlsb` reader is checked against.
Format: `Sheet!A1<TAB>kind<TAB>value` where kind is `n` number (dates are raw serials), `s` string, `b` boolean, `e` error.
