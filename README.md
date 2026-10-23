# Excel Period Transfer

A small desktop application that copies period (quarter) columns from a source Excel workbook into an
existing destination table, matching records by user-selected key columns instead of row position.

* Formats: `.xlsx`, `.xlsm` (VBA project preserved) and `.xlsb` (binary workbooks, read and written natively).
* Single or composite keys (for example **ProductCode + Region**), with explicit source-to-destination column mappings.
* Several period columns per run (for example `2026 Q1 -> Q1 Actual`, `2026 Q2 -> Q2 Actual`).
* Mapped destination columns are cleared across all data rows, then filled from matching records. Unmatched
  destination rows end up blank in those columns; unmatched source rows are reported and skipped; rows are never appended.
* Validation before anything is written: duplicate keys, missing keys, formulas without stored results,
  protected sheets, merged cells, partial shared formulas, calculated table columns, output path safety.
* The result is always a new workbook. Inputs are never modified; the edit happens on a working copy, the copy is
  re-read and verified cell by cell and part by part, and only then moved into place.
* One self-contained executable per platform (Windows x64/Arm64, macOS, Linux). No runtime, no installer.

| Document | Contents |
|----------|----------|
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | Technology-stack comparison and recommendation, design, matching rules, limitations, what was verified |
| [docs/USER-GUIDE.md](docs/USER-GUIDE.md) | Step-by-step guide for colleagues |
| [docs/DEVELOPER-GUIDE.md](docs/DEVELOPER-GUIDE.md) | Setup, build, test, packaging commands, release checklist |
| [examples/README.md](examples/README.md) | The ProductCode + Region example with before/after values |

## Quick start (developers)

```bash
dotnet test                                  # core, scenario and headless UI tests
dotnet run --project tools/ExampleGenerator  # regenerates examples/*.xlsx and runs the example transfer
dotnet run --project src/ExcelTransfer.App   # starts the application
```

Packaging (one file, no .NET installation needed on the target machine):

```bash
dotnet publish src/ExcelTransfer.App -c Release -r win-x64   -o publish/win-x64
dotnet publish src/ExcelTransfer.App -c Release -r win-arm64 -o publish/win-arm64
dotnet publish src/ExcelTransfer.App -c Release -r osx-arm64 -o publish/osx-arm64
```

See the developer guide for signing, SmartScreen and distribution notes.

## Repository layout

```
src/ExcelTransfer.Core      engine: readers/editors for xlsx/xlsm (Open XML SDK) and xlsb (own BIFF12 code), matching, planning, execution, verification
src/ExcelTransfer.App       Avalonia desktop wizard (code-behind style, no MVVM framework)
tests/ExcelTransfer.Tests   xUnit v3: unit tests, end-to-end scenarios, xlsb reader/editor tests against Excel-made fixtures, headless UI tests
tools/ExampleGenerator      builds the example workbooks and runs the example transfer
examples/                   generated example workbooks (source, destination, result)
docs/                       documentation
```
