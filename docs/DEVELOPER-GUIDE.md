# Developer guide

## Prerequisites

| Tool | Version | Notes |
|---|---|---|
| .NET SDK | 10.0.1xx (LTS) | `dotnet --version` should print 10.0.x. Install from https://dotnet.microsoft.com/download/dotnet/10.0 or your package manager (`apt install dotnet-sdk-10.0` on Ubuntu 24.04). |
| IDE (optional) | Visual Studio 2022 17.14+, Rider, or VS Code with C# Dev Kit | The Avalonia XAML previewer needs the Avalonia extension. |
| Git | any | |

No Excel installation is needed to build, test or run.

## Dependency versions

| Package | Version | Used by | Licence |
|---|---|---|---|
| DocumentFormat.OpenXml | 3.5.1 | Core (`.xlsx`/`.xlsm`) | MIT |
| Avalonia, Avalonia.Desktop, Avalonia.Themes.Fluent, Avalonia.Fonts.Inter | 12.1.4 | App | MIT |
| ClosedXML | 0.105.1 | Tests and ExampleGenerator only (builds fixtures, reads results independently) | MIT |
| xunit.v3 / xunit.runner.visualstudio / Microsoft.NET.Test.Sdk | 3.2.2 / 3.1.5 / 18.11.0 | Tests | Apache 2.0 / MIT |
| Avalonia.Headless.XUnit, Avalonia.Skia | 12.1.4 | Tests (headless UI tests, screenshots) | MIT |

The product itself depends only on the Open XML SDK and Avalonia. The `.xlsb` engine has no third-party dependency.

## Build and run

```bash
git clone <repo> && cd <repo>
dotnet build                                   # whole solution (ExcelTransfer.slnx)
dotnet run --project src/ExcelTransfer.App     # start the wizard
dotnet run --project tools/ExampleGenerator    # regenerate examples/ and run the example transfer
```

## Tests

```bash
dotnet test                                    # everything: unit, scenario, xlsb, headless UI
dotnet test --filter "FullyQualifiedName~Biff12"           # only the .xlsb engine tests
EXCELTRANSFER_SCREENSHOT_DIR=docs/images dotnet test --filter "FullyQualifiedName~UiTests"   # also writes the guide's screenshots
```

Test layout:

* `CellsTests`, `KeyNormalizerTests`: pure unit tests of references, formats, dates, normalisation, composite keys.
* `FileCheckTests`, `PlannerTests`: file rejection messages, mapping validation, output path rules, duplicate/missing keys.
* `TransferScenarioTests`: end-to-end xlsx/xlsm runs built with ClosedXML and read back with ClosedXML (an independent parser).
* `Biff12Tests`: reader vs SheetJS reference dumps, record round-trips, editor on an Excel-made workbook, transfers to/from xlsb.
* `UiTests`: Avalonia headless; drives the wizard through all seven steps.

Fixtures and their provenance are listed in `tests/ExcelTransfer.Tests/Fixtures/README.md`.

### Cross-checking `.xlsb` output with independent parsers (manual)

Node and Python are only needed for this optional check:

```bash
npm install xlsx@0.18.5 && pip install pyxlsb==1.0.10
node -e "const X=require('xlsx');const wb=X.readFile(process.argv[1]);for(const n of wb.SheetNames)console.log(n, X.utils.sheet_to_csv(wb.Sheets[n]))" out.xlsb
python3 -c "from pyxlsb import open_workbook as o;import sys
with o(sys.argv[1]) as wb:
  for n in wb.sheets:
    with wb.get_sheet(n) as s: print(n, [[c.v for c in r] for r in s.rows()])" out.xlsb
```

## Packaging

The app project publishes as a self-contained single file (settings in `ExcelTransfer.App.csproj`, applied only
during `dotnet publish`). One command per target; cross-publishing from any OS works for Windows and Linux
targets, macOS bundles should be produced on macOS when they are to be signed/notarised.

```bash
dotnet publish src/ExcelTransfer.App -c Release -r win-x64   -o publish/win-x64     # Windows 10/11 x64  -> ExcelPeriodTransfer.exe
dotnet publish src/ExcelTransfer.App -c Release -r win-arm64 -o publish/win-arm64   # Windows on Arm
dotnet publish src/ExcelTransfer.App -c Release -r osx-arm64 -o publish/osx-arm64   # Apple silicon
dotnet publish src/ExcelTransfer.App -c Release -r osx-x64   -o publish/osx-x64     # Intel Mac
dotnet publish src/ExcelTransfer.App -c Release -r linux-x64 -o publish/linux-x64
```

Each folder contains exactly one executable (symbol files are excluded by the project file). The win-x64 build measured 51.8 MB (compressed single file). The executable extracts
native libraries to `%TEMP%\.net` (Windows) or `~/.net` on first start.

Windows distribution notes:

* **Code signing** avoids the SmartScreen warning once the certificate has reputation:
  `signtool sign /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 /a publish\win-x64\ExcelPeriodTransfer.exe`.
  Use the organisation's certificate; EV certificates no longer grant instant reputation.
* Without signing, distribute from a trusted internal share and tell users about *More info -> Run anyway*;
  or have IT whitelist the file hash.
* Antivirus false positives are rare for .NET single-file apps but possible for any new binary; submit the file
  to the vendor if it happens.

macOS notes: wrap the published executable in an `.app` bundle (an `Info.plist` and `Contents/MacOS/`), then
`codesign --deep --options runtime` and notarise with `notarytool`; unsigned apps need right-click -> Open.

### Release checklist

1. Bump `<Version>` in `Directory.Build.props`.
2. `dotnet test` green.
3. `dotnet run --project tools/ExampleGenerator` and open the three example files in Excel.
4. Open one transferred `.xlsb` and one `.xlsm` in Excel; confirm no repair prompt, macros intact, recalculation happened.
5. Publish per target, sign, distribute.

## Performance check

`TransferPlanner` and the editors hold the selected sheets in memory. Measured with a generated 100,000-row
source and destination (8 and 10 columns, destination as an Excel Table with a formula column), Release build,
Linux container with 4 vCPU:

| Step | Time | Notes |
|---|---:|---|
| Plan (read both sheets, key index, matching, structural checks) | 16 s | 100,000 x 100,000 records, composite key |
| Execute (copy, edit 400,000 cleared + 400,000 written cells, save, read back and verify) | 18 s | includes the full cell-by-cell and part-by-part verification |
| Peak working set | 1.8 GB | dominated by the Open XML SDK DOM of the two sheets plus the in-memory snapshots |
| Files | source 3.6 MB, destination 4.8 MB, output 4.9 MB | |

For the typical sizes of planning workbooks (a few thousand rows) everything completes in well under a second.
If much larger sheets become common, the first optimisation is to read sheets with the SDK's streaming
`OpenXmlReader` instead of the DOM and to drop the snapshot after planning; the design keeps that change local
to `OpenXmlSource`.

To repeat the measurement build a scratch console project that references `ExcelTransfer.Core` and ClosedXML,
generate the two workbooks, and time `TransferPlanner.Plan` and `TransferExecutor.Execute` (the snippet used is
in the repository history under `scratchpad/harness/PerfSmoke` in the development notes).

## Extending

* **New container format**: implement `IWorkbookSource` (snapshot) and `IWorkbookEditor` (clears/writes/finish),
  add the format to `WorkbookFormat`/`WorkbookFormats`, `WorkbookSources.Open`, `WorkbookEditors.CreateWorkingCopy`
  and `WorkbookFileCheck`. Planner, executor, verifier and UI need no change.
* **New validation rule**: add it in `TransferPlanner` and return a `ValidationIssue` with a stable `Code`; the UI
  renders issues generically.
* **Changing matching rules**: `KeyNormalizer` holds the logic, `MatchingRules.Describe()` the text shown to users;
  keep both in sync and extend `KeyNormalizerTests`.
