# Architecture and technology decision

## 1. What the application must do

Copy several period/quarter columns from a source worksheet into an existing destination table, matching
records by one or more user-selected key columns. Supported containers: `.xlsx`, `.xlsm` and `.xlsb`.
Colleagues must be able to run it without installing a runtime; the destination workbook's other content must
survive untouched; nothing may be written until validation and an explicit confirmation have passed.

## 2. Technology-stack comparison

"End user" means a colleague running the tool. "Developer" means whoever maintains it.

| Criterion | C#/.NET 10 desktop (chosen) | Packaged Python (openpyxl + PyInstaller) | Browser app, local processing (SheetJS / ExcelJS) | Java (JavaFX + Apache POI, jpackage) | Excel-hosted (VBA / Office Scripts / Office.js add-in) |
|---|---|---|---|---|---|
| Operating systems | Windows 10 1607+/11, macOS 15+, Linux ([supported OS list](https://github.com/dotnet/core/blob/main/release-notes/10.0/supported-os.md)); GUI via Avalonia on all three | Windows/macOS/Linux, but a build per OS on that OS | Anything with a modern browser | Windows/macOS/Linux, build per OS | Where Excel runs; Office Scripts need Microsoft 365 |
| End-user installation | None: one self-contained `.exe` (or app) copied anywhere; no .NET needed ([single-file publishing](https://github.com/dotnet/docs/blob/main/docs/core/deploying/single-file/overview.md)) | None in theory; in practice PyInstaller one-file binaries are frequently quarantined by antivirus heuristics (widely reported, e.g. [discussion](https://discuss.python.org/t/pyinstaller-false-positive/43171)) | None: open an `.html` file | None if jpackage bundles a runtime | Macro security prompts/policies; add-ins need sideloading or admin deployment |
| Developer requirements | .NET 10 SDK, any editor | Python 3.x, PyInstaller, per-OS build machines | Node toolchain for TypeScript/bundling and tests | JDK 21, JavaFX SDK, jpackage per OS | Excel, VBA editor or TypeScript tooling |
| Distribution size | ~52 MB single compressed file (runtime included; measured for win-x64) | 30-120 MB | < 5 MB | 60-100 MB | tiny |
| Excel handling: `.xlsx`/`.xlsm` | Open XML SDK (Microsoft, MIT): edit individual parts, everything else byte-identical, VBA kept | openpyxl drops charts/images on save ([docs](https://openpyxl.readthedocs.io/en/stable/usage.html)) and keeps either formulas or cached values, not both | SheetJS CE writes data but no styles (styling is a Pro feature per its README); ExcelJS keeps styles but not images, many conditional formats or custom table themes (README) | POI XSSF good fidelity | Excel itself: perfect |
| Excel handling: `.xlsb` | No OSS library edits xlsb with preservation: this project includes its own BIFF12 reader/editor (verified against Excel-made files and two independent parsers) | pyxlsb reads only; no writer | SheetJS CE reads/writes xlsb but regenerates the workbook (formatting lost) | POI reads xlsb (XSSFB) only | Excel itself: perfect |
| Formula results | Stored results read; recalculation requested on open | Same limitation, harder to express | Same | Same | Live |
| Licensing | MIT (Open XML SDK, Avalonia, ClosedXML for tests); no commercial component | MIT/PSF; EPPlus-style dual licences avoided | Apache 2.0 (SheetJS CE) / MIT (ExcelJS) | Apache 2.0 | Microsoft 365 licences |
| Development complexity | Medium: one language, strong typing, good tests; own xlsb code to maintain | Low for xlsx, blocked for xlsb editing | Medium-high: browser file API limits (no in-place save, no paths), weak libraries for preservation | Medium-high | VBA hard to test/version; Office Scripts cannot open a second local file |

Other options considered and rejected: Excel COM automation as the engine (Windows + Excel only, fragile
unattended, impossible to test in CI; kept as a possible future fallback engine), EPPlus (Polyform Noncommercial
since v5, commercial licence required for business use - confirmed in its NuGet package description),
Electron/Tauri (adds a web stack on top of the same library gaps).

### Facts verified against official sources while deciding

* .NET 10 is the current LTS: released 11 Nov 2025, supported until 14 Nov 2028 ([.NET support policy](https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core)).
* Single-file, self-contained publishing bundles the runtime; builds are per OS/architecture; native libraries need `IncludeNativeLibrariesForSelfExtract` ([docs source](https://github.com/dotnet/docs/blob/main/docs/core/deploying/single-file/overview.md)).
* Open XML SDK is MIT licensed and explicitly low level ("not intended to directly provide higher-level abstractions") ([README](https://github.com/dotnet/Open-XML-SDK)).
* Avalonia 12 requires .NET 8+ and recommends .NET 10; file dialogs go through `IStorageProvider`; headless xUnit testing moved to xUnit v3 ([breaking changes](https://docs.avaloniaui.net/docs/avalonia12-breaking-changes), [headless testing](https://docs.avaloniaui.net/docs/testing/headless-xunit)).
* EPPlus: "From version 5 EPPlus changes the license model using a dual license, Polyform Non Commercial / Commercial license" (NuGet package description).
* SheetJS README: styling, images/charts/pivot tables and formula evaluation are Pro features; the Community Edition is Apache 2.0.
* ExcelJS README (MIT): "Images are not currently supported" in several places, custom table themes unsupported, pivot tables "with limitations".
* openpyxl documentation: images and charts in existing files are lost when a file is opened and saved.
* Unsigned executables trigger "Windows protected your PC" until the file/publisher builds reputation ([Microsoft Learn, SmartScreen reputation](https://learn.microsoft.com/windows/apps/package-and-deploy/smartscreen-reputation)).
* BIFF12 record layouts: cross-checked against SheetJS's writer (Excel accepts its output), LibreOffice Calc's importer (`sc/source/filter/oox`, used in production to open xlsb) and the MS-XLSB specification where reachable.

## 3. Recommendation and trade-offs

**Chosen: C# on .NET 10, Avalonia 12 for the GUI, Open XML SDK for `.xlsx`/`.xlsm`, and an in-house BIFF12 engine for `.xlsb`.**

Why: it is the only option on the list that gives (a) a one-file, no-install distribution, (b) preservation-safe
editing of all three formats, (c) a native GUI on Windows and macOS from one code base, (d) a strongly typed,
well-tested code base a colleague can maintain, all under permissive licences.

Trade-offs accepted:

* **Executable size (~52 MB)** instead of a kilobyte-sized HTML file. Nothing is downloaded at run time and no
  runtime is installed, which is the easier story for most IT departments.
* **Unsigned binaries show SmartScreen warnings.** Either sign with the organisation's code-signing certificate
  (`signtool`) or distribute from an internal share/Intune. If your IT blocks unsigned executables outright, the
  core design (package-level patching) could be ported to a browser app, but with the file-saving and xlsb
  limitations above.
* **Own `.xlsb` code.** About 1,500 lines that only this project maintains. It is covered by round-trip tests
  against real Excel-made files and by cross-checks with two independent parsers (SheetJS, pyxlsb). What it is
  *not* covered by is Excel itself, which was not available in the development environment: open one produced
  `.xlsb` in Excel on the first day and keep it as a smoke test.
* **No formula recalculation.** The tool copies the results Excel stored at the source's last save and flags the
  output so Excel recalculates everything when it is opened (which makes Excel ask to save on close).
* **Avalonia instead of WinForms/WPF.** Slightly less native look, in exchange for macOS/Linux support and the
  ability to run the whole wizard in headless UI tests.

## 4. Questions that shaped the architecture, and the defaults used

| Question | Default assumed (change if wrong) |
|---|---|
| Colleagues' operating systems? | Windows 10/11 x64. macOS and Arm64 builds exist; macOS needs notarisation to avoid Gatekeeper prompts. |
| Workbook formats and special features? | `.xlsx`, `.xlsm`, `.xlsb`; charts, pivot tables, VBA, comments, data validation, conditional formatting must survive untouched (they do: untouched parts are copied byte for byte). Password-encrypted workbooks are not supported. |
| Typical file sizes? | Up to ~100k rows x ~20 columns per selected sheet, held in memory. Measured numbers are in section 10. |
| Destination: real Excel Table or ordinary range? | Both: Excel Tables (ListObjects) are listed by name and their header/data/totals rows are derived automatically; ordinary ranges need the header row and the first/last data row (suggested automatically). |
| Duplicate keys | Block, show the records; no first/last/aggregate behaviour. |
| Missing keys | Block by default; the user can opt to skip such records (still reported). |
| Hidden/filtered rows | Included by default; a checkbox excludes them from matching and clearing. |
| Output location | New file `<destination> - updated.<ext>` next to the destination, same format as the destination. |

## 5. Structure

```
ExcelTransfer.Core
  Cells/        CellValue (typed cell content), CellRef/CellRange, NumberFormats (date detection), ExcelDates (1900/1904)
  Workbooks/    IWorkbookSource + SheetSnapshot (format-independent view), IWorkbookEditor (format-independent edits)
                OpenXmlSource / OpenXmlEditor   : .xlsx/.xlsm via DocumentFormat.OpenXml
                Biff12/                         : .xlsb (OpcZipPackage, Biff12Stream, Biff12Source, Biff12Editor)
                WorkbookFileCheck, WorkbookInspector, WorkbookFormat
  Data/         DataSetSelection -> SheetReader -> LoadedDataSet (columns, records, metadata)
  Matching/     KeyNormalizer, KeyPart, CompositeKey (structural equality), MatchingRules (text shown to users)
  Transfer/     TransferSpec -> TransferPlanner -> TransferPlan (issues, matches, clears, writes)
                TransferExecutor (working copy -> edit -> verify -> move), OutputVerifier, OutputPathRules
ExcelTransfer.App     Avalonia wizard: MainWindow (steps), DataSetStep (sheet/table/range + preview), MappingRow
ExcelTransfer.Tests   xUnit v3; ClosedXML builds xlsx fixtures; Excel-made xlsb/xlsm fixtures; Avalonia.Headless UI tests
tools/ExampleGenerator
```

Flow: **select files -> select sheet/table -> map keys -> map periods -> validate & preview -> transfer -> save.**
The UI never touches a workbook itself; every operation is a call into `ExcelTransfer.Core` on a background thread.

### Reading

Both engines produce a `SheetSnapshot`: every stored cell as a `CellValue` (Blank, Text, Number, Date, Boolean,
Error, FormulaWithoutResult) plus hidden-row flags, merged ranges, shared/array/data-table formula ranges,
protection and table definitions. Dates are detected from the cell's number format (built-in ids and custom
format codes) and normalised to 1900-system serials. Formula cells contribute their *cached* result; a formula
without a cached result becomes `FormulaWithoutResult` and blocks the transfer if it is needed. Rich text is
flattened to plain text.

### Planning

`TransferPlanner` loads both data sets, validates mappings, builds keys for every record (ignoring completely
empty records), reports duplicates/missing keys, matches, and computes the exact edit list: every cell of the
mapped destination columns in scope is cleared, then matched values are written. Structural blockers are
detected here (protected sheet, merged cells in the target area, shared or array formulas that extend beyond the
cleared area, what-if data tables, calculated table columns). Nothing is written during planning.

### Executing and verifying

1. The destination is copied to `~<output>.<guid>.tmp.<ext>` in the output folder (for `.xlsb` the new package
   is streamed straight to that file).
2. Only the target worksheet part is edited: cells are cleared (content removed, style kept) or written. New
   strings go to the shared string table; date values written into cells without a date format get a cloned
   cell style with a built-in date format. The calculation chain part is removed and a full recalculation on
   open is requested (`fullCalcOnLoad`; for xlsb additionally `recalcID = 0`).
3. The saved copy is read back and compared with the plan: every planned clear/write happened, every other cell
   in the data set is unchanged, headers and row counts agree, and every package part the editor did not touch
   is byte-identical to the original. Any discrepancy discards the copy and reports an error.
4. The copy is moved to the output path.

## 6. Matching rules (as shown in the application)

1. Records match when every key pair agrees, compared part by part in the selected order; parts are never joined, so ("AB","C") and ("A","BC") differ.
2. Text: trimmed, internal whitespace runs collapsed, non-breaking spaces treated as spaces, zero-width characters removed, case-insensitive, Unicode NFC.
3. Leading zeros are kept: "007" != "7" != 7.
4. Numbers compare by value after rounding to 15 significant digits; -0 = 0.
5. Text matches a number only when the text is exactly the number's canonical form ("123" = 123; "123.0", "1,234", "1e3" stay text).
6. Dates compare by Excel serial (1904-system workbooks converted); a date equals the plain number with the same serial; a date never matches text.
7. TRUE/FALSE compare as booleans only.
8. Blank/whitespace-only cells are missing keys; error values and formulas without stored results are invalid keys.
9. Records whose key and transferred columns are all blank are ignored.
10. Duplicate keys on either side block the transfer.

Rule 5 is deliberately stricter than "parse everything as a number": it matches the common case (identifier typed
as text in one file and as a number in the other) without ever changing an identifier's meaning. Note that Excel's
own lookup functions would *not* match "123" to 123.

## 7. Format support and limitations

| Area | Status |
|---|---|
| `.xlsx` read/write | Implemented and tested (Open XML SDK). Parts re-serialised by the SDK: the edited worksheet, shared strings, styles, workbook part, `[Content_Types].xml`, workbook relationships. Everything else is copied unchanged (verified byte for byte on every run). |
| `.xlsm` | Same engine; `vbaProject.bin` preserved (tested with a real VBA project). Output keeps `.xlsm`. |
| `.xlsb` read | Own BIFF12 reader: cells (all 12 record types incl. RK numbers, rich strings, inline strings), shared strings, styles/number formats, sheets incl. hidden/very hidden/chart sheets, tables, merges, protection, shared/array/data-table formula ranges, 1904 dates, calc properties. Verified: byte-identical record round-trip on 7 Excel-made files; values identical to SheetJS on every cell of every fixture. |
| `.xlsb` write | Own BIFF12 editor: rewrites only the rows that change, recomputes row spans, appends shared strings and date styles, patches calc properties, removes the stale calc chain and binary index parts and their relationships/content types. Output verified structurally and by three parsers (ours, SheetJS, pyxlsb). **Not yet opened in Microsoft Excel** in this environment: do that once before relying on it. |
| `.xls`, `.csv`, `.ods`, templates (`.xltx`/`.xltm`), add-ins, encrypted/password-protected workbooks, Strict Open XML | Rejected with a specific message. |
| Formulas | Never written, never recalculated. Destination formulas inside mapped period columns are replaced by values (warned). Shared or array formulas that only partly overlap the cleared area block the transfer. |
| Dynamic arrays / spills in the destination period columns | Treated as array formulas (block when partly overlapping). Cell metadata records attached to cells are carried along. |
| Rich text | Flattened to plain text when copied; destination cells receive plain text. |
| Dates into cells without a date format | A cloned cell style with Excel's built-in short date (or date-time) format is applied; reported in the preview. |
| Very large sheets | Whole selected sheets are held in memory (snapshot + plan); see measurements below. |
| Excel Table features | Header, data and totals rows handled; calculated columns block; table style/banding untouched (table definition part is not modified). |
| Workbook/sheet protection | Protected destination sheets block (unprotect first); workbook structure protection is irrelevant and untouched. |

What it would take to support other formats: `.xls` (BIFF8) read could reuse NPOI (Apache 2.0) through a third
`IWorkbookSource`, but writing with preservation would need a BIFF8 editor similar to the BIFF12 one; `.csv`
needs no engine but has no sheets/tables/formatting (a one-day addition as a *source* only); `.ods` would need an
ODF engine (different XML vocabulary and package rules). Encrypted workbooks would need the MS-OFFCRYPTO
agile-encryption implementation plus a password prompt.

## 8. Safety properties

* Inputs are opened read-only with shared access and never written.
* The output path must differ from both inputs, must keep the destination's extension, must be in an existing
  writable folder, and existing files are only replaced after an explicit checkbox.
* Clearing and writing happen on a working copy; a failure or a failed verification deletes the copy and leaves no output.
* Every error message says what happened and what to do next (file locked, password protected, wrong extension, protected sheet, ...).

## 9. Tested versus assumed

**Implemented and covered by automated tests (116 tests, all passing):** single and composite keys; the
("AB","C")/("A","BC") collision; shuffled rows; whitespace/case variants; text-vs-number canonical rule; duplicate
keys on either side with record lists; missing keys blocking and the skip option; completely empty records;
formula results copied and missing results blocking; destination formulas replaced with warning; clearing of
unmatched destination rows while an unselected Q3 column and keys/descriptions/totals/formatting/table survive;
byte-identical preservation of a second sheet, pictures, comments, data validations and document properties;
hidden rows included/excluded; date formats applied; 1904-system dates; ordinary-range destinations; protection,
merged cells, partial shared formulas and calculated columns blocking; manual-calculation warning; xlsm VBA
preservation and output-extension rule; failed verification leaving no output; xlsb reader against reference
values; xlsb record round-trip; xlsb editing of a real workbook with VBA/chart/comments; transfers xlsx->xlsb,
xlsb->xlsx and into an Excel-made xlsb; file rejection messages; output path rules; and the complete wizard run
through headless UI tests.

**Verified manually in this environment:** xlsb outputs parsed by SheetJS 0.18.5 and pyxlsb 1.0.10 with
identical values; Windows x64 single-file publish from Linux.

**Assumed / not verifiable here:** Microsoft Excel opening the produced `.xlsb` files without a repair prompt
(structural checks and two independent parsers pass, but Excel is the only authority); appearance of the Avalonia
window on real Windows/macOS screens (rendered and screenshotted headlessly only); behaviour with workbooks far
larger than the measured sizes; antivirus/SmartScreen behaviour for your specific organisation.

## 10. Performance

Measured on the development container (4 vCPU, Linux, .NET 10, Release build) with generated 100,000-row
workbooks: planning 16 s, executing plus verification 18 s, peak memory 1.8 GB (details in
`docs/DEVELOPER-GUIDE.md`, section "Performance check"). Typical planning workbooks with a few thousand rows
complete in well under a second. The whole selected sheets are held in memory; streaming reads are the first
optimisation if much larger files become common.
