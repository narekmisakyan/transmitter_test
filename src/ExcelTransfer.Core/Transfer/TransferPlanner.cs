using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;
using ExcelTransfer.Core.Matching;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Transfer;

/// <summary>Validates a transfer and computes exactly what it would do, without touching any file.</summary>
public static class TransferPlanner
{
    public static TransferPlan Plan(TransferSpec spec, IProgress<ProgressReport>? progress = null, CancellationToken ct = default)
    {
        var issues = new List<ValidationIssue>();
        progress?.Report(new ProgressReport("Checking files", 0.02));

        var sourceCheck = WorkbookFileCheck.Check(spec.SourcePath);
        if (!sourceCheck.Ok) issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "source-file", "Source: " + sourceCheck.FullMessage));
        var destCheck = WorkbookFileCheck.Check(spec.DestinationPath);
        if (!destCheck.Ok) issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "destination-file", "Destination: " + destCheck.FullMessage));
        if (destCheck.Ok && destCheck.Format is WorkbookFormat destFormat)
            issues.AddRange(OutputPathRules.Validate(spec.OutputPath, spec.SourcePath, spec.DestinationPath, destFormat, spec.Options.OverwriteExistingOutput));
        if (issues.Any(i => i.IsError))
            return new TransferPlan { Spec = spec, Issues = issues, DestinationFormat = destCheck.Format };

        ct.ThrowIfCancellationRequested();
        progress?.Report(new ProgressReport("Reading source data", 0.1));
        LoadedDataSet source;
        try
        {
            source = SheetReader.Load(spec.SourcePath, spec.Source);
        }
        catch (TransferException ex)
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "source-read", "Source: " + ex.FullMessage));
            return new TransferPlan { Spec = spec, Issues = issues, DestinationFormat = destCheck.Format };
        }

        ct.ThrowIfCancellationRequested();
        progress?.Report(new ProgressReport("Reading destination data", 0.35));
        LoadedDataSet destination;
        try
        {
            destination = SheetReader.Load(spec.DestinationPath, spec.Destination);
        }
        catch (TransferException ex)
        {
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "destination-read", "Destination: " + ex.FullMessage));
            return new TransferPlan { Spec = spec, Issues = issues, Source = source, DestinationFormat = destCheck.Format };
        }

        ct.ThrowIfCancellationRequested();
        progress?.Report(new ProgressReport("Checking mappings", 0.6));
        ValidateMappings(spec, source, destination, issues);
        if (issues.Any(i => i.IsError))
            return new TransferPlan { Spec = spec, Issues = issues, Source = source, Destination = destination, DestinationFormat = destination.Format };

        progress?.Report(new ProgressReport("Matching records", 0.65));
        var plan = Match(spec, source, destination, issues, ct);
        progress?.Report(new ProgressReport("Checking destination structure", 0.9));
        CheckDestinationStructure(spec, destination, plan, issues);
        CheckSourceFreshness(source, issues);
        progress?.Report(new ProgressReport("Validation complete", 1));

        return new TransferPlan
        {
            Spec = spec,
            Issues = issues,
            Source = source,
            Destination = destination,
            DestinationFormat = destination.Format,
            Matches = plan.Matches,
            UnmatchedSource = plan.UnmatchedSource,
            UnmatchedDestination = plan.UnmatchedDestination,
            ColumnsToClear = plan.ColumnsToClear,
            DestinationRowsInScope = plan.RowsInScope,
            Writes = plan.Writes,
            SourceRecordsConsidered = plan.SourceConsidered,
            DestinationRecordsConsidered = plan.DestinationConsidered,
            SourceEmptyRecordsSkipped = plan.SourceEmpty,
            DestinationEmptyRecordsSkipped = plan.DestinationEmpty,
            SourceHiddenRowsSkipped = plan.SourceHidden,
            DestinationHiddenRowsSkipped = plan.DestinationHidden,
            SourceRecordsWithMissingKeys = plan.SourceMissingKeys,
            DestinationRecordsWithMissingKeys = plan.DestinationMissingKeys,
            BlankSourceValues = plan.BlankSourceValues,
            FormulaResultsUsed = plan.FormulaResultsUsed,
            DateValuesNeedingDateFormat = plan.DateValuesNeedingDateFormat,
        };
    }

    private static void ValidateMappings(TransferSpec spec, LoadedDataSet source, LoadedDataSet destination, List<ValidationIssue> issues)
    {
        if (spec.KeyMappings.Count == 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "keys-missing", "Select at least one key column pair."));
        if (spec.PeriodMappings.Count == 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.General, "periods-missing", "Select at least one period column pair to transfer."));

        foreach (var m in spec.KeyMappings)
        {
            if (source.FindColumn(m.SourceColumnIndex) is null)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "key-column-missing", $"Key column {CellRef.ColumnLetter(m.SourceColumnIndex)} is not part of the source data."));
            if (destination.FindColumn(m.DestinationColumnIndex) is null)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "key-column-missing", $"Key column {CellRef.ColumnLetter(m.DestinationColumnIndex)} is not part of the destination data."));
        }
        foreach (var m in spec.PeriodMappings)
        {
            if (source.FindColumn(m.SourceColumnIndex) is null)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "period-column-missing", $"Period column {CellRef.ColumnLetter(m.SourceColumnIndex)} is not part of the source data."));
            if (destination.FindColumn(m.DestinationColumnIndex) is null)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "period-column-missing", $"Period column {CellRef.ColumnLetter(m.DestinationColumnIndex)} is not part of the destination data."));
        }

        var dupKeySrc = spec.KeyMappings.GroupBy(m => m.SourceColumnIndex).Where(g => g.Count() > 1).Select(g => CellRef.ColumnLetter(g.Key)).ToList();
        if (dupKeySrc.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "key-duplicate-source", $"Each source key column may be used once. Repeated: {string.Join(", ", dupKeySrc)}."));
        var dupKeyDst = spec.KeyMappings.GroupBy(m => m.DestinationColumnIndex).Where(g => g.Count() > 1).Select(g => CellRef.ColumnLetter(g.Key)).ToList();
        if (dupKeyDst.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "key-duplicate-destination", $"Each destination key column may be used once. Repeated: {string.Join(", ", dupKeyDst)}."));

        var dupPeriodDst = spec.PeriodMappings.GroupBy(m => m.DestinationColumnIndex).Where(g => g.Count() > 1).Select(g => CellRef.ColumnLetter(g.Key)).ToList();
        if (dupPeriodDst.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "period-duplicate-destination", $"Two source columns cannot write to the same destination column. Repeated destination: {string.Join(", ", dupPeriodDst)}."));

        var keyDest = spec.KeyMappings.Select(m => m.DestinationColumnIndex).ToHashSet();
        var periodIntoKey = spec.PeriodMappings.Where(m => keyDest.Contains(m.DestinationColumnIndex)).Select(m => CellRef.ColumnLetter(m.DestinationColumnIndex)).Distinct().ToList();
        if (periodIntoKey.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "period-targets-key", $"Destination key columns cannot be cleared or overwritten: {string.Join(", ", periodIntoKey)}."));

        var keySrc = spec.KeyMappings.Select(m => m.SourceColumnIndex).ToHashSet();
        var keyAsPeriod = spec.PeriodMappings.Where(m => keySrc.Contains(m.SourceColumnIndex)).Select(m => CellRef.ColumnLetter(m.SourceColumnIndex)).Distinct().ToList();
        if (keyAsPeriod.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Source, "key-as-period", $"Source key column(s) {string.Join(", ", keyAsPeriod)} are also mapped as period columns; their values will be copied."));

        var dupPeriodSrc = spec.PeriodMappings.GroupBy(m => m.SourceColumnIndex).Where(g => g.Count() > 1).Select(g => CellRef.ColumnLetter(g.Key)).ToList();
        if (dupPeriodSrc.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Info, DataSide.Source, "period-duplicate-source", $"Source column(s) {string.Join(", ", dupPeriodSrc)} are mapped to more than one destination column; the same values will be written to each."));
    }

    private sealed class MatchWork
    {
        public List<MatchedPair> Matches = new();
        public List<UnmatchedRecord> UnmatchedSource = new();
        public List<UnmatchedRecord> UnmatchedDestination = new();
        public List<ClearColumnPlan> ColumnsToClear = new();
        public List<int> RowsInScope = new();
        public List<CellWrite> Writes = new();
        public int SourceConsidered, DestinationConsidered, SourceEmpty, DestinationEmpty, SourceHidden, DestinationHidden, SourceMissingKeys, DestinationMissingKeys, BlankSourceValues, FormulaResultsUsed, DateValuesNeedingDateFormat;
    }

    private sealed class KeyIndex
    {
        public readonly Dictionary<CompositeKey, List<int>> Rows = new();
        public readonly Dictionary<int, CompositeKey> KeyByRow = new();
        public readonly List<(int Row, string Reason)> MissingKeys = new();
        public int Considered, Empty, Hidden;
    }

    private static KeyIndex BuildIndex(LoadedDataSet data, IReadOnlyList<int> keyColumns, IReadOnlyList<int> valueColumns, bool includeHidden)
    {
        var index = new KeyIndex();
        var keyCells = new CellValue[keyColumns.Count];
        foreach (var record in data.Records)
        {
            if (!includeHidden && record.Hidden)
            {
                index.Hidden++;
                continue;
            }
            var allBlank = true;
            for (var i = 0; i < keyColumns.Count; i++)
            {
                keyCells[i] = record[keyColumns[i]];
                if (!keyCells[i].IsBlank) allBlank = false;
            }
            if (allBlank)
            {
                foreach (var c in valueColumns)
                {
                    if (!record[c].IsBlank) { allBlank = false; break; }
                }
            }
            if (allBlank)
            {
                index.Empty++;
                continue;
            }
            index.Considered++;
            var result = KeyNormalizer.BuildKey(keyCells);
            if (!result.Ok)
            {
                var reasons = result.Problems.Select(p =>
                {
                    var col = data.FindColumn(keyColumns[p.ComponentIndex]);
                    var label = col?.Label ?? CellRef.ColumnLetter(keyColumns[p.ComponentIndex]);
                    return p.Problem switch
                    {
                        KeyProblem.Missing => $"{label} is blank",
                        KeyProblem.ErrorValue => $"{label} contains an error value",
                        KeyProblem.FormulaWithoutResult => $"{label} has a formula without a stored result",
                        _ => label,
                    };
                });
                index.MissingKeys.Add((record.RowNumber, string.Join("; ", reasons)));
                continue;
            }
            var key = result.Key!;
            index.KeyByRow[record.RowNumber] = key;
            if (!index.Rows.TryGetValue(key, out var rows))
            {
                rows = new List<int>(1);
                index.Rows[key] = rows;
            }
            rows.Add(record.RowNumber);
        }
        return index;
    }

    private static MatchWork Match(TransferSpec spec, LoadedDataSet source, LoadedDataSet destination, List<ValidationIssue> issues, CancellationToken ct)
    {
        var work = new MatchWork();
        var srcKeyCols = spec.KeyMappings.Select(m => m.SourceColumnIndex).ToList();
        var dstKeyCols = spec.KeyMappings.Select(m => m.DestinationColumnIndex).ToList();
        var srcPeriodCols = spec.PeriodMappings.Select(m => m.SourceColumnIndex).Distinct().ToList();
        var dstPeriodCols = spec.PeriodMappings.Select(m => m.DestinationColumnIndex).ToList();

        var srcIndex = BuildIndex(source, srcKeyCols, srcPeriodCols, spec.Source.IncludeHiddenRows);
        ct.ThrowIfCancellationRequested();
        var dstIndex = BuildIndex(destination, dstKeyCols, dstPeriodCols, spec.Destination.IncludeHiddenRows);
        ct.ThrowIfCancellationRequested();

        work.SourceConsidered = srcIndex.Considered;
        work.DestinationConsidered = dstIndex.Considered;
        work.SourceEmpty = srcIndex.Empty;
        work.DestinationEmpty = dstIndex.Empty;
        work.SourceHidden = srcIndex.Hidden;
        work.DestinationHidden = dstIndex.Hidden;
        work.SourceMissingKeys = srcIndex.MissingKeys.Count;
        work.DestinationMissingKeys = dstIndex.MissingKeys.Count;

        ReportMissingKeys(spec, DataSide.Source, srcIndex, issues, "skipped");
        ReportMissingKeys(spec, DataSide.Destination, dstIndex, issues, "left blank in the transferred columns");
        ReportDuplicates(DataSide.Source, srcIndex, issues);
        ReportDuplicates(DataSide.Destination, dstIndex, issues);

        // Matching: only keys that are unique on both sides can be matched.
        var dstByKey = dstIndex.Rows;
        var matchedSourceRows = new HashSet<int>();
        foreach (var record in destination.Records)
        {
            if (!dstIndex.KeyByRow.TryGetValue(record.RowNumber, out var key)) continue;
            if (dstByKey[key].Count > 1) continue;
            if (!srcIndex.Rows.TryGetValue(key, out var srcRows))
            {
                work.UnmatchedDestination.Add(new UnmatchedRecord(record.RowNumber, key));
                continue;
            }
            if (srcRows.Count > 1) continue;
            work.Matches.Add(new MatchedPair(srcRows[0], record.RowNumber, key));
            matchedSourceRows.Add(srcRows[0]);
        }
        foreach (var record in source.Records)
        {
            if (!srcIndex.KeyByRow.TryGetValue(record.RowNumber, out var key)) continue;
            if (srcIndex.Rows[key].Count > 1) continue;
            if (!matchedSourceRows.Contains(record.RowNumber))
                work.UnmatchedSource.Add(new UnmatchedRecord(record.RowNumber, key));
        }
        if (work.UnmatchedSource.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Source, "unmatched-source",
                $"{work.UnmatchedSource.Count} source record(s) have no matching destination record and will be skipped (rows are never appended).",
                work.UnmatchedSource.Select(u => $"Row {u.Row}: {u.Key.Display()}"), work.UnmatchedSource.Select(u => u.Row)));
        if (work.UnmatchedDestination.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Info, DataSide.Destination, "unmatched-destination",
                $"{work.UnmatchedDestination.Count} destination record(s) have no matching source record; their transferred columns will be blank.",
                work.UnmatchedDestination.Select(u => $"Row {u.Row}: {u.Key.Display()}"), work.UnmatchedDestination.Select(u => u.Row)));

        // Scope for clearing: every data row (hidden rows unless excluded).
        foreach (var record in destination.Records)
        {
            if (!spec.Destination.IncludeHiddenRows && record.Hidden) continue;
            work.RowsInScope.Add(record.RowNumber);
        }
        var recordsByRow = destination.Records.ToDictionary(r => r.RowNumber);
        foreach (var col in dstPeriodCols)
        {
            var column = destination.FindColumn(col)!;
            var nonBlank = 0;
            var formulas = 0;
            foreach (var row in work.RowsInScope)
            {
                var v = recordsByRow[row][col];
                if (!v.IsBlank) nonBlank++;
                if (v.FromFormula) formulas++;
            }
            work.ColumnsToClear.Add(new ClearColumnPlan(column, work.RowsInScope.Count, nonBlank, formulas));
        }

        // Writes.
        var srcByRow = source.Records.ToDictionary(r => r.RowNumber);
        var formulaWithoutResult = new List<string>();
        var errorValues = new List<string>();
        foreach (var pair in work.Matches)
        {
            var srcRecord = srcByRow[pair.SourceRow];
            var dstRecord = recordsByRow[pair.DestinationRow];
            foreach (var m in spec.PeriodMappings)
            {
                var v = srcRecord[m.SourceColumnIndex];
                if (v.IsBlank)
                {
                    work.BlankSourceValues++;
                    continue;
                }
                if (v.Kind == CellKind.FormulaWithoutResult)
                {
                    formulaWithoutResult.Add($"Row {pair.SourceRow}, {source.FindColumn(m.SourceColumnIndex)!.Label}: ={v.Text}");
                    continue;
                }
                if (v.Kind == CellKind.Error)
                    errorValues.Add($"Row {pair.SourceRow}, {source.FindColumn(m.SourceColumnIndex)!.Label}: {v.Text}");
                if (v.FromFormula) work.FormulaResultsUsed++;
                if (v.Kind == CellKind.Date && !dstRecord[m.DestinationColumnIndex].DateStyled) work.DateValuesNeedingDateFormat++;
                work.Writes.Add(new CellWrite(pair.DestinationRow, m.DestinationColumnIndex, v));
            }
        }
        if (formulaWithoutResult.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Source, "formula-without-result",
                $"{formulaWithoutResult.Count} source cell(s) contain formulas whose results are not stored in the file, so there is no value to copy. Open the source workbook in Excel, let it calculate, save it, and try again.",
                formulaWithoutResult));
        if (errorValues.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Source, "error-values",
                $"{errorValues.Count} source value(s) are Excel error values; they will be copied as error values.", errorValues));
        if (work.FormulaResultsUsed > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Info, DataSide.Source, "formula-results",
                $"{work.FormulaResultsUsed} value(s) come from formulas in the source; the results stored at the source's last save are copied (formulas are not recalculated)."));
        if (work.DateValuesNeedingDateFormat > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Info, DataSide.Destination, "date-format-applied",
                $"{work.DateValuesNeedingDateFormat} date value(s) will be written into cells that have no date number format; a short date format will be applied to those cells so they display as dates."));
        if (work.SourceHidden > 0 || work.DestinationHidden > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Info, DataSide.General, "hidden-rows-skipped",
                $"Hidden rows excluded: {work.SourceHidden} in the source, {work.DestinationHidden} in the destination."));
        return work;
    }

    private static void ReportMissingKeys(TransferSpec spec, DataSide side, KeyIndex index, List<ValidationIssue> issues, string consequence)
    {
        if (index.MissingKeys.Count == 0) return;
        var details = index.MissingKeys.Select(m => $"Row {m.Row}: {m.Reason}");
        var rows = index.MissingKeys.Select(m => m.Row);
        var sideName = side == DataSide.Source ? "source" : "destination";
        if (spec.Options.AllowMissingKeys)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, side, "missing-keys",
                $"{index.MissingKeys.Count} {sideName} record(s) have missing or invalid key values and will be {consequence}.", details, rows));
        else
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, side, "missing-keys",
                $"{index.MissingKeys.Count} {sideName} record(s) have missing or invalid key values. Fix them, narrow the data range, or allow records with missing keys to be skipped.", details, rows));
    }

    private static void ReportDuplicates(DataSide side, KeyIndex index, List<ValidationIssue> issues)
    {
        var duplicates = index.Rows.Where(kv => kv.Value.Count > 1).ToList();
        if (duplicates.Count == 0) return;
        var sideName = side == DataSide.Source ? "source" : "destination";
        var details = duplicates.Select(kv => $"{kv.Key.Display()} -> rows {string.Join(", ", kv.Value)}");
        var rows = duplicates.SelectMany(kv => kv.Value);
        issues.Add(ValidationIssue.Create(IssueSeverity.Error, side, "duplicate-keys",
            $"{duplicates.Count} key(s) occur more than once in the {sideName} data ({duplicates.Sum(kv => kv.Value.Count)} records). Each key must identify exactly one record; the transfer never picks the first or last duplicate or adds values together.",
            details, rows));
    }

    private static void CheckDestinationStructure(TransferSpec spec, LoadedDataSet destination, MatchWork work, List<ValidationIssue> issues)
    {
        if (destination.Metadata.Protected)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "sheet-protected",
                $"Destination sheet \"{destination.SheetName}\" is protected. Unprotect it in Excel (Review > Unprotect Sheet), save, and try again."));

        var periodCols = spec.PeriodMappings.Select(m => m.DestinationColumnIndex).ToHashSet();
        var rowsInScope = work.RowsInScope.ToHashSet();
        bool RangeFullyTargeted(CellRange range)
        {
            for (var c = range.FirstColumn; c <= range.LastColumn; c++)
                if (!periodCols.Contains(c)) return false;
            for (var r = range.FirstRow; r <= range.LastRow; r++)
                if (!rowsInScope.Contains(r)) return false;
            return true;
        }
        bool RangeTouchesTargets(CellRange range)
        {
            for (var c = range.FirstColumn; c <= range.LastColumn; c++)
            {
                if (!periodCols.Contains(c)) continue;
                for (var r = range.FirstRow; r <= range.LastRow; r++)
                    if (rowsInScope.Contains(r)) return true;
            }
            return false;
        }

        var mergedHits = destination.Metadata.MergedRanges.Where(RangeTouchesTargets).Select(r => r.ToString()).ToList();
        if (mergedHits.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "merged-cells",
                $"{mergedHits.Count} merged cell range(s) overlap the destination period columns. Unmerge them in Excel first, so every record has its own cells.", mergedHits));

        var sharedHits = destination.Metadata.SharedFormulas.Where(g => RangeTouchesTargets(g.Range) && !RangeFullyTargeted(g.Range)).Select(g => g.Range.ToString()).Distinct().ToList();
        if (sharedHits.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "shared-formula-partial",
                $"{sharedHits.Count} shared formula group(s) extend beyond the cells that would be cleared (for example a formula filled across several columns). Clearing part of such a group would break the remaining formulas. Convert these formulas to values in Excel, or map every column of the group.",
                sharedHits));

        var arrayHits = destination.Metadata.ArrayFormulas.Where(r => RangeTouchesTargets(r) && !RangeFullyTargeted(r)).Select(r => r.ToString()).ToList();
        if (arrayHits.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "array-formula-partial",
                $"{arrayHits.Count} array formula range(s) partly overlap the destination period columns. An array formula cannot be partly cleared; remove or move it in Excel first.", arrayHits));

        var tableHits = destination.Metadata.DataTableFormulas.Where(RangeTouchesTargets).Select(r => r.ToString()).ToList();
        if (tableHits.Count > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "data-table-formula",
                $"{tableHits.Count} what-if data table range(s) overlap the destination period columns. Remove the data table in Excel first.", tableHits));

        if (destination.Table is not null)
        {
            var calculated = new List<string>();
            foreach (var m in spec.PeriodMappings)
            {
                var idx = m.DestinationColumnIndex - destination.Table.Range.FirstColumn;
                if (idx >= 0 && idx < destination.Table.Columns.Count && destination.Table.Columns[idx].HasCalculatedColumnFormula)
                    calculated.Add(destination.FindColumn(m.DestinationColumnIndex)!.Label);
            }
            if (calculated.Count > 0)
                issues.Add(ValidationIssue.Create(IssueSeverity.Error, DataSide.Destination, "calculated-column",
                    $"Destination table column(s) {string.Join(", ", calculated)} are calculated columns (Excel fills them with a formula). Writing values into them is not supported; remove the column formula in Excel or map a different column."));
        }

        var formulaCells = work.ColumnsToClear.Sum(c => c.FormulaCount);
        if (formulaCells > 0)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Destination, "formulas-replaced",
                $"{formulaCells} destination cell(s) in the mapped period columns contain formulas. They will be cleared and replaced by values."));
    }

    private static void CheckSourceFreshness(LoadedDataSet source, List<ValidationIssue> issues)
    {
        if (source.ManualCalculation)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Source, "manual-calculation",
                "The source workbook is set to manual calculation; stored formula results may be out of date. Press F9 in Excel and save the source before transferring if in doubt."));
        else if (source.FullCalcOnLoad)
            issues.Add(ValidationIssue.Create(IssueSeverity.Warning, DataSide.Source, "pending-recalculation",
                "The source workbook is flagged for recalculation on next open, so stored formula results may be out of date. Open and save it in Excel before transferring if in doubt."));
    }
}
