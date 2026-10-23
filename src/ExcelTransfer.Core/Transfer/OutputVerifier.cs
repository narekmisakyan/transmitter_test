using System.IO.Compression;
using ExcelTransfer.Core.Cells;
using ExcelTransfer.Core.Data;

namespace ExcelTransfer.Core.Transfer;

/// <summary>
/// Reads a saved output workbook back and checks it against the plan: every planned clear and write happened,
/// every other cell of the data set is unchanged, and every package part outside the edited ones is byte-identical
/// to the original destination workbook.
/// </summary>
public static class OutputVerifier
{
    public static void Verify(TransferPlan plan, string outputPath, IReadOnlySet<string> changedParts)
    {
        var before = plan.Destination ?? throw new TransferException("The plan has no destination data.");
        LoadedDataSet after;
        try
        {
            after = SheetReader.Load(outputPath, before.Selection);
        }
        catch (TransferException ex)
        {
            throw new TransferException("The saved workbook could not be read back, so it was discarded. " + ex.Message, ex.Hint, ex);
        }

        var problems = new List<string>();
        if (after.Records.Count != before.Records.Count || after.FirstColumn != before.FirstColumn || after.LastColumn != before.LastColumn)
            problems.Add($"The data range changed ({before.Describe()} vs {after.Describe()}).");
        for (var i = 0; i < Math.Min(before.Columns.Count, after.Columns.Count); i++)
            if (!string.Equals(before.Columns[i].Header, after.Columns[i].Header, StringComparison.Ordinal))
                problems.Add($"Header of column {before.Columns[i].Letter} changed from \"{before.Columns[i].Header}\" to \"{after.Columns[i].Header}\".");

        if (problems.Count == 0)
        {
            var writes = new Dictionary<(int, int), CellValue>();
            foreach (var w in plan.Writes) writes[(w.Row, w.Column)] = w.Value;
            var clearColumns = plan.ColumnsToClear.Select(c => c.Column.ColumnIndex).ToHashSet();
            var scope = plan.DestinationRowsInScope.ToHashSet();
            var afterByRow = after.Records.ToDictionary(r => r.RowNumber);
            foreach (var record in before.Records)
            {
                if (!afterByRow.TryGetValue(record.RowNumber, out var afterRecord))
                {
                    problems.Add($"Row {record.RowNumber} is missing in the output.");
                    continue;
                }
                foreach (var column in before.Columns)
                {
                    var c = column.ColumnIndex;
                    CellValue expected;
                    if (clearColumns.Contains(c) && scope.Contains(record.RowNumber))
                        expected = writes.TryGetValue((record.RowNumber, c), out var w) ? w : CellValue.Blank;
                    else
                        expected = record[c];
                    var actual = afterRecord[c];
                    if (!SameValue(expected, actual))
                        problems.Add($"{CellRef.ToA1(record.RowNumber, c)}: expected \"{expected.Display()}\" but found \"{actual.Display()}\".");
                    if (problems.Count >= 20) break;
                }
                if (problems.Count >= 20) break;
            }
        }
        if (problems.Count > 0)
            throw new TransferException("Verification of the saved workbook failed, so it was discarded: " + string.Join(" ", problems.Take(5)) + (problems.Count > 5 ? $" ({problems.Count} problems in total)" : string.Empty));

        VerifyUntouchedParts(before.WorkbookPath, outputPath, changedParts);
    }

    /// <summary>Numbers and dates compare by value; the display format of the destination cell may legitimately differ.</summary>
    internal static bool SameValue(CellValue expected, CellValue actual)
    {
        static bool IsNumeric(CellKind k) => k is CellKind.Number or CellKind.Date;
        if (IsNumeric(expected.Kind) && IsNumeric(actual.Kind))
            return expected.Number.Equals(actual.Number) || Math.Abs(expected.Number - actual.Number) <= 1e-9 * Math.Max(1, Math.Abs(expected.Number));
        if (expected.Kind == CellKind.FormulaWithoutResult)
            return actual.Kind == CellKind.FormulaWithoutResult || actual.FromFormula;
        return expected.ValueEquals(actual);
    }

    /// <summary>Every zip entry of the original that the editor did not declare as changed must be byte-identical in the output.</summary>
    public static void VerifyUntouchedParts(string originalPath, string outputPath, IReadOnlySet<string> changedParts)
    {
        var allowed = new HashSet<string>(changedParts.Select(Normalize), StringComparer.OrdinalIgnoreCase)
        {
            "[Content_Types].xml",
            "xl/_rels/workbook.xml.rels",
            "xl/_rels/workbook.bin.rels",
        };
        using var original = ZipFile.OpenRead(originalPath);
        using var output = ZipFile.OpenRead(outputPath);
        var outputEntries = output.Entries.Where(e => !e.FullName.EndsWith('/')).ToDictionary(e => Normalize(e.FullName), e => e, StringComparer.OrdinalIgnoreCase);
        var problems = new List<string>();
        foreach (var entry in original.Entries)
        {
            if (entry.FullName.EndsWith('/')) continue;
            var name = Normalize(entry.FullName);
            if (allowed.Contains(name)) continue;
            if (!outputEntries.TryGetValue(name, out var other))
            {
                problems.Add($"part \"{name}\" is missing");
                continue;
            }
            if (!SameBytes(entry, other)) problems.Add($"part \"{name}\" changed");
        }
        foreach (var name in outputEntries.Keys)
        {
            if (allowed.Contains(name)) continue;
            if (original.GetEntry(name) is null && !original.Entries.Any(e => string.Equals(Normalize(e.FullName), name, StringComparison.OrdinalIgnoreCase)))
                problems.Add($"unexpected new part \"{name}\"");
        }
        if (problems.Count > 0)
            throw new TransferException("Verification of the saved workbook failed, so it was discarded: parts outside the edited worksheet differ from the original (" + string.Join("; ", problems.Take(5)) + ").");
    }

    private static string Normalize(string name) => name.Replace('\\', '/').TrimStart('/');

    private static bool SameBytes(ZipArchiveEntry a, ZipArchiveEntry b)
    {
        if (a.Length != b.Length) return false;
        using var sa = a.Open();
        using var sb = b.Open();
        var bufA = new byte[81920];
        var bufB = new byte[81920];
        while (true)
        {
            var ra = ReadFull(sa, bufA);
            var rb = ReadFull(sb, bufB);
            if (ra != rb) return false;
            if (ra == 0) return true;
            if (!bufA.AsSpan(0, ra).SequenceEqual(bufB.AsSpan(0, rb))) return false;
        }
    }

    private static int ReadFull(Stream s, byte[] buffer)
    {
        var total = 0;
        while (total < buffer.Length)
        {
            var n = s.Read(buffer, total, buffer.Length - total);
            if (n <= 0) break;
            total += n;
        }
        return total;
    }
}
