using ExcelTransfer.Core.Transfer;
using ExcelTransfer.Tests.Support;
using static ExcelTransfer.Tests.Support.WorkbookBuilder;

namespace ExcelTransfer.Tests;

public class PlannerTests
{
    private static (TempDir Dir, string Source, string Destination, string Output) Files()
    {
        var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"));
        var destination = CreateDestination(dir.File("destination.xlsx"));
        return (dir, source, destination, dir.File("output.xlsx"));
    }

    [Fact]
    public void Missing_mappings_are_errors()
    {
        var (dir, s, d, o) = Files();
        using (dir)
        {
            var plan = TransferPlanner.Plan(Spec(s, d, o, keys: Array.Empty<ColumnMapping>(), periods: Array.Empty<ColumnMapping>()));
            Assert.False(plan.CanTransfer);
            Assert.Contains(plan.Errors, e => e.Code == "keys-missing");
            Assert.Contains(plan.Errors, e => e.Code == "periods-missing");
        }
    }

    [Fact]
    public void Two_sources_cannot_target_the_same_destination_column()
    {
        var (dir, s, d, o) = Files();
        using (dir)
        {
            var plan = TransferPlanner.Plan(Spec(s, d, o, periods: new[] { new ColumnMapping(SrcQ1, DstQ1), new ColumnMapping(SrcQ2, DstQ1) }));
            Assert.Contains(plan.Errors, e => e.Code == "period-duplicate-destination");
        }
    }

    [Fact]
    public void Destination_key_columns_cannot_be_cleared()
    {
        var (dir, s, d, o) = Files();
        using (dir)
        {
            var plan = TransferPlanner.Plan(Spec(s, d, o, periods: new[] { new ColumnMapping(SrcQ1, DstRegion) }));
            Assert.Contains(plan.Errors, e => e.Code == "period-targets-key");
        }
    }

    [Fact]
    public void Columns_outside_the_data_are_errors()
    {
        var (dir, s, d, o) = Files();
        using (dir)
        {
            var plan = TransferPlanner.Plan(Spec(s, d, o, periods: new[] { new ColumnMapping(40, DstQ1) }));
            Assert.Contains(plan.Errors, e => e.Code == "period-column-missing");
        }
    }

    [Fact]
    public void Output_path_rules_are_enforced()
    {
        var (dir, s, d, _) = Files();
        using (dir)
        {
            Assert.Contains(TransferPlanner.Plan(Spec(s, d, d)).Errors, e => e.Code == "output-overwrites-input");
            Assert.Contains(TransferPlanner.Plan(Spec(s, d, s)).Errors, e => e.Code == "output-overwrites-input");
            Assert.Contains(TransferPlanner.Plan(Spec(s, d, dir.File("out.xlsb"))).Errors, e => e.Code == "output-extension");
            Assert.Contains(TransferPlanner.Plan(Spec(s, d, Path.Combine(dir.Path, "missing-folder", "out.xlsx"))).Errors, e => e.Code == "output-folder-missing");

            var existing = dir.File("existing.xlsx");
            File.WriteAllText(existing, "x");
            Assert.Contains(TransferPlanner.Plan(Spec(s, d, existing)).Errors, e => e.Code == "output-exists");
            var allowed = TransferPlanner.Plan(Spec(s, d, existing, options: new TransferOptions { OverwriteExistingOutput = true }));
            Assert.DoesNotContain(allowed.Errors, e => e.Code == "output-exists");
        }
    }

    [Fact]
    public void Suggested_output_name_keeps_the_destination_format_and_avoids_collisions()
    {
        using var dir = new TempDir();
        var dest = dir.File("Plan.xlsb");
        File.WriteAllText(dest, "x");
        var first = OutputPathRules.SuggestOutputPath(dest);
        Assert.Equal(dir.File("Plan - updated.xlsb"), first);
        File.WriteAllText(first, "x");
        Assert.Equal(dir.File("Plan - updated (2).xlsb"), OutputPathRules.SuggestOutputPath(dest));
    }

    [Fact]
    public void Plan_reports_matches_unmatched_and_cells_to_clear()
    {
        var (dir, s, d, o) = Files();
        using (dir)
        {
            var plan = TransferPlanner.Plan(Spec(s, d, o));
            Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
            Assert.Equal(4, plan.Matches.Count);
            Assert.Equal("P-999 | West", Assert.Single(plan.UnmatchedSource).Key.Display());
            Assert.Equal("P-004 | East", Assert.Single(plan.UnmatchedDestination).Key.Display());
            Assert.Equal(2, plan.ColumnsToClear.Count);
            Assert.Equal(5, plan.DestinationRowsInScope.Count);
            Assert.Equal(10, plan.TotalCellsToClear);
            Assert.Equal(10, plan.NonBlankCellsToClear);
            Assert.Equal(7, plan.ValuesToWrite); // 4 matches x 2 columns minus one blank source value
            Assert.Equal(1, plan.BlankSourceValues);
            Assert.Contains(plan.Issues, i => i.Code == "unmatched-source" && i.Severity == IssueSeverity.Warning);
            Assert.Contains(plan.Issues, i => i.Code == "unmatched-destination" && i.Severity == IssueSeverity.Info);
            Assert.NotEmpty(plan.MatchingRules);
        }
    }

    [Fact]
    public void Duplicate_keys_block_on_either_side_and_list_the_records()
    {
        using var dir = new TempDir();
        var dupSource = CreateSource(dir.File("dup-source.xlsx"), DefaultSource.Concat(new[] { new SourceRow("P-001", "north", 1.0, 2.0) }));
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var plan = TransferPlanner.Plan(Spec(dupSource, destination, dir.File("o.xlsx")));
        var issue = Assert.Single(plan.Errors.Where(e => e.Code == "duplicate-keys"));
        Assert.Equal(DataSide.Source, issue.Side);
        Assert.Contains(issue.Details, dtl => dtl.StartsWith("P-001 | North -> rows 3, 7"));
        Assert.Equal(new[] { 3, 7 }, issue.Rows.OrderBy(r => r));

        var source = CreateSource(dir.File("source.xlsx"));
        var dupDest = CreateDestination(dir.File("dup-destination.xlsx"), DefaultDestination.Concat(new[] { new DestRow("P-002", "NORTH", "dup", 0.0, 0.0, 0.0) }));
        var plan2 = TransferPlanner.Plan(Spec(source, dupDest, dir.File("o2.xlsx")));
        var issue2 = Assert.Single(plan2.Errors.Where(e => e.Code == "duplicate-keys"));
        Assert.Equal(DataSide.Destination, issue2.Side);
    }

    [Fact]
    public void Missing_keys_block_unless_explicitly_allowed()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"));
        var destination = CreateDestination(dir.File("destination.xlsx"), DefaultDestination.Concat(new[] { new DestRow(null, "North", "no code", 1.0, 1.0, 1.0) }));
        var blocked = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx")));
        var issue = Assert.Single(blocked.Errors.Where(e => e.Code == "missing-keys"));
        Assert.Equal(DataSide.Destination, issue.Side);
        Assert.Contains(issue.Details, d => d.Contains("Row 7") && d.Contains("blank"));

        var allowed = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), options: new TransferOptions { AllowMissingKeys = true }));
        Assert.True(allowed.CanTransfer);
        Assert.Contains(allowed.Warnings, w => w.Code == "missing-keys");
        Assert.Equal(1, allowed.DestinationRecordsWithMissingKeys);
    }

    [Fact]
    public void Completely_empty_records_are_ignored()
    {
        using var dir = new TempDir();
        var rows = DefaultSource.Take(2).Concat(new[] { new SourceRow(null, null, null, null) }).Concat(DefaultSource.Skip(2));
        var source = CreateSource(dir.File("source.xlsx"), rows);
        var destination = CreateDestination(dir.File("destination.xlsx"));
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), sourceSelection: SourceSelection(lastDataRow: 7)));
        Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
        Assert.Equal(1, plan.SourceEmptyRecordsSkipped);
        Assert.Equal(4, plan.Matches.Count);
    }

    [Fact]
    public void Text_and_number_keys_match_only_in_canonical_form()
    {
        using var dir = new TempDir();
        var source = CreateSource(dir.File("source.xlsx"), new[]
        {
            new SourceRow("1001", "N", 1.0, 2.0),
            new SourceRow("007", "N", 3.0, 4.0),
        });
        var destination = CreateDestination(dir.File("destination.xlsx"), Array.Empty<DestRow>(), customize: (_, ws) =>
        {
            ws.Cell(2, DstCode).Value = 1001; ws.Cell(2, DstRegion).Value = "n"; ws.Cell(2, DstDescription).Value = "numeric code";
            ws.Cell(3, DstCode).Value = 7; ws.Cell(3, DstRegion).Value = "N"; ws.Cell(3, DstDescription).Value = "seven";
        });
        var plan = TransferPlanner.Plan(Spec(source, destination, dir.File("o.xlsx"), destinationSelection: DestinationRangeSelection(lastDataRow: 3)));
        Assert.True(plan.CanTransfer, string.Join("|", plan.Errors.Select(e => e.Message)));
        var match = Assert.Single(plan.Matches);
        Assert.Equal(2, match.SourceRow);
        Assert.Equal(2, match.DestinationRow);
        Assert.Equal("007 | N", Assert.Single(plan.UnmatchedSource).Key.Display());
    }
}
