using System.Diagnostics;
using ExcelTransfer.Core.Workbooks;

namespace ExcelTransfer.Core.Transfer;

public sealed record TransferResult(
    string OutputPath,
    WorkbookFormat Format,
    int RecordsMatched,
    int CellsCleared,
    int ValuesWritten,
    int DateStylesApplied,
    TimeSpan Duration,
    IReadOnlyList<string> Notes);

/// <summary>
/// Runs a validated <see cref="TransferPlan"/>: the destination is copied to a temporary file next to the output,
/// the copy is edited and saved, the saved copy is read back and verified, and only then is it moved to the
/// output path. The source and destination workbooks are never modified.
/// </summary>
public static class TransferExecutor
{
    public static TransferResult Execute(TransferPlan plan, IProgress<ProgressReport>? progress = null, CancellationToken ct = default)
    {
        if (!plan.CanTransfer || plan.Destination is null || plan.Source is null)
            throw new TransferException("The transfer cannot run until every validation error is resolved.");

        var stopwatch = Stopwatch.StartNew();
        var destination = plan.Destination;
        var format = destination.Format;
        var outputPath = Path.GetFullPath(plan.Spec.OutputPath);
        var outputDir = Path.GetDirectoryName(outputPath) ?? throw new TransferException("The output path has no folder.");

        var recheck = OutputPathRules.Validate(outputPath, plan.Spec.SourcePath, plan.Spec.DestinationPath, format, plan.Spec.Options.OverwriteExistingOutput);
        var blocking = recheck.FirstOrDefault(i => i.IsError);
        if (blocking is not null) throw new TransferException(blocking.Message);

        var tempPath = Path.Combine(outputDir, $"~{Path.GetFileNameWithoutExtension(outputPath)}.{Guid.NewGuid():N}.tmp{format.Extension()}");
        try
        {
            progress?.Report(new ProgressReport("Creating working copy", 0.05));
            SheetEditStats stats;
            IReadOnlySet<string> changedParts;
            using (var editor = WorkbookEditors.CreateWorkingCopy(plan.Spec.DestinationPath, tempPath, format))
            {
                ct.ThrowIfCancellationRequested();
                stats = editor.EditSheet(destination.SheetName, plan.CellsToClear().ToList(), plan.Writes, progress, ct);
                progress?.Report(new ProgressReport("Saving working copy", 0.65));
                changedParts = editor.Finish();
            }

            ct.ThrowIfCancellationRequested();
            progress?.Report(new ProgressReport("Verifying the saved workbook", 0.8));
            OutputVerifier.Verify(plan, tempPath, changedParts);

            progress?.Report(new ProgressReport("Moving into place", 0.95));
            MoveIntoPlace(tempPath, outputPath, plan.Spec.Options.OverwriteExistingOutput);
            progress?.Report(new ProgressReport("Done", 1));

            var notes = new List<string>
            {
                "The source and destination workbooks were not modified.",
                "Excel will recalculate formulas when the output is opened and will ask to save on close.",
            };
            if (stats.DateStylesApplied > 0)
                notes.Add($"{stats.DateStylesApplied} cell(s) received a date number format so copied dates display as dates.");
            return new TransferResult(outputPath, format, plan.Matches.Count, stats.CellsCleared, stats.ValuesWritten, stats.DateStylesApplied, stopwatch.Elapsed, notes);
        }
        catch (TransferException)
        {
            TryDelete(tempPath);
            throw;
        }
        catch (OperationCanceledException)
        {
            TryDelete(tempPath);
            throw;
        }
        catch (Exception ex)
        {
            TryDelete(tempPath);
            throw new TransferException($"The transfer failed and no output was written: {ex.Message}", null, ex);
        }
    }

    private static void MoveIntoPlace(string tempPath, string outputPath, bool overwrite)
    {
        try
        {
            if (File.Exists(outputPath))
            {
                if (!overwrite) throw new TransferException($"The output file \"{Path.GetFileName(outputPath)}\" appeared while the transfer was running. Choose another name.");
                File.Move(tempPath, outputPath, overwrite: true);
            }
            else
            {
                File.Move(tempPath, outputPath);
            }
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new TransferException($"The output file \"{Path.GetFileName(outputPath)}\" could not be written: access denied.", "If the file is open in Excel, close it. Otherwise check the folder permissions.", ex);
        }
        catch (IOException ex)
        {
            throw new TransferException($"The output file \"{Path.GetFileName(outputPath)}\" could not be written: {ex.Message}", "If the file is open in Excel or another program, close it and try again.", ex);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
