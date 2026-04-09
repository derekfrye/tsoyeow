using ExcelXmlQueryResults.App.Models;
using ExcelXmlWriter;

namespace ExcelXmlQueryResults.Desktop.ViewModels;

public sealed partial class MainWindowViewModel
{
    public async Task RunExportAsync()
    {
        if (IsBusy)
        {
            return;
        }

        GeneratedFiles.Clear();
        exportCancellation = new CancellationTokenSource();

        try
        {
            IsBusy = true;
            StatusText = "Starting export...";
            ProgressText = "Preparing request.";

            var settings = CreateSettingsSnapshot();
            await settingsStore.SaveAsync(settings);

            var progress = new Progress<ExportProgress>(HandleProgress);
            if (QueryFiles.Count > 0)
            {
                await RunBatchExportAsync(settings, progress, exportCancellation.Token);
            }
            else
            {
                await RunSingleExportAsync(settings, progress, exportCancellation.Token);
            }
        }
        catch (OperationCanceledException)
        {
            StatusText = "Export canceled.";
            ProgressText = "The active export was canceled.";
        }
        catch (Exception exception)
        {
            StatusText = "Export failed.";
            ProgressText = exception.Message;
        }
        finally
        {
            exportCancellation?.Dispose();
            exportCancellation = null;
            IsBusy = false;
        }
    }

    public void CancelExport()
    {
        exportCancellation?.Cancel();
    }

    async Task RunBatchExportAsync(
        AppSettings settings,
        IProgress<ExportProgress> progress,
        CancellationToken cancellationToken)
    {
        var batchResult = await workflowService.ExportQueryFilesAsync(
            QueryFiles.ToArray(),
            OutputPath,
            settings,
            progress,
            cancellationToken);

        foreach (var file in batchResult.GeneratedFiles)
        {
            GeneratedFiles.Add(file);
        }

        StatusText = $"Finished batch export. Generated {GeneratedFiles.Count} workbook files.";
    }

    async Task RunSingleExportAsync(
        AppSettings settings,
        IProgress<ExportProgress> progress,
        CancellationToken cancellationToken)
    {
        var result = await workflowService.ExportQueryAsync(
            QueryText,
            OutputPath,
            settings,
            progress,
            cancellationToken);

        foreach (var workbook in result.Workbooks)
        {
            GeneratedFiles.Add(workbook.Path);
        }

        StatusText = $"Finished export. Wrote {result.TotalRecordsRead:N0} rows across {result.Workbooks.Count} workbook file(s).";
    }

    void HandleProgress(ExportProgress progress)
    {
        switch (progress.Stage)
        {
            case ExportStage.QueryStarting:
                ProgressText = "Connecting and starting query execution.";
                break;
            case ExportStage.QueryRunning:
                ProgressText = progress.RowsPerSecond is { } rowsPerSecond
                    ? $"Streaming results at {rowsPerSecond:N2} rows/sec. Total rows: {progress.TotalRows:N0}."
                    : $"Streaming results. Total rows: {progress.TotalRows:N0}.";
                break;
            case ExportStage.SavingWorkbook:
                ProgressText = string.IsNullOrWhiteSpace(progress.OutputPath)
                    ? progress.Message ?? "Saving workbook."
                    : $"{progress.Message} {progress.OutputPath}";
                break;
            case ExportStage.Completed:
                ProgressText = $"Export complete. Total rows: {progress.TotalRows:N0}.";
                break;
        }
    }
}
