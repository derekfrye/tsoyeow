using System.Collections.ObjectModel;
using System.Globalization;
using ExcelXmlQueryResults.App.Models;
using ExcelXmlQueryResults.App.Services;
using ExcelXmlQueryResults.Desktop.Infrastructure;
using ExcelXmlWriter;

namespace ExcelXmlQueryResults.Desktop.ViewModels;

public sealed class MainWindowViewModel : ViewModelBase, IDisposable
{
    readonly AppSettingsStore settingsStore;
    readonly ExportWorkflowService workflowService;

    CancellationTokenSource? exportCancellation;
    bool initialized;
    bool isBusy;
    bool useIntegratedSecurity = true;
    string server = string.Empty;
    string database = string.Empty;
    string username = string.Empty;
    string password = string.Empty;
    string queryTimeoutSeconds = "0";
    string maxRowsPerWorksheet = "150000";
    string maxWorkbookSizeGiB = "1";
    string maxResultSetsPerWorkbook = "99";
    string boundaryColumns = string.Empty;
    string resultSetNames = string.Empty;
    string queryText = string.Empty;
    string outputPath = string.Empty;
    string statusText = "Ready.";
    string progressText = "No export is running.";

    public MainWindowViewModel(AppSettingsStore settingsStore, ExportWorkflowService workflowService)
    {
        this.settingsStore = settingsStore;
        this.workflowService = workflowService;
    }

    public ObservableCollection<string> QueryFiles { get; } = [];

    public ObservableCollection<string> GeneratedFiles { get; } = [];

    public bool IsBusy
    {
        get => isBusy;
        private set
        {
            if (SetProperty(ref isBusy, value))
            {
                RaisePropertyChanged(nameof(IsSingleQueryMode));
            }
        }
    }

    public bool IsSingleQueryMode => QueryFiles.Count == 0 && !IsBusy;

    public bool UseIntegratedSecurity
    {
        get => useIntegratedSecurity;
        set => SetProperty(ref useIntegratedSecurity, value);
    }

    public string Server
    {
        get => server;
        set => SetProperty(ref server, value);
    }

    public string Database
    {
        get => database;
        set => SetProperty(ref database, value);
    }

    public string Username
    {
        get => username;
        set => SetProperty(ref username, value);
    }

    public string Password
    {
        get => password;
        set => SetProperty(ref password, value);
    }

    public string QueryTimeoutSeconds
    {
        get => queryTimeoutSeconds;
        set => SetProperty(ref queryTimeoutSeconds, value);
    }

    public string MaxRowsPerWorksheet
    {
        get => maxRowsPerWorksheet;
        set => SetProperty(ref maxRowsPerWorksheet, value);
    }

    public string MaxWorkbookSizeGiB
    {
        get => maxWorkbookSizeGiB;
        set => SetProperty(ref maxWorkbookSizeGiB, value);
    }

    public string MaxResultSetsPerWorkbook
    {
        get => maxResultSetsPerWorkbook;
        set => SetProperty(ref maxResultSetsPerWorkbook, value);
    }

    public string BoundaryColumns
    {
        get => boundaryColumns;
        set => SetProperty(ref boundaryColumns, value);
    }

    public string ResultSetNames
    {
        get => resultSetNames;
        set => SetProperty(ref resultSetNames, value);
    }

    public string QueryText
    {
        get => queryText;
        set => SetProperty(ref queryText, value);
    }

    public string OutputPath
    {
        get => outputPath;
        set => SetProperty(ref outputPath, value);
    }

    public string StatusText
    {
        get => statusText;
        private set => SetProperty(ref statusText, value);
    }

    public string ProgressText
    {
        get => progressText;
        private set => SetProperty(ref progressText, value);
    }

    public async Task InitializeAsync()
    {
        if (initialized)
        {
            return;
        }

        initialized = true;
        var settings = await settingsStore.LoadAsync();
        ApplySettings(settings);
        StatusText = "Settings loaded.";
    }

    public async Task SaveAsync()
    {
        await settingsStore.SaveAsync(CreateSettingsSnapshot());
    }

    public void SetOutputPath(string path)
    {
        OutputPath = path;
    }

    public void SetQueryFiles(IEnumerable<string> files)
    {
        QueryFiles.Clear();
        foreach (var file in files)
        {
            QueryFiles.Add(file);
        }

        RaisePropertyChanged(nameof(IsSingleQueryMode));
    }

    public void ClearQueryFiles()
    {
        QueryFiles.Clear();
        RaisePropertyChanged(nameof(IsSingleQueryMode));
    }

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
                var batchResult = await workflowService.ExportQueryFilesAsync(
                    QueryFiles.ToArray(),
                    OutputPath,
                    settings,
                    progress,
                    exportCancellation.Token);

                foreach (var file in batchResult.GeneratedFiles)
                {
                    GeneratedFiles.Add(file);
                }

                StatusText = $"Finished batch export. Generated {GeneratedFiles.Count} workbook files.";
            }
            else
            {
                var result = await workflowService.ExportQueryAsync(
                    QueryText,
                    OutputPath,
                    settings,
                    progress,
                    exportCancellation.Token);

                foreach (var workbook in result.Workbooks)
                {
                    GeneratedFiles.Add(workbook.Path);
                }

                StatusText = $"Finished export. Wrote {result.TotalRecordsRead:N0} rows across {result.Workbooks.Count} workbook file(s).";
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

    public void Dispose()
    {
        exportCancellation?.Dispose();
        exportCancellation = null;
    }

    void ApplySettings(AppSettings settings)
    {
        Server = settings.Server;
        Database = settings.Database;
        UseIntegratedSecurity = settings.ConnectionMode == ConnectionMode.IntegratedSecurity;
        Username = settings.Username;
        Password = settings.Password;
        QueryTimeoutSeconds = settings.QueryTimeoutSeconds.ToString(CultureInfo.InvariantCulture);
        MaxRowsPerWorksheet = settings.MaxRowsPerWorksheet.ToString(CultureInfo.InvariantCulture);
        MaxWorkbookSizeGiB = ((decimal)settings.MaxWorkbookBytes / (1024 * 1024 * 1024)).ToString("0.###", CultureInfo.InvariantCulture);
        MaxResultSetsPerWorkbook = settings.MaxResultSetsPerWorkbook.ToString(CultureInfo.InvariantCulture);
        BoundaryColumns = string.Join(", ", settings.WorksheetBoundaryColumns);
        ResultSetNames = string.Join(", ", settings.ResultSetNames.OrderBy(pair => pair.Key).Select(pair => $"{pair.Key}={pair.Value}"));
        OutputPath = settings.OutputPath;
    }

    AppSettings CreateSettingsSnapshot()
    {
        return new AppSettings
        {
            Server = Server.Trim(),
            Database = Database.Trim(),
            ConnectionMode = UseIntegratedSecurity ? ConnectionMode.IntegratedSecurity : ConnectionMode.UsernamePassword,
            Username = Username.Trim(),
            Password = Password,
            QueryTimeoutSeconds = ParseInt(QueryTimeoutSeconds, nameof(QueryTimeoutSeconds)),
            MaxRowsPerWorksheet = ParseInt(MaxRowsPerWorksheet, nameof(MaxRowsPerWorksheet)),
            MaxWorkbookBytes = ParseWorkbookBytes(MaxWorkbookSizeGiB),
            MaxResultSetsPerWorkbook = ParseInt(MaxResultSetsPerWorkbook, nameof(MaxResultSetsPerWorkbook)),
            IncludeHeadersForEmptyResults = true,
            RewriteOverpunchNumbers = true,
            OutputPath = OutputPath.Trim(),
            WorksheetBoundaryColumns = ParseCsvList(BoundaryColumns),
            ResultSetNames = ParseResultSetNames(ResultSetNames),
        };
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

    static int ParseInt(string value, string fieldName)
    {
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) || parsed < 0)
        {
            throw new InvalidOperationException($"{fieldName} must be a non-negative whole number.");
        }

        return parsed;
    }

    static long ParseWorkbookBytes(string gibibytes)
    {
        if (!decimal.TryParse(gibibytes, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed) || parsed <= 0)
        {
            throw new InvalidOperationException("Max workbook size must be a positive number of GiB.");
        }

        return decimal.ToInt64(decimal.Round(parsed * 1024 * 1024 * 1024, 0, MidpointRounding.AwayFromZero));
    }

    static List<string> ParseCsvList(string text)
    {
        return text
            .Split([',', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
    }

    static Dictionary<int, string> ParseResultSetNames(string text)
    {
        var result = new Dictionary<int, string>();
        var entries = text.Split([',', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var entry in entries)
        {
            var parts = entry.Split('=', 2, StringSplitOptions.TrimEntries);
            if (parts.Length != 2 || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var key))
            {
                throw new InvalidOperationException("Result set names must use the format '1=Sheet Name, 2=Another Name'.");
            }

            if (!string.IsNullOrWhiteSpace(parts[1]))
            {
                result[key] = parts[1];
            }
        }

        return result;
    }
}
