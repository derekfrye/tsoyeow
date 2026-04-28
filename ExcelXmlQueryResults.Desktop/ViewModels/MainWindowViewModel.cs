using System.Collections.ObjectModel;
using ExcelXmlQueryResults.App.Services;
using ExcelXmlQueryResults.Desktop.Infrastructure;

namespace ExcelXmlQueryResults.Desktop.ViewModels;

public sealed partial class MainWindowViewModel : ViewModelBase, IDisposable
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

    public void Save()
    {
        settingsStore.Save(CreateSettingsSnapshot());
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

    public void Dispose()
    {
        exportCancellation?.Cancel();
        exportCancellation?.Dispose();
        exportCancellation = null;
    }
}
