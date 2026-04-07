using ExcelXmlQueryResults.App.Models;
using ExcelXmlWriter;

namespace ExcelXmlQueryResults.App.Services;

public sealed class ExportWorkflowService
{
    readonly ConnectionStringFactory connectionStringFactory;
    readonly ExcelExporter exporter;

    public ExportWorkflowService()
        : this(new ExcelExporter(), new ConnectionStringFactory())
    {
    }

    public ExportWorkflowService(ExcelExporter exporter, ConnectionStringFactory connectionStringFactory)
    {
        this.exporter = exporter;
        this.connectionStringFactory = connectionStringFactory;
    }

    public async Task<ExportResult> ExportQueryAsync(
        string queryText,
        string outputPath,
        AppSettings settings,
        IProgress<ExportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ValidateOutputPath(outputPath);
        var request = CreateRequest(queryText, settings);
        return await exporter.ExportAsync(request, outputPath, progress, cancellationToken);
    }

    public async Task<BatchExportResult> ExportQueryFilesAsync(
        IReadOnlyList<string> queryFilePaths,
        string outputPath,
        AppSettings settings,
        IProgress<ExportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(queryFilePaths);

        if (queryFilePaths.Count == 0)
        {
            throw new ArgumentException("At least one query file is required.", nameof(queryFilePaths));
        }

        ValidateOutputPath(outputPath);

        var exports = new List<ExportResult>(queryFilePaths.Count);
        for (int index = 0; index < queryFilePaths.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var queryFile = queryFilePaths[index];
            var queryText = await File.ReadAllTextAsync(queryFile, cancellationToken);
            var batchOutputPath = queryFilePaths.Count == 1
                ? outputPath
                : BuildBatchOutputPath(outputPath, index + 1);

            exports.Add(await ExportQueryAsync(queryText, batchOutputPath, settings, progress, cancellationToken));
        }

        return new BatchExportResult
        {
            Exports = exports,
        };
    }

    public void ValidateOutputPath(string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var fullPath = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException("The output directory does not exist.");
        }

        var probeFile = Path.Combine(directory, $"{Path.GetRandomFileName()}.tmp");
        using (File.Create(probeFile))
        {
        }

        File.Delete(probeFile);
    }

    ExportRequest CreateRequest(string queryText, AppSettings settings)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(queryText);
        ArgumentNullException.ThrowIfNull(settings);

        var request = new ExportRequest
        {
            QueryText = queryText,
            ConnectionString = connectionStringFactory.Build(settings),
            SourceKind = ExportSourceKind.SqlQuery,
            Format = OutputWorkbookFormat.Xlsx,
            QueryTimeoutSeconds = settings.QueryTimeoutSeconds,
            MaxRowsPerWorksheet = settings.MaxRowsPerWorksheet,
            IncludeHeadersForEmptyResults = settings.IncludeHeadersForEmptyResults,
            MaxWorkbookBytes = settings.MaxWorkbookBytes,
            RewriteOverpunchNumbers = settings.RewriteOverpunchNumbers,
            MaxResultSetsPerWorkbook = settings.MaxResultSetsPerWorkbook,
        };

        foreach (var column in settings.WorksheetBoundaryColumns)
        {
            request.WorksheetBoundaryColumns.Add(column);
        }

        foreach (var pair in settings.ResultSetNames)
        {
            request.ResultSetNames[pair.Key] = pair.Value;
        }

        return request;
    }

    static string BuildBatchOutputPath(string outputPath, int sequence)
    {
        var fullPath = Path.GetFullPath(outputPath);
        return Path.Combine(
            Path.GetDirectoryName(fullPath) ?? string.Empty,
            $"{Path.GetFileNameWithoutExtension(fullPath)}_batch{sequence}{Path.GetExtension(fullPath)}");
    }
}
