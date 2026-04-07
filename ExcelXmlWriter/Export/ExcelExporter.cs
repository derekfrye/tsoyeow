using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter;

public sealed class ExcelExporter
{
    public Task<ExportResult> ExportAsync(
        ExportRequest request,
        string destinationPath,
        IProgress<ExportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => Export(request, destinationPath, progress, cancellationToken), cancellationToken);
    }

    ExportResult Export(
        ExportRequest request,
        string destinationPath,
        IProgress<ExportProgress>? progress,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ValidateRequest(request);

        var workbookPaths = new List<ExportedWorkbook>();
        var options = CreateWorkbookParameters(request);
        var normalizedDestinationPath = NormalizeDestinationPath(destinationPath, request.Format);

        string? currentOutputPath = null;
        int totalRecordsRead = 0;

        using var session = new WorkbookExportSession(options, cancellationToken);
        session.QueryStarted += (_, _) =>
        {
            progress?.Report(new ExportProgress
            {
                Stage = ExportStage.QueryStarting,
            });
        };
        session.QueryProgress += (_, args) =>
        {
            progress?.Report(new ExportProgress
            {
                Stage = ExportStage.QueryRunning,
                RowsPerSecond = args.RowsPerSecond,
                TotalRows = args.TotalRows,
                OutputPath = currentOutputPath,
            });
        };
        session.SaveStarted += (_, args) =>
        {
            progress?.Report(new ExportProgress
            {
                Stage = ExportStage.SavingWorkbook,
                Message = args.Message,
                TotalRows = session.TotalRecordsRead,
                OutputPath = currentOutputPath,
            });
        };
        session.Finished += (_, args) =>
        {
            totalRecordsRead = args.TotalRecordsRead;
            progress?.Report(new ExportProgress
            {
                Stage = ExportStage.Completed,
                TotalRows = args.TotalRecordsRead,
                OutputPath = currentOutputPath,
            });
        };

        session.Open();

        int sequence = 1;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            currentOutputPath = sequence == 1
                ? normalizedDestinationPath
                : BuildIncrementedPath(normalizedDestinationPath, sequence);

            var status = session.WriteQueryResults(currentOutputPath);
            workbookPaths.Add(new ExportedWorkbook
            {
                SequenceNumber = sequence,
                Path = currentOutputPath,
            });

            if (status == WorkBookStatus.Completed)
            {
                break;
            }

            if (status is not WorkBookStatus.OverSize and not WorkBookStatus.BreakCompleted)
            {
                throw new InvalidOperationException($"Unexpected workbook export status '{status}'.");
            }

            sequence++;
        }

        return new ExportResult
        {
            TotalRecordsRead = totalRecordsRead,
            Workbooks = workbookPaths,
        };
    }

    static void ValidateRequest(ExportRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.QueryText);

        if (request.SourceKind == ExportSourceKind.SqlQuery)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(request.ConnectionString);
        }

        if (request.Format != OutputWorkbookFormat.Xlsx)
        {
            throw new NotSupportedException($"The '{request.Format}' output format is not supported.");
        }

        if (request.MaxRowsPerWorksheet <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request.MaxRowsPerWorksheet));
        }

        if (request.MaxWorkbookBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request.MaxWorkbookBytes));
        }

        if (request.MaxResultSetsPerWorkbook <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request.MaxResultSetsPerWorkbook));
        }
    }

    static WorkBookParams CreateWorkbookParameters(ExportRequest request)
    {
        var parameters = new WorkBookParams
        {
            Query = request.QueryText,
            ConnectionString = request.ConnectionString,
            FromFile = request.SourceKind == ExportSourceKind.XmlDataSetFile,
            MaxRowsPerSheet = request.MaxRowsPerWorksheet,
            QueryTimeout = request.QueryTimeoutSeconds,
            WriteEmptyResultSetColumns = request.IncludeHeadersForEmptyResults,
            MaxWorkBookSize = request.MaxWorkbookBytes,
            AutoRewriteOverpunch = request.RewriteOverpunchNumbers,
            MaximumResultSetsPerWorkbook = request.MaxResultSetsPerWorkbook,
            DupeKeysToDelayStartingNewWorksheet = request.WorksheetBoundaryColumns.ToArray(),
        };

        foreach (var resultName in request.ResultSetNames)
        {
            parameters.ResultNames[resultName.Key] = resultName.Value;
        }

        foreach (var mapping in request.ColumnTypeMappings)
        {
            parameters.ColumnTypeMappings[mapping.Key] = mapping.Value;
        }

        return parameters;
    }

    static string NormalizeDestinationPath(string destinationPath, OutputWorkbookFormat format)
    {
        var fullPath = Path.GetFullPath(destinationPath);
        var extension = Path.GetExtension(fullPath);
        if (format == OutputWorkbookFormat.Xlsx && !string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return Path.ChangeExtension(fullPath, ".xlsx");
        }

        return fullPath;
    }

    static string BuildIncrementedPath(string basePath, int sequence)
    {
        return Path.Combine(
            Path.GetDirectoryName(basePath) ?? string.Empty,
            $"{Path.GetFileNameWithoutExtension(basePath)}_{sequence}{Path.GetExtension(basePath)}");
    }
}
