using System.Globalization;
using Microsoft.Data.SqlClient;
using ExcelXmlWriter.Xlsx;

namespace ExcelXmlWriter.Workbook
{
    internal sealed class ExportFinishedEventArgs : EventArgs
    {
        internal ExportFinishedEventArgs(int totalRecordsRead)
        {
            TotalRecordsRead = totalRecordsRead;
        }

        internal int TotalRecordsRead { get; }
    }

    internal sealed class ExportProgressEventArgs : EventArgs
    {
        internal ExportProgressEventArgs(decimal rowsPerSecond, int totalRows)
        {
            RowsPerSecond = rowsPerSecond;
            TotalRows = totalRows;
        }

        internal decimal RowsPerSecond { get; }

        internal int TotalRows { get; }
    }

    internal sealed class SaveStartedEventArgs : EventArgs
    {
        internal SaveStartedEventArgs(string message)
        {
            Message = message;
        }

        internal string Message { get; }
    }

    internal sealed class WorkbookExportSession : IDisposable
    {
        readonly WorkBookParams options;
        readonly QueryReader queryReader;

        CancellationTokenSource? progressCancellationSource;
        Task? progressTask;
        bool queryOpened;
        bool closed;

        internal WorkbookExportSession(WorkBookParams options)
        {
            this.options = options;
            queryReader = new QueryReader(options.Query, options.QueryTimeout, options.FromFile, options.ConnectionString);
        }

        internal event EventHandler<EventArgs>? QueryStarted;

        internal event EventHandler<ExportProgressEventArgs>? QueryProgress;

        internal event EventHandler<SaveStartedEventArgs>? SaveStarted;

        internal event EventHandler<ExportFinishedEventArgs>? Finished;

        internal void Open()
        {
            QueryStarted?.Invoke(this, EventArgs.Empty);
            queryReader.OpenReader();
            queryOpened = true;
            StartProgressLoop();
        }

        internal WorkBookStatus WriteQueryResults(string path)
        {
            using var fileStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            return WriteQueryResults(fileStream);
        }

        internal WorkBookStatus WriteQueryResults(Stream stream)
        {
            if (!queryOpened)
            {
                throw new WorkbookNotStartedException();
            }

            using IExcelBackend backend = CreateBackend(stream);
            var tracking = new WorkbookTracking();
            int resultSetTotal = 1;
            bool hasResultSet = queryReader.HasActiveResultSet;

            while (hasResultSet)
            {
                if (resultSetTotal > options.MaximumResultSetsPerWorkbook)
                {
                    OnSaveStarted("Saving incremental result...");
                    backend.Close();
                    return WorkBookStatus.BreakCompleted;
                }

                var schema = ResultSetSchema.Create(queryReader.GetSchemaTable());
                if (options.WriteEmptyResultSetColumns && schema is not null)
                {
                    InitializeSheet(backend, tracking, schema);
                }

                while (queryReader.Read())
                {
                    if (!tracking.WorksheetOpen && schema is not null)
                    {
                        InitializeSheet(backend, tracking, schema);
                    }

                    if (!tracking.WorksheetOpen || schema is null)
                    {
                        continue;
                    }

                    DetermineIfRowDependsOnPreviousRow(tracking, schema);
                    WriteCurrentRow(backend, tracking, schema);

                    if (tracking.Status == WorkBookStatus.OverSize)
                    {
                        OnSaveStarted("Saving incremental result...");
                        if (tracking.WorksheetOpen)
                        {
                            backend.CloseSheet();
                        }

                        backend.Close();
                        return tracking.Status;
                    }
                }

                if (tracking.WorksheetOpen)
                {
                    backend.CloseSheet();
                    tracking.WorksheetOpen = false;
                    tracking.RowCount = 0;
                }

                resultSetTotal++;
                hasResultSet = queryReader.NextResult();
            }

            OnSaveStarted("Saving final result...");
            backend.Close();
            Close();
            return WorkBookStatus.Completed;
        }

        internal void Close()
        {
            if (closed)
            {
                return;
            }

            closed = true;
            queryReader.CloseReader();
            StopProgressLoop();
            Finished?.Invoke(this, new ExportFinishedEventArgs(queryReader.TotalRecordsRead));
        }

        static IExcelBackend CreateBackend(Stream stream)
        {
            return new XlsxParts(stream);
        }

        void StartProgressLoop()
        {
            progressCancellationSource = new CancellationTokenSource();
            progressTask = Task.Run(() => EmitRowsOverTimeAsync(progressCancellationSource.Token));
        }

        void StopProgressLoop()
        {
            if (progressCancellationSource is null)
            {
                return;
            }

            progressCancellationSource.Cancel();

            try
            {
                progressTask?.GetAwaiter().GetResult();
            }
            catch (OperationCanceledException)
            {
            }
            finally
            {
                progressCancellationSource.Dispose();
                progressCancellationSource = null;
                progressTask = null;
            }
        }

        async Task EmitRowsOverTimeAsync(CancellationToken cancellationToken)
        {
            var start = DateTime.UtcNow;
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(5));

            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                var totalRows = queryReader.TotalRecordsRead;
                var elapsed = DateTime.UtcNow - start;
                QueryProgress?.Invoke(
                    this,
                    new ExportProgressEventArgs((decimal)totalRows / Math.Max((decimal)elapsed.TotalSeconds, 0.001m), totalRows));
            }
        }

        void InitializeSheet(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema)
        {
            backend.CreateSheet(
                queryReader.CurrentResult,
                tracking.SheetSubCount,
                RetrieveSheetName(queryReader.CurrentResult, tracking.SheetSubCount),
                schema);
            tracking.WorksheetOpen = true;
        }

        void DetermineIfRowDependsOnPreviousRow(WorkbookTracking tracking, ResultSetSchema schema)
        {
            tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey = ReadDuplicateKeyValues(schema);
            if (tracking.RowCount == 1)
            {
                tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey = tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;
            }

            if (tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey is not null
                && tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey is not null)
            {
                tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = tracking.PreviousAndCurrentRowKeyColumns
                    .CurrentRowDupKey
                    .Zip(tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey, (current, previous) => string.Equals(current, previous, StringComparison.Ordinal))
                    .Any(matches => !matches);
            }
            else
            {
                tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = true;
            }

            tracking.PreviousAndCurrentRowKeyColumns.PrevDupKey = tracking.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;
        }

        string[]? ReadDuplicateKeyValues(ResultSetSchema schema)
        {
            if (options.DupeKeysToDelayStartingNewWorksheet.Length == 0)
            {
                return null;
            }

            var values = new string[options.DupeKeysToDelayStartingNewWorksheet.Length];
            for (int index = 0; index < values.Length; index++)
            {
                var columnName = options.DupeKeysToDelayStartingNewWorksheet[index];
                if (schema.ContainsColumn(columnName))
                {
                    values[index] = queryReader[columnName]?.ToString() ?? string.Empty;
                }
            }

            return values;
        }

        void WriteCurrentRow(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema)
        {
            if (tracking.RowCount > 0 && tracking.RowCount % options.MaxRowsPerSheet == 0)
            {
                tracking.Status = WorkBookStatus.BreakWanted;
            }

            if (tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent)
            {
                if (tracking.Status == WorkBookStatus.BreakWanted)
                {
                    backend.CloseSheet();
                    tracking.SheetSubCount++;
                    tracking.WorksheetOpen = false;
                    tracking.RowCount = 0;
                    InitializeSheet(backend, tracking, schema);
                    tracking.Status = WorkBookStatus.BreakCompleted;
                }
                else if (backend.FileSize >= options.MaxWorkBookSize)
                {
                    tracking.Status = WorkBookStatus.OverSize;
                    return;
                }
                else
                {
                    tracking.Status = WorkBookStatus.Pending;
                }
            }
            else
            {
                tracking.Status = WorkBookStatus.Pending;
            }

            backend.WriteRow(queryReader);
            tracking.RowCount++;
        }

        string RetrieveSheetName(int sheetCount, int sheetSubCount)
        {
            if (options.ResultNames.TryGetValue(sheetCount, out var sheetName))
            {
                return sheetName + "_" + sheetSubCount.ToString(CultureInfo.InvariantCulture);
            }

            return "Sheet" + sheetCount.ToString(CultureInfo.InvariantCulture) + "_" + sheetSubCount.ToString(CultureInfo.InvariantCulture);
        }

        void OnSaveStarted(string message)
        {
            SaveStarted?.Invoke(this, new SaveStartedEventArgs(message));
        }

        public void Dispose()
        {
            if (!closed)
            {
                Close();
            }

            queryReader.Dispose();
        }
    }
}
