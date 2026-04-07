using System.Data;
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
        readonly CancellationToken cancellationToken;

        CancellationTokenSource? progressCancellationSource;
        Task? progressTask;
        bool queryOpened;
        bool closed;
        object?[]? bufferedRow;

        internal WorkbookExportSession(WorkBookParams options, CancellationToken cancellationToken = default)
        {
            this.options = options;
            this.cancellationToken = cancellationToken;
            queryReader = new QueryReader(options.Query, options.QueryTimeout, options.FromFile, options.ConnectionString);
        }

        internal event EventHandler<EventArgs>? QueryStarted;

        internal event EventHandler<ExportProgressEventArgs>? QueryProgress;

        internal event EventHandler<SaveStartedEventArgs>? SaveStarted;

        internal event EventHandler<ExportFinishedEventArgs>? Finished;

        internal int TotalRecordsRead => queryReader.TotalRecordsRead;

        internal void Open()
        {
            cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();
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

                if (schema is not null && bufferedRow is not null)
                {
                    if (!tracking.WorksheetOpen)
                    {
                        InitializeSheet(backend, tracking, schema);
                    }

                    DetermineIfRowDependsOnPreviousRow(tracking, schema);
                    WriteCurrentRow(backend, tracking, schema, new BufferedDataRecord(bufferedRow));
                    bufferedRow = null;

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

                while (queryReader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!tracking.WorksheetOpen && schema is not null)
                    {
                        InitializeSheet(backend, tracking, schema);
                    }

                    if (!tracking.WorksheetOpen || schema is null)
                    {
                        continue;
                    }

                    DetermineIfRowDependsOnPreviousRow(tracking, schema);
                    WriteCurrentRow(backend, tracking, schema, queryReader);

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
            progressCancellationSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            progressTask = Task.Run(() => EmitRowsOverTimeAsync(progressCancellationSource.Token), CancellationToken.None);
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

        void WriteCurrentRow(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema, IDataRecord dataRecord)
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
                    bufferedRow = Enumerable.Range(0, queryReader.FieldCount)
                        .Select(queryReader.GetValue)
                        .ToArray();
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

            backend.WriteRow(dataRecord);
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

        sealed class BufferedDataRecord : IDataRecord
        {
            readonly object?[] values;

            internal BufferedDataRecord(object?[] values)
            {
                this.values = values;
            }

            public int FieldCount => values.Length;

            public object this[int i] => GetValue(i);

            public object this[string name] => throw new NotSupportedException();

            public bool GetBoolean(int i) => Convert.ToBoolean(values[i], CultureInfo.InvariantCulture);

            public byte GetByte(int i) => Convert.ToByte(values[i], CultureInfo.InvariantCulture);

            public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => throw new NotSupportedException();

            public char GetChar(int i) => Convert.ToChar(values[i], CultureInfo.InvariantCulture);

            public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => throw new NotSupportedException();

            public IDataReader GetData(int i) => throw new NotSupportedException();

            public string GetDataTypeName(int i) => GetFieldType(i).Name;

            public DateTime GetDateTime(int i) => Convert.ToDateTime(values[i], CultureInfo.InvariantCulture);

            public decimal GetDecimal(int i) => Convert.ToDecimal(values[i], CultureInfo.InvariantCulture);

            public double GetDouble(int i) => Convert.ToDouble(values[i], CultureInfo.InvariantCulture);

            public Type GetFieldType(int i) => values[i]?.GetType() ?? typeof(object);

            public float GetFloat(int i) => Convert.ToSingle(values[i], CultureInfo.InvariantCulture);

            public Guid GetGuid(int i) => values[i] is Guid guid ? guid : Guid.Parse(values[i]?.ToString() ?? string.Empty);

            public short GetInt16(int i) => Convert.ToInt16(values[i], CultureInfo.InvariantCulture);

            public int GetInt32(int i) => Convert.ToInt32(values[i], CultureInfo.InvariantCulture);

            public long GetInt64(int i) => Convert.ToInt64(values[i], CultureInfo.InvariantCulture);

            public string GetName(int i) => i.ToString(CultureInfo.InvariantCulture);

            public int GetOrdinal(string name) => throw new NotSupportedException();

            public string GetString(int i) => values[i]?.ToString() ?? string.Empty;

            public object GetValue(int i) => values[i] ?? DBNull.Value;

            public int GetValues(object[] targetValues)
            {
                var length = Math.Min(targetValues.Length, values.Length);
                for (int index = 0; index < length; index++)
                {
                    targetValues[index] = GetValue(index);
                }

                return length;
            }

            public bool IsDBNull(int i) => values[i] is null or DBNull;
        }
    }
}
