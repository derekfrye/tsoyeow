using System.Data;

namespace ExcelXmlWriter.Workbook
{
    internal sealed partial class WorkbookExportSession
    {
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
                    return CloseBackend(backend, tracking, WorkBookStatus.BreakCompleted, "Saving incremental result...");
                }

                var schema = ResultSetSchema.Create(queryReader.GetSchemaTable());
                WriteBufferedOrEmptyResultSet(backend, tracking, schema);

                while (queryReader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!TryOpenWorksheet(backend, tracking, schema))
                    {
                        continue;
                    }

                    var activeSchema = schema!;
                    DetermineIfRowDependsOnPreviousRow(tracking, activeSchema);
                    if (TryWriteCurrentRow(backend, tracking, activeSchema, queryReader) is { } status)
                    {
                        return CloseBackend(backend, tracking, status, "Saving incremental result...");
                    }
                }

                CloseWorksheet(tracking, backend);
                resultSetTotal++;
                hasResultSet = queryReader.NextResult();
            }

            OnSaveStarted("Saving final result...");
            backend.Close();
            Close();
            return WorkBookStatus.Completed;
        }

        void WriteBufferedOrEmptyResultSet(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema? schema)
        {
            if (schema is null)
            {
                return;
            }

            if (options.WriteEmptyResultSetColumns)
            {
                InitializeSheet(backend, tracking, schema);
            }

            if (bufferedRow is null)
            {
                return;
            }

            if (!tracking.WorksheetOpen)
            {
                InitializeSheet(backend, tracking, schema);
            }

            DetermineIfRowDependsOnPreviousRow(tracking, schema);
            TryWriteCurrentRow(backend, tracking, schema, new BufferedDataRecord(bufferedRow));
            bufferedRow = null;
        }

        static void CloseWorksheet(WorkbookTracking tracking, IExcelBackend backend)
        {
            if (!tracking.WorksheetOpen)
            {
                return;
            }

            backend.CloseSheet();
            tracking.WorksheetOpen = false;
            tracking.RowCount = 0;
        }

        bool TryOpenWorksheet(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema? schema)
        {
            if (!tracking.WorksheetOpen && schema is not null)
            {
                InitializeSheet(backend, tracking, schema);
            }

            return tracking.WorksheetOpen && schema is not null;
        }

        WorkBookStatus? TryWriteCurrentRow(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema, IDataRecord dataRecord)
        {
            if (tracking.RowCount > 0 && tracking.RowCount % options.MaxRowsPerSheet == 0)
            {
                tracking.Status = WorkBookStatus.BreakWanted;
            }

            if (!AdvanceWorksheetState(backend, tracking, schema))
            {
                return tracking.Status;
            }

            backend.WriteRow(dataRecord);
            tracking.RowCount++;
            return null;
        }

        bool AdvanceWorksheetState(IExcelBackend backend, WorkbookTracking tracking, ResultSetSchema schema)
        {
            if (!tracking.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent)
            {
                tracking.Status = WorkBookStatus.Pending;
                return true;
            }

            if (tracking.Status == WorkBookStatus.BreakWanted)
            {
                backend.CloseSheet();
                tracking.SheetSubCount++;
                tracking.WorksheetOpen = false;
                tracking.RowCount = 0;
                InitializeSheet(backend, tracking, schema);
                tracking.Status = WorkBookStatus.BreakCompleted;
                return false;
            }

            if (backend.FileSize >= options.MaxWorkBookSize)
            {
                bufferedRow = Enumerable.Range(0, queryReader.FieldCount)
                    .Select(queryReader.GetValue)
                    .ToArray();
                tracking.Status = WorkBookStatus.OverSize;
                return false;
            }

            tracking.Status = WorkBookStatus.Pending;
            return true;
        }

        WorkBookStatus CloseBackend(IExcelBackend backend, WorkbookTracking tracking, WorkBookStatus status, string message)
        {
            OnSaveStarted(message);
            CloseWorksheet(tracking, backend);
            backend.Close();
            return status;
        }
    }
}
