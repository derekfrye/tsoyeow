using Microsoft.Data.SqlClient;

namespace ExcelXmlWriter.Workbook
{
    public class Workbook : IDisposable
    {
        readonly WorkbookExportSession session;

        /// <summary>
        /// Event raised when the file has been written and closed.
        /// </summary>
        public event EventHandler<ReaderFinishedEventArgs>? ReaderFinished;

        /// <summary>
        /// Event raised when the query has started execution.
        /// </summary>
        public event EventHandler<EventArgs>? QueryStarted;

        /// <summary>
        /// Event raised every 5 seconds after the query has executed, emitting the total rows/second received from the Sql server.
        /// </summary>
        public event EventHandler<QueryRowsOverTimeEventArgs>? QueryRowsOverTime;

        /// <summary>
        /// Event raised when an exception has been thrown during query execution.
        /// </summary>
        public event EventHandler<QueryExceptionEventArgs>? QueryException;

        /// <summary>
        /// Raised when the workbook has completed reading records, and has begun saving results to output.
        /// </summary>
        public event EventHandler<SaveFileEventArgs>? SaveFile;

        public Workbook(WorkBookParams parameters)
        {
            var compatParameters = CreateCompatibilityParameters(parameters);
            session = new WorkbookExportSession(compatParameters);
            session.QueryStarted += (_, args) => QueryStarted?.Invoke(this, args);
            session.QueryProgress += (_, args) => QueryRowsOverTime?.Invoke(this, new QueryRowsOverTimeEventArgs(args.RowsPerSecond, args.TotalRows));
            session.SaveStarted += (_, args) => SaveFile?.Invoke(this, new SaveFileEventArgs(args.Message));
            session.Finished += (_, args) => ReaderFinished?.Invoke(this, new ReaderFinishedEventArgs(args.TotalRecordsRead));
        }

        static WorkBookParams CreateCompatibilityParameters(WorkBookParams parameters)
        {
            var backend = parameters.BackendMethod == ExcelBackend.Xml
                ? ExcelBackend.Xlsx
                : parameters.BackendMethod;

            var compatibilityParameters = new WorkBookParams
            {
                Query = parameters.Query,
                ConnectionString = parameters.ConnectionString,
                BackendMethod = backend,
                FromFile = parameters.FromFile,
                MaxRowsPerSheet = parameters.MaxRowsPerSheet,
                QueryTimeout = parameters.QueryTimeout,
                WriteEmptyResultSetColumns = parameters.WriteEmptyResultSetColumns,
                MaxWorkBookSize = parameters.MaxWorkBookSize,
                AutoRewriteOverpunch = parameters.AutoRewriteOverpunch,
                DupeKeysToDelayStartingNewWorksheet = parameters.DupeKeysToDelayStartingNewWorksheet.ToArray(),
                MaximumResultSetsPerWorkbook = parameters.MaximumResultSetsPerWorkbook,
                ResultNames = new Dictionary<int, string>(parameters.ResultNames),
            };

            foreach (var mapping in parameters.ColumnTypeMappings)
            {
                compatibilityParameters.ColumnTypeMappings.Add(mapping.Key, mapping.Value);
            }

            return compatibilityParameters;
        }

        public bool NextResult()
        {
            throw new NotSupportedException("The compatibility Workbook facade no longer exposes manual result-set traversal.");
        }

        public bool RunQuery()
        {
            try
            {
                session.Open();
                return true;
            }
            catch (SqlException exception)
            {
                if (QueryException is null)
                {
                    throw;
                }

                QueryException(this, new QueryExceptionEventArgs(exception));
                return false;
            }
        }

        public WorkBookStatus WriteQueryResults(string path)
        {
            return session.WriteQueryResults(path);
        }

        public WorkBookStatus WriteQueryResults(Stream stream)
        {
            return session.WriteQueryResults(stream);
        }

        public void QueryClose()
        {
            session.Close();
        }

        public void Dispose()
        {
            session.Dispose();
            GC.SuppressFinalize(this);
        }
    }
}
