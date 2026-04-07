namespace ExcelXmlWriter.Workbook
{
    public sealed class ReaderFinishedEventArgs : EventArgs
    {
        public ReaderFinishedEventArgs(int totalRecordsRead)
        {
            TotalRecordsRead = totalRecordsRead;
        }

        public int TotalRecordsRead { get; }
    }

    public sealed class QueryRowsOverTimeEventArgs : EventArgs
    {
        public QueryRowsOverTimeEventArgs(decimal rowsPerSecond, int totalRows)
        {
            RowsPerSecond = rowsPerSecond;
            TotalRows = totalRows;
        }

        public decimal RowsPerSecond { get; }

        public int TotalRows { get; }
    }

    public sealed class QueryExceptionEventArgs : EventArgs
    {
        public QueryExceptionEventArgs(Exception exception)
        {
            Exception = exception;
        }

        public Exception Exception { get; }
    }

    public sealed class SaveFileEventArgs : EventArgs
    {
        public SaveFileEventArgs(string message)
        {
            Message = message;
        }

        public string Message { get; }
    }
}
