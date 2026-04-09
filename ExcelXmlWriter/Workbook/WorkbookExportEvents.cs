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
}
