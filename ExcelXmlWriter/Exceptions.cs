namespace ExcelXmlWriter
{
    [Serializable()]
    public class WorkbookUnfinishedException : Exception
    {
        public WorkbookUnfinishedException() : base(Resource1.WorkbookUnfinished) { }
        public WorkbookUnfinishedException(string message) : base(message) { }
        public WorkbookUnfinishedException(string message, Exception innerException) : base(message, innerException) { }
    }

    [Serializable()]
    public class WorkbookNotStartedException : Exception
    {
        public WorkbookNotStartedException() : base(Resource1.WorkbookNotStarted) { }
        public WorkbookNotStartedException(string message) : base(message) { }
        public WorkbookNotStartedException(string message, Exception innerException) : base(message, innerException) { }
    }

    [Serializable()]
    public class WorkbookCannotWriteException : Exception
    {
        public WorkbookCannotWriteException() : base(Resource1.WorkbookCannotWrite) { }
        public WorkbookCannotWriteException(string message) : base(message) { }
        public WorkbookCannotWriteException(string message, Exception innerException) : base(message, innerException) { }
    }
}
