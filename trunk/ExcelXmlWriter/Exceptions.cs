using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace ExcelXmlWriter
{
    [Serializable()]
    public class WorkbookUnfinishedException : Exception
    {
        public WorkbookUnfinishedException() : base(Resource1.WorkbookUnfinished) { }
        public WorkbookUnfinishedException(string message) : base(message) { throw new NotImplementedException(); }
        public WorkbookUnfinishedException(string message, Exception innerException) : base(message, innerException) { throw new NotImplementedException(); }
        protected WorkbookUnfinishedException(SerializationInfo info, StreamingContext context) : base(info, context) { throw new NotImplementedException(); }
    }
    [Serializable()]
    public class WorkbookNotStartedException : Exception
    {
        public WorkbookNotStartedException() : base(Resource1.WorkbookNotStarted) { }
        public WorkbookNotStartedException(string message) : base(message) { throw new NotImplementedException(); }
        public WorkbookNotStartedException(string message, Exception innerException) : base(message, innerException) { throw new NotImplementedException(); }
        protected WorkbookNotStartedException(SerializationInfo info, StreamingContext context) : base(info, context) { throw new NotImplementedException(); }
    }
    [Serializable()]
    public class WorkbookCannotWriteException : Exception
    {
        public WorkbookCannotWriteException() : base(Resource1.WorkbookCannotWrite) { }
        public WorkbookCannotWriteException(string message) : base(message) { throw new NotImplementedException(); }
        public WorkbookCannotWriteException(string message, Exception innerException) : base(message, innerException) { throw new NotImplementedException(); }
        protected WorkbookCannotWriteException(SerializationInfo info, StreamingContext context) : base(info, context) { throw new NotImplementedException(); }
    }
}