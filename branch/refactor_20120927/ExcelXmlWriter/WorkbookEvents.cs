using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelXmlWriter
{
    public class ReaderFinishedEvents: EventArgs
    {
        public int totalRecordsRead
        {
            get;
            private set;
        }

        public ReaderFinishedEvents(int a)
        {
            this.totalRecordsRead = a;
        }
    }

    public class QueryRowsOverTimeEvents : EventArgs
    {
        public decimal rowsPerSecond
        { get; private set; }
        public int total
        {
            get;
            private set;
        }

        public QueryRowsOverTimeEvents(decimal rowsPerSecond, int total)
        {
            this.rowsPerSecond = rowsPerSecond;
            this.total = total;
        }
    }

    public class QueryExceptionEvents : EventArgs
    {
        public Exception e 
        { get; private set; }

        public QueryExceptionEvents(Exception e)
        {
            this.e = e;
        }
    }
}
