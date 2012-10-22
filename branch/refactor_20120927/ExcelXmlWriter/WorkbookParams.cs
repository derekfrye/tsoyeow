using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Data;
using System.Collections;
using ExcelXmlWriter;
using System.Threading;
using System.Globalization;
using System.ComponentModel;
using System.Data.SqlClient;
using ExcelXmlWriter.Properties;
using System.Xml.Linq;
using System.Xml;
using ExcelXmlWriter.Xlsx;

namespace ExcelXmlWriter
{
    /// <summary>
    /// An Excel data type. A setting of general forces the workbook to infer each cell's type 
    /// throughout query execution (without truncating numbers over Excel's length limit). The default is string.
    /// </summary>
    public enum ExcelDataType { String, Number, Date, General, OverpunchNumber }

    /// <summary>
    /// An Excel data type. A setting of general forces the workbook to infer each cell's type 
    /// throughout query execution (without truncating numbers over Excel's length limit). The default is string.
    /// </summary>
    public enum ExcelBackend { Xml, Xlsx }

    /// <summary>
    /// 
    /// </summary>
    public enum WorkBookStatus
    {
        /// <summary>
        /// The WorkBook has written an entire result set to the stream.
        /// </summary>
        Completed,
        /// <summary>
        /// The WorkBook exceeded the maximum file size before writing the entire result set to the stream.
        /// </summary>
        OverSize
    }

    class WorkerProgress
    {
        public Thread t;
        public bool cancelled;
    }

    /// <summary>
    /// Required parameters for constructing an Excel 2003 XML-compatible workbook file from a query.
    /// </summary>
    public class WorkBookParams
    {
        /// <summary>
        /// The query to execute and output results to the workbook.
        /// </summary>
        public string Query
        { get; set; }
        /// <summary>
        /// The Sql Server connection string.
        /// </summary>
        public string ConnectionString
        { get; set; }
        /// <summary>
        /// The filetype to write.
        /// </summary>
        public ExcelBackend BackendMethod
        { get; set; }
        /// <summary>
        /// If true, the query is a path to an XML file representation of a DataSet.
        /// </summary>
        public bool FromFile
        { get; set; }
        /// <summary>
        /// The maximum number of rows to write per worksheet before starting a new worksheet with remaining rows.
        /// </summary>
        public int MaxRowsPerSheet
        { get; set; }
        /// <summary>
        /// If provided, a 1-based index of column numbers and the corresponding ExcelDataType to cast each written value.
        /// </summary>
        public Dictionary<int, ExcelDataType> ColumnTypeMappings
        { get; private set; }
        /// <summary>
        /// If provided, a 1-based index of result set numbers and the corresponding WorkSheet name.
        /// </summary>
        public Dictionary<int, string> ResultNames
        { get; set; }
        /// <summary>
        /// If provided, the time in seconds to wait for the query to execute. The default is 30 seconds.
        /// </summary>
        public int QueryTimeout
        { get; set; }
        /// <summary>
        /// If a result set returns with no rows, this controls whether an empty worksheet is written or not (containing just the column headers). The default is true.
        /// </summary>
        public bool WriteEmptyResultSetColumns
        { get; set; }
        /// <summary>
        /// The maximum workbook size in bytes; this value is checked after writing each row.
        /// Excel refuses to open files larger than 2GiB, so this value defaults to 2,000,000,000 to be safe.
        /// If this value is met or exceeded, the workbook finishes writing the necessary data to close the workbook and stops processing the query results.
        /// The recommended action after exceeding this value is to finish writing the remaining query results to a new stream.
        /// </summary>
        public int MaxWorkBookSize
        { get; set; }

        public bool AutoRewriteOverpunch
        { get; set; }

        public WorkBookParams()
        {
            QueryTimeout = 30;
            ResultNames = new Dictionary<int, string>();
            ColumnTypeMappings = new Dictionary<int, ExcelDataType>();
            WriteEmptyResultSetColumns = true;
            MaxWorkBookSize = 2000000000;
            BackendMethod = ExcelBackend.Xlsx;
        }
    }
}
