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

namespace ExcelXmlWriter.Workbook
{
    

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
        /// The maximum uncompressed total size of all worksheets, in bytes. If this value is exceeded before starting a new worksheet, writing output stops.
        /// Empirical evidence suggests the default of around 2,500,000,000 is safe. The recommended action after exceeding this value is to finish writing the remaining query results to a new stream.
        /// </summary>
        public long MaxWorkBookSize
        { get; set; }

        public bool AutoRewriteOverpunch
        { get; set; }

        public string[] DupeKeysToDelayStartingNewWorksheet
        { get; set; }

        public WorkBookParams()
        {
            QueryTimeout = 30;
            ResultNames = new Dictionary<int, string>();
            ColumnTypeMappings = new Dictionary<int, ExcelDataType>();
            WriteEmptyResultSetColumns = true;
            MaxWorkBookSize = 2500000000;
            BackendMethod = ExcelBackend.Xlsx;
        }
    }
}
