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

    public class Workbook
    {
        #region LocalFields

        WorkBookParams runParameters;
        QueryReader queryReader;
        bool queryRan;

        IExcelBackend pts;
        readonly object workerLock = new object();
        // lock to read/write
        IList<WorkerProgress> workers;

        int fileSize;
        bool firstResult = true;

        #endregion

        #region Fields

        /// <summary>
        /// Advance the reader to the next result set.
        /// </summary>
        public bool NextResult()
        {
            // the first result doesn't need to advance the reader
            if (firstResult)
            {
                firstResult = false;
                return true;
            }
            else
            {
                return queryReader.MoveToNextResultSet();
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// Event raised when the file has been written and closed.
        /// </summary>
        public event EventHandler<ReaderFinishedEvents> ReaderFinished;
        /// <summary>
        /// Event raised when the query has started execution.
        /// </summary>
        public event EventHandler<EventArgs> QueryStarted;
        /// <summary>
        /// Event raised every 5 seconds after the query has executed, emitting the total rows/second received from the Sql server.
        /// </summary>
        public event EventHandler<QueryRowsOverTimeEvents> QueryRowsOverTime;
        /// <summary>
        /// Event raised when an exception has been thrown during query execution.
        /// </summary>
        public event EventHandler<QueryExceptionEvents> QueryException;

        #endregion

        #region PublicMethods

        /// <summary>
        /// Represents an Excel-compatible workbook file, with a specified WorkbookParams.
        /// </summary>
        /// <param name="p1"></param>
        public Workbook(WorkBookParams p1)
        {
            this.runParameters = p1;
            // max rows is actually max rows subtract 1, since we're including row headers
            this.runParameters.MaxRowsPerSheet = this.runParameters.MaxRowsPerSheet - 1;

            queryReader = new QueryReader(runParameters.Query, runParameters.QueryTimeout, runParameters.FromFile, runParameters.ConnectionString);

            workers = new List<WorkerProgress>();

            Thread t = new Thread(EmitRowsOverTime);
            t.IsBackground = true;
            WorkerProgress w = new WorkerProgress();
            w.t = t;
            w.cancelled = false;

            lock (workerLock)
                workers.Add(w);
        }

        /// <summary>
        /// Execute the query used to populate the workbook.
        /// </summary>
        /// <returns>False if the query raises an exception.</returns>
        public bool RunQuery()
        {
            // start the query
            OnQueryStarted();
            try
            {
                queryReader.OpenReader();
            }
            catch (SqlException e)
            {
                if (QueryException != null)
                    QueryException(this, new QueryExceptionEvents(e));
                else
                    throw;

                return false;
            }

            lock (workerLock)
                if (!workers.First().cancelled)
                    workers.First().t.Start();

            queryRan = true;
            return queryRan;
        }

        /// <summary>
        /// Write the query result set(s) to the specified filename, starting a new worksheet for each resultset.
        /// </summary>
        public WorkBookStatus WriteQueryResults(string path)
        {
        	FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            WorkBookStatus t=WriteQueryResults(fs, true);
            fs.Close();
            return t;
        }
        
         /// <summary>
        /// Write the query result set(s) to the specified stream, starting a new worksheet for each resultset.
        /// </summary>
        public WorkBookStatus WriteQueryResults(Stream stream)
        {
            return WriteQueryResults(stream, true);
        }

        /// <summary>
        /// Write the current query result set to the specified filename.
        /// </summary>
        public WorkBookStatus WriteQueryResult(string path)
        {
            FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            WorkBookStatus t=WriteQueryResults(fs, false);
            fs.Close();
            return t;
        }

        /// <summary>
        /// Write the current query result set to the specified stream.
        /// </summary>
        public WorkBookStatus WriteQueryResult(Stream stream)
        {
            return WriteQueryResults(stream, false);
        }

        /// <summary>
        /// Release any workbook-created resources.
        /// </summary>
        public void QueryClose()
        {
            queryReader.CloseReader();
            OnReaderFinished();
        }

        #endregion

        WorkBookStatus WriteQueryResults(Stream path, bool multipleSheetsPerStream)
        {
            if (!queryRan)
            {
                throw new WorkbookNotStartedException();
            }

            switch (runParameters.BackendMethod)
            {
                case ExcelBackend.Xlsx:
                    pts = new XlsxParts(path);
                    break;
                    // FIXME tghis isnt an xml backend :(
                case ExcelBackend.Xml:
                    pts = new XlsxParts(path);
                    break;
                default:
                    pts = new XlsxParts(path);
                    break;
            }

            int rowCount = 0;
            int sheetSubCount = 1;
            bool worksheetOpen = false;
            bool workbookTooBig = false;

        READ:
            if (runParameters.WriteEmptyResultSetColumns)
            {
                // create the worksheet
                pts.CreateSheet(queryReader.CurrentResult
                    , sheetSubCount
                    , retrieveSheetName(queryReader.CurrentResult, sheetSubCount)
                    , queryReader.GetSchemaTable().Rows
                    );
                // mark the worksheet as open
                worksheetOpen = true;
            }

            while (queryReader.MoveNext())
            {
                if (!worksheetOpen)
                {
                    // create the worksheet
                    pts.CreateSheet(queryReader.CurrentResult
                        , sheetSubCount
                        , retrieveSheetName(queryReader.CurrentResult, sheetSubCount)
                        , queryReader.GetSchemaTable().Rows
                        );
                    // mark the worksheet as open
                    worksheetOpen = true;
                }
                // write the row
                
                pts.WriteRow(queryReader);
                // increment the row count for this worksheet
                rowCount++;

                if (fileSize >= runParameters.MaxWorkBookSize)
                {
                    workbookTooBig = true;
                    fileSize = 0;
                    break;
                }

                // if we've hit the max number of rows per sheet
                if (rowCount % runParameters.MaxRowsPerSheet == 0)
                {
                    // close the worksheet
                    pts.CloseSheet();
                    // mark the worksheet as closed
                    worksheetOpen = false;
                    // reset the rowcount
                    rowCount = 0;
                    // increment the sheet sub count
                    sheetSubCount++;
                    // start writing a new worksheet
                    goto READ;
                }
            }

            sheetSubCount = 1;

            // if the worksheet is still open, close it
            if (worksheetOpen)
            {
                pts.CloseSheet();
                worksheetOpen = false;
                rowCount = 0;
            }

            // if we're writing all result sets to the same workbook
            if (multipleSheetsPerStream && !workbookTooBig)
            {
                // if there's another result set, start writing the next sheet
                if (queryReader.MoveToNextResultSet())
                    goto READ;
                // otherwise, close the file and release the query
                else
                {
                    pts.Close();
                    QueryClose();
                }
            }
            
            fileSize = 0;

            if (workbookTooBig)
                return WorkBookStatus.OverSize;
            else
                return WorkBookStatus.Completed;
        }

        string retrieveSheetName(int sheetCount, int sheetSubCount)
        {

            if (runParameters.ResultNames != null && runParameters.ResultNames.ContainsKey(sheetCount))
            {

                return runParameters.ResultNames[sheetCount] + "_" + sheetSubCount.ToString();
            }
            else
                return "Sheet" + sheetCount.ToString() + "_" + sheetSubCount.ToString();
        }

        #region ProgressMethods

        void OnReaderFinished()
        {
            lock (workerLock)
            {
                WorkerProgress wb = workers.First();
                if (wb.t != null)
                    wb.cancelled = true;
            }

            if (ReaderFinished != null)
                //ReaderFinished(queryReader.TotalRecordsRead);
                ReaderFinished(this, new ReaderFinishedEvents(queryReader.TotalRecordsRead));
        }

        void OnQueryStarted()
        {
            if (QueryStarted != null)
                QueryStarted(this, EventArgs.Empty);
        }

        // called from separate thread
        void OnQueryRowsOverTime(decimal rowsPerSecond, int totalRows)
        {
            if (QueryRowsOverTime != null)
                QueryRowsOverTime(this, new QueryRowsOverTimeEvents(rowsPerSecond, totalRows));
        }

        // called from separate thread
        void EmitRowsOverTime()
        {
            DateTime t = DateTime.Now;
            int totalRows = queryReader.TotalRecordsRead;

            while (true)
            {
                Thread.Sleep(5000);
                totalRows = queryReader.TotalRecordsRead;
                DateTime now = DateTime.Now;
                TimeSpan ts = now - t;

                bool cancelled = false;
                lock (workerLock)
                {
                    WorkerProgress wb = workers.First();
                    if (wb.t != null && wb.t == Thread.CurrentThread && wb.cancelled)
                        cancelled = true;
                }
                if (cancelled)
                    break;
                else
                    OnQueryRowsOverTime((decimal)totalRows / (decimal)ts.TotalSeconds, totalRows);
            }
        }

        #endregion

    }
}