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
    public class Workbook:IDisposable
    {
        #region LocalFields

        WorkBookParams runParameters;
        QueryReader queryReader;
        bool queryRan;

        IExcelBackend pts;
        readonly object workerLock = new object();
        // lock to read/write
        IList<WorkerProgress> workers;

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
        /// <summary>
        /// Raised when the workbook has completed reading records, and has begun saving results to output.
        /// </summary>
        public event EventHandler<SaveFileEvent> SaveFile;

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
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                WorkBookStatus t = WriteQueryResults(fs, true);
                return t;
            }
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
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                WorkBookStatus t = WriteQueryResults(fs, false);
                return t;
            }
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
                // FIXME this isnt an xml backend
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
            string[] prevDupKey = null;
            string[] newDupKey = null;
            bool canBreak = true;
            bool breakWanted=false;

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
                newDupKey = pts.WriteRow(queryReader, runParameters.DupeKeysToDelayStartingNewWorksheet);

                // first time through, set these values equal
                if (rowCount == 0)
                {
                    prevDupKey = newDupKey;
                }
                // otherwise, if the values in the key columns are identical to previous row, then the desire is to keep them together
                // i.e., cannot split to another sheet or workbook between these rows
                if (newDupKey != null & prevDupKey != null)
                {
                    for (int i = 0; i < newDupKey.Length; i++)
                    {
                        if (string.Equals(newDupKey[i], prevDupKey[i], StringComparison.InvariantCulture))
                        {
                            canBreak = false;                            
                        }
                        else
                        {
                            canBreak = true;
                            break;
                        }
                    }
                }
                else
                    canBreak = true;

                // set the prevDupKey for next cycle
                prevDupKey = newDupKey;

                // increment the row count for this worksheet
                rowCount++;

                if(rowCount % runParameters.MaxRowsPerSheet == 0)
                    breakWanted=true;

                if (canBreak && pts.FileSize >= runParameters.MaxWorkBookSize)
                {
                    workbookTooBig = true;
                    break;
                }

                // if we've hit the max number of rows per sheet
                if (canBreak && breakWanted)
                {
                    breakWanted = false;
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
            breakWanted = false;

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
                    OnSave("Saving final result...");
                    pts.Close();
                    QueryClose();
                }
            }

            if (workbookTooBig)
            {
                OnSave("Saving incremental result...");
                pts.Close();
                return WorkBookStatus.OverSize;
            }
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

        void OnSave(string message)
        {
            if (SaveFile != null)
                SaveFile(this, new SaveFileEvent(message) );
        }

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


        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                queryReader.Dispose();                
            }
        }

        #endregion
    }
}