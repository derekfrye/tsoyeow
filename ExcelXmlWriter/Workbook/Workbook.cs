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
            //// max rows is actually max rows subtract 1, since we're including row headers
            //this.runParameters.MaxRowsPerSheet = this.runParameters.MaxRowsPerSheet - 1;

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
        /// Will exit with the query still open if MaximumResultSetsPerWorkbook is exceeded.
        /// </summary>
        public WorkBookStatus WriteQueryResults(string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                WorkBookStatus t = WriteResults(fs);
                return t;
            }
        }
        
         /// <summary>
        /// Write the query result set(s) to the specified stream, starting a new worksheet for each resultset.
        /// </summary>
        public WorkBookStatus WriteQueryResults(Stream stream)
        {
            return WriteResults(stream);
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

        void InitSheet(WorkbookTracking w)
        {
            var t = queryReader.GetSchemaTable();
            if (t != null)
            {
                pts.CreateSheet(queryReader.CurrentResult
                    , w.SheetSubCount
                    , retrieveSheetName(queryReader.CurrentResult, w.SheetSubCount)
                    , queryReader.GetSchemaTable().Rows
                    );
                w.WorksheetOpen = true;
            }
        }

        /// <summary>
        /// Writes the results.
        /// </summary>
        /// <returns>The results.</returns>
        /// <param name="path">Path.</param>
        WorkBookStatus WriteResults(Stream path)
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

            WorkbookTracking w = new WorkbookTracking();

            // start to keep track of the result sets
            int resultSetTotal = 1;

            // loop through the result sets and write the results
            while (queryReader.MoveToNextResultSet())
            {

                // if this result set is over the max results per workbook, close this one and indicate a break has happened
                if (resultSetTotal > runParameters.MaximumResultSetsPerWorkbook)
                {
                    OnSave("Saving incremental result...");
                    pts.Close();
                    return WorkBookStatus.BreakCompleted;
                }

                resultSetTotal++;

                // write empty columns if requested
                if (runParameters.WriteEmptyResultSetColumns)
                {
                    InitSheet(w);
                }

                // loop through all result sets
                while (queryReader.MoveNext())
                {
                    if (!w.WorksheetOpen)
                    {
                        InitSheet(w);
                    }
                    if (w.WorksheetOpen)
                    {
                        this.DetermineIfRowDependsOnPreviousRow(w);
                        queryReader.Reset();

                        // write the row, or determine why we couldn't write the row
                        WriteARow(w);

                        // if we're over-size, we must return now
                        // but keep the query open bc the caller may request to write the rest of the results
                        if (w.Status == WorkBookStatus.OverSize)
                        {
                            OnSave("Saving incremental result...");
                            if (w.WorksheetOpen)
                            {
                                pts.CloseSheet();
                            }
                            pts.Close();
                            return w.Status;
                        }
                    }
                }

                if (w.WorksheetOpen)
                {
                    pts.CloseSheet();
                    w.WorksheetOpen = false;
                    w.RowCount = 0;
                }

                
            }

            OnSave("Saving final result...");
            pts.Close();
            QueryClose();
            return WorkBookStatus.Completed;
        }

        void DetermineIfRowDependsOnPreviousRow(WorkbookTracking w)
        {
            w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey = pts.ReadKeyValues(queryReader, runParameters.DupeKeysToDelayStartingNewWorksheet);

            // first time through, set these values equal
            if (w.RowCount == 1)
            {
                w.PreviousAndCurrentRowKeyColumns.PrevDupKey = w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;
            }
            // otherwise, if the values in the key columns are identical to previous row, then the desire is to keep them together
            // i.e., cannot split to another sheet or workbook between these rows
            if (w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey != null & w.PreviousAndCurrentRowKeyColumns.PrevDupKey != null)
            {
                for (int i = 0; i < w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey.Length; i++)
                {
                    if (string.Equals(w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey[i], w.PreviousAndCurrentRowKeyColumns.PrevDupKey[i], StringComparison.InvariantCulture))
                    {
                        w.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = false;
                    }
                    else
                    {
                        w.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = true;
                        break;
                    }
                }
            }
            else
                w.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent = true;

            // set the prevDupKey for next cycle
            w.PreviousAndCurrentRowKeyColumns.PrevDupKey = w.PreviousAndCurrentRowKeyColumns.CurrentRowDupKey;

            //return w.keyColumnStatus;
        }

        void WriteARow( WorkbookTracking w)
        {
            if (w.RowCount>0&&w.RowCount % runParameters.MaxRowsPerSheet == 0)
                w.Status = WorkBookStatus.BreakWanted;

            if (w.PreviousAndCurrentRowKeyColumns.PreviousDiffersFromCurrent)
            {
                if (w.Status == WorkBookStatus.BreakWanted)
                {
                    pts.CloseSheet();
                    w.SheetSubCount++;
                    w.WorksheetOpen = false;
                    w.RowCount = 0;
                    InitSheet(w);
                    w.Status= WorkBookStatus.BreakCompleted;

                }
                else if (pts.FileSize >= runParameters.MaxWorkBookSize)
                {
                    w.Status = WorkBookStatus.OverSize;
                    return;
                }
                else
                    w.Status = WorkBookStatus.Pending;
            }
            else
                w.Status = WorkBookStatus.Pending;

            pts.WriteRow(queryReader, runParameters.DupeKeysToDelayStartingNewWorksheet);
            w.RowCount++;
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