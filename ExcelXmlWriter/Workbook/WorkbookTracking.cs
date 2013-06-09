/*
 * Created by SharpDevelop.
 * User: djfrye
 * Date: 10/30/2012
 * Time: 2:50 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace ExcelXmlWriter.Workbook
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
    /// Various statues a workbook can be in.
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
        OverSize,
        Pending,
        BreakCompleted,
        BreakWanted,
    }
    
	 class DupKeyResults
    {
        public string[] PrevDupKey
        { get; set; }
        public string[] CurrentRowDupKey
        { get; set; }
        public bool PreviousDiffersFromCurrent
        { get; set; }
    }

    class WorkbookTracking
    {
        public WorkBookStatus Status
        { get; set; }
        public int SheetSubCount
        { get; set; }
        public bool WorksheetOpen 
        { get; set; }
        public int RowCount
        { get; set; }
        public DupKeyResults PreviousAndCurrentRowKeyColumns
        { get; set; }
        

        public WorkbookTracking()
        {
            SheetSubCount = 1;
            WorksheetOpen = false;
            RowCount = 1;
            Status = WorkBookStatus.Pending;
            PreviousAndCurrentRowKeyColumns = new DupKeyResults();
        }
    }
}
