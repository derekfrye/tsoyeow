/*
 * Created by SharpDevelop.
 * User: djfrye
 * Date: 10/30/2012
 * Time: 2:50 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace ExcelXmlWriter.Workbook
{
    /// <summary>
     /// Various statues a workbook can be in.
     /// </summary>
    internal enum WorkBookStatus
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

    sealed class DupKeyResults
    {
        public string[]? PrevDupKey
        {
            get; set;
        }
        public string[]? CurrentRowDupKey
        {
            get; set;
        }
        public bool PreviousDiffersFromCurrent
        {
            get; set;
        }
    }

    sealed class WorkbookTracking
    {
        public WorkBookStatus Status
        {
            get; set;
        }
        public int SheetSubCount
        {
            get; set;
        }
        public bool WorksheetOpen
        {
            get; set;
        }
        public int RowCount
        {
            get; set;
        }
        public DupKeyResults PreviousAndCurrentRowKeyColumns
        {
            get; set;
        }


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
