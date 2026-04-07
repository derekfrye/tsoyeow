using System.Data;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter
{
    interface IExcelBackend : IDisposable
    {
        void CreateSheet(int sheetCount, int subSheetCount, string sheetName, ResultSetSchema schema);
        void WriteRow(IDataRecord data);
        void CloseSheet();
        void Close();
        long FileSize
        {
            get;
        }
    }
}
