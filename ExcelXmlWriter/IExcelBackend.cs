using System.Data;

namespace ExcelXmlWriter
{
    interface IExcelBackend
    {
        void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders);
        void WriteRow(IDataReader data);
        void WriteRow(IDataReader data, string[] columnValuesToReturn);
        string[] ReadKeyValues(IDataReader queryReader, string[] colsToObtainValsFrom);
        void CloseSheet();
        void Close();
        long FileSize
        {
            get;
        }
    }





}
