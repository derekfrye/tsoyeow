using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections;

namespace ExcelXmlWriter
{
    interface IExcelBackend
    {
        void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders);
        void WriteRow(IDataReader data);
        string[] WriteRow(IDataReader data, string[] columnValuesToReturn);
        void CloseSheet();
        void Close();
        long FileSize
        { get; }
    }
}
