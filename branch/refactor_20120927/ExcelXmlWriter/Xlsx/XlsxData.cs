using System;
using System.Collections.Generic;
using System.Text;
using ExcelXmlWriter;
using System.Globalization;
using System.IO;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter.Xlsx
{
    static class XlsxData
    {
        static readonly DateTime ExcelSince = new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);

        internal static string ConvertToWriteableValue(string value, ExcelDataType dataType)
        {
            if (dataType == ExcelDataType.Date)
                return (((TimeSpan)(Convert.ToDateTime(value, CultureInfo.CurrentCulture) - ExcelSince)).TotalDays).ToString();
            else if (dataType == ExcelDataType.Number)
                return value.Trim();
            else
                return value;

        }
    }
}
