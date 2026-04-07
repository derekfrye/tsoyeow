using System.Globalization;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter.Xlsx
{
    internal static class XlsxData
    {
        static readonly DateTime ExcelSince = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);

        internal static string ConvertToWriteableValue(string value, ExcelDataType dataType)
        {
            return dataType switch
            {
                ExcelDataType.Date => (Convert.ToDateTime(value, CultureInfo.CurrentCulture) - ExcelSince)
                    .TotalDays
                    .ToString(CultureInfo.InvariantCulture),
                ExcelDataType.Number => value.Trim(),
                _ => value,
            };
        }
    }
}
