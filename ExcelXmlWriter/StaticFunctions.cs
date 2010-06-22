using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;

namespace ExcelXmlWriter
{
    static class StaticFunctions
    {
        static readonly int excelMaxNumberLength = Convert.ToInt32(Resource1.ExcelMaximumNumberLength, CultureInfo.InvariantCulture);

        // move to static helper class
        internal static void copyStream(Stream input, Stream output)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = input.Read(buf, 0, bufSize)) > 0)
                output.Write(buf, 0, bytesRead);
        }

        internal static ExcelDataType ResolveDataType(string dataValue)
        {
            decimal throwaway = new decimal();
            DateTime d = new DateTime();

            // Credit Card problem
            if (Decimal.TryParse(dataValue, out throwaway)
                && dataValue.Trim().Length <= excelMaxNumberLength
                && !dataValue.Trim().EndsWith("-")
                )
                return ExcelDataType.Number;
            else if (DateTime.TryParse(dataValue, out d)
                // excel doesn't like dates before 1900
                && d >= Convert.ToDateTime("1900-01-01")
                && d <= Convert.ToDateTime("9999-12-31"))
                return ExcelDataType.Date;
            else
                return ExcelDataType.String;
        }
    }
}
