using System;
using System.Collections.Generic;
using System.Text;
using ExcelXmlWriter;
using System.Globalization;
using System.IO;

namespace ExcelXmlWriter
{
   
    static class XlsxData
    {
        
        static readonly DateTime ExcelSince = new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc); 

        internal static string DataVal(string val, ExcelDataType e)
        {
          //  StringBuilder sb = new StringBuilder();

          if(e == ExcelDataType.Date)
          	return (((TimeSpan) (Convert.ToDateTime(val, CultureInfo.CurrentCulture) - ExcelSince)).TotalDays).ToString();
          else if(e==ExcelDataType.Number)
                    return val.Trim();
                    else 
                    	return val;
                
            }
    } 
}
