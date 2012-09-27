using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelXmlWriter
{
    static class XlsxCell
    {
        internal static string hdr(ExcelDataType excelDataType)
        {
            StringBuilder sb = new StringBuilder();
            switch (excelDataType)
            {
                case ExcelDataType.Date:
                    sb.Append(@"<c s=""1""><v>");
                    break;
                case ExcelDataType.Number:
                    sb.Append(@"<c><v>");
                    break;
                case ExcelDataType.String:
                case ExcelDataType.General:
                default:
                    sb.Append(@"<c t=""s""><v>");
                    break;    
            }
            
            return sb.ToString();
        }
        internal static string hdrclose
        {
            get
            {
                return "</v></c>" + Environment.NewLine;
            }
        }
    }
}
