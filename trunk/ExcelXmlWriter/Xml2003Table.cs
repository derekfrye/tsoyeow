using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace ExcelXmlWriter
{
    static class Xml2003Table
    {
        internal static string hdr(int colCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<Table ");
            sb.Append("ss:ExpandedColumnCount=\"" + colCount.ToString(CultureInfo.CurrentCulture) 
                + "\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:DefaultRowHeight=\"15\">");
            sb.Append(Environment.NewLine);
            return sb.ToString();
        }
        internal static string hdrclose
        {
            get
            {
                return "</Table>" + Environment.NewLine;
            }
        }
    }
}
