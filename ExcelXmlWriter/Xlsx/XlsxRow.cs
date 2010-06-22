using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelXmlWriter
{
    static class XlsxRow
    {
        internal static string hdr
        {
            get
            {
                return "<row>" + Environment.NewLine;
            }
        }
        internal static string hdrclose
        {
            get
            {
                return "</row>" + Environment.NewLine;
            }
        }
    }
}
