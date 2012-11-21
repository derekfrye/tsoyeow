using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlQueryResults
{
    enum QueryState
    {
        Running,
        Finished, 
        Saving
    }

    class ExcelXmlQueryResultsParams
    {
        public WorkBookParams e { get; set; }
        public string filenm { get; set; }
    }

    public class Utility
    {
        public static string getIncrFileName(int i, string p3)
        {
            string a =Path.Combine(Path.GetDirectoryName(p3)
                , Path.GetFileNameWithoutExtension(p3)
                + "_" + i.ToString()
                + Path.GetExtension(p3));
            while (File.Exists(a))
            {
                i++;
                a = Path.Combine(Path.GetDirectoryName(p3)
                , Path.GetFileNameWithoutExtension(p3)
                + "_" + i.ToString()
                + Path.GetExtension(p3));
            }

            return a;
        }
    }
}
