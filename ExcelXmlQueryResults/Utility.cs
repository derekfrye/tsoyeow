using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ExcelXmlWriter.Workbook;
using System.Text.RegularExpressions;

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
        /// <summary>
        /// Create a new filename of format originalfilename_i.originalextension. 
        /// If originalfilename_i.originalextension exists, it'll create originalfilename_(i+1).originalextension and return that.
        /// </summary>
        /// <param name="i">The number to append to the filename</param>
        /// <param name="originalFileName">The original filename</param>
        /// <returns></returns>
        public static string getIncrFileName(int i, string originalFileName)
        {
            string a = Path.Combine(Path.GetDirectoryName(originalFileName)
                , Path.GetFileNameWithoutExtension(originalFileName)
                + "_" + i.ToString()
                + Path.GetExtension(originalFileName));
            while (File.Exists(a))
            {
                i++;
                a = Path.Combine(Path.GetDirectoryName(originalFileName)
                , Path.GetFileNameWithoutExtension(originalFileName)
                + "_" + i.ToString()
                + Path.GetExtension(originalFileName));
            }

            return a;
        }

        /// <summary>
        /// Create a new filename of format originalfilename_i.originalextension. 
        /// If originalfilename_i.originalextension exists, it'll create originalfilename_(i+1).originalextension and return that.
        /// </summary>
        /// <param name="i">The number to append to the filename</param>
        /// <param name="originalFileName">The original filename</param>
        /// <returns></returns>
        public static string getIncrPaddedFileName(int i, string originalFileName, int padleft)
        {

            var basenm = Path.GetFileName(originalFileName);
            var r = new Regex(@"(.+?)_(\d+)(\.xlsx)", RegexOptions.IgnoreCase);
            var m = r.Match(basenm);

            string a = string.Empty;

            a = Path.Combine(Path.GetDirectoryName(originalFileName)
                    , Path.GetFileNameWithoutExtension(basenm)
                    + "_" + i.ToString().PadLeft(padleft, '0')
                    + Path.GetExtension(originalFileName));


            while (File.Exists(a))
            {
                i++;
                a = Path.Combine(Path.GetDirectoryName(originalFileName)
                    , Path.GetFileNameWithoutExtension(basenm)
                    + "_" + i.ToString().PadLeft(padleft, '0')
                    + Path.GetExtension(originalFileName));
            }

            return a;
        }
    }
}
