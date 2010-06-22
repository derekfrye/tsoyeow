using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;

namespace ExcelXmlWriter
{
    class XlsxWorksheetPartCollection
    {
        IList<XlsxWorksheet> worksheets;
        IList<PackagePart> packageparts;

        internal XlsxWorksheetPartCollection()
        {
            worksheets = new List<XlsxWorksheet>();
            packageparts = new List<PackagePart>();
        }

        internal void Add(XlsxWorksheet w1, System.IO.Packaging.PackagePart pt1)
        {
            worksheets.Add(w1);
            packageparts.Add(pt1);
        }

        internal IList<XlsxWorksheet> aa
        { get { return worksheets; } }

        internal Stream retrieveStream(XlsxWorksheet ww)
        {
            int i = worksheets.IndexOf(ww);
            return packageparts[i].GetStream();
        }
    }
}
