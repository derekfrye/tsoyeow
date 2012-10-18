using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /[Content_Types].xml
    /// </summary>
    internal class ContentTypes : XlsxPart
    {
        XNamespace xn11 = "http://schemas.openxmlformats.org/package/2006/content-types";

        internal ContentTypes(List<XlsxWorksheet> worksheets)
        {
            appXml.Add(
                new XElement(xn11 + "Types"
                    , new XElement(xn11 + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml"))
                    , new XElement(xn11 + "Default", new XAttribute("Extension", "rels")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/xl/workbook.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/docProps/core.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/xl/theme/theme1.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/xl/styles.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))
                    , new XElement(xn11 + "Override", new XAttribute("PartName", "/docProps/app.xml")
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"))
                ));

            foreach (var x in worksheets)
            {
                appXml.Elements().First(xx => xx.Name.LocalName == "Types" && xx.Name.Namespace == xn11).Add(new XElement(xn11 + "Override", new XAttribute("PartName", x.PackageFileName)
                       , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            }
        }
    }
}
