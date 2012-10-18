using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using System.IO.Packaging;
using ExcelXmlWriter.Xlsx;

namespace ExcelXmlWriter
{

    class ZipAAA
    {
        public string path
        { get; set; }
        public string RelType
        {get;set;}
    }

    /// <summary>
    /// /xl/workbook.xml
    /// </summary>
    class XlsxWorkbookMetadata 
    {
        protected XDocument appXml = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));
        XNamespace xn1 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace xn11 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal ZipAAA LinkToPackage()
        {
            //return base.LinkToPackage(p, new Uri("/xl/workbook.xml", UriKind.Relative)
            //    , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            //    );
            return new ZipAAA() { path = "xl/workbook.xml", RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" };
        }

        internal void SetSheetCount(IList<XlsxWorksheet> sheets)
        {
            int order = 1;
            foreach (var lk in sheets)
            {
                appXml.Element(xn11 + "workbook").Element(xn11 + "sheets").Add(
                    new XElement(xn11 + "sheet"
                        , new XAttribute("name", lk.sheetname)
                        , new XAttribute("sheetId", order)
                        , new XAttribute(xn1 + "id", lk.Id)
                    )
                );
                order++;
            }

            //base.close();
        }

        public string Write()
        {
            //XNamespace xn2 = "http://schemas.openxmlformats.org/package/2006/relationships";

            ////<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

            //var t = new XElement(xn2 + "Relationships");
            //foreach (var z in rels)
            //{
            //    t.Add(new XElement(xn2 + "Relationship", new XAttribute("Type", z.Value.RelType), new XAttribute("Target", z.Value.path), new XAttribute("Id", z.Key.ToString())));
            //}
            //XDocument x = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), t);

            StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8);

            var za = new XmlWriterSettings();
            za.Encoding = Encoding.UTF8;

            XmlWriter apo = XmlWriter.Create(sb, za);
            appXml.Save(apo);
            apo.Close();
            return sb.ToString();
        }

        internal XlsxWorkbookMetadata()
        {
            appXml.Add(
                new XElement(xn11 + "workbook"

                    , new XAttribute(XNamespace.Xmlns + "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                    , new XElement(xn11 + "fileVersion"
                        , new XAttribute("appName", "xl")
                        , new XAttribute("lastEdited", "4")
                        , new XAttribute("lowestEdited", "4")
                        , new XAttribute("rupBuild", "4506"))
                        , new XElement(xn11 + "workbookPr"
                            , new XAttribute("defaultThemeVersion", "124226")
                        )
                        , new XElement(xn11 + "bookViews"
                            , new XElement(xn11 + "workbookView"
                                , new XAttribute("xWindow", "360")
                                , new XAttribute("yWindow", "120")
                                , new XAttribute("windowWidth", "22860")
                                , new XAttribute("windowHeight", "11385")
                                , new XAttribute("activeTab", "0")
                            )
                        )
                        , new XElement(xn11 + "sheets")
                        , new XElement(xn11 + "calcPr"
                            , new XAttribute("calcId", "125725")
                        )
                    )
                );
        }
    }
}