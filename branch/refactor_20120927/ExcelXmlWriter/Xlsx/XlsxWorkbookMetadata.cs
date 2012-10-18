﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Xml;
using System.IO.Packaging;

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
    class XlsxWorkbookMetadata : XlsxPart
    {

        XNamespace xn1 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace xn11 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal ZipAAA LinkToPackage()
        {
            //return base.LinkToPackage(p, new Uri("/xl/workbook.xml", UriKind.Relative)
            //    , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            //    );
            return new ZipAAA() { path = "/xl/workbook.xml", RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" };
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