﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using ExcelXmlWriter.Xlsx;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /docProps/app.xml
    /// </summary>
    class XlsxAppXml : XlsxPart
    {

        XNamespace xn = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        XNamespace xn2 = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

        internal void SetSheetCount(IList<XlsxWorksheet> sheets)
        {
            XElement xx = appXml.Element(xn2 + "Properties").Element(xn2 + "TitlesOfParts")
                    .Element(xn + "vector");
            xx.RemoveNodes();
            int totalsheets = 0;
            foreach (var lk in sheets)
            {
                xx.Add(new XElement(xn + "lpstr", lk.Sheetname));
                totalsheets++;
            }

            xx.Attribute("size").SetValue(totalsheets);

            appXml.Element(xn2 + "Properties").Element(xn2 + "HeadingPairs")
                .Element(xn + "vector").Elements(xn + "variant").Where(x => x.Elements(xn + "i4").Any()).First().Element(xn + "i4").SetValue(totalsheets);
        }

        internal XlsxAppXml()
        {
            appXml.Add(
                new XElement(xn2 + "Properties"
                    , new XAttribute(XNamespace.Xmlns + "vt"
                        , "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
                    , new XElement(xn2 + "Application", "Microsoft Excel")
                    , new XElement(xn2 + "DocSecurity", "0")
                    , new XElement(xn2 + "ScaleCrop", "false")
                    , new XElement(xn2 + "HeadingPairs"
                        , new XElement(xn + "vector", new XAttribute("size", 2), new XAttribute("baseType", "variant")
                            , new XElement(xn + "variant"
                                , new XElement(xn + "lpstr", "Worksheets")
                            )
                            , new XElement(xn + "variant"
                                , new XElement(xn + "i4", 0)
                            )
                        )
                    )
                    , new XElement(xn2 + "TitlesOfParts"
                        , new XElement(xn + "vector"
                            , new XAttribute("size", 0), new XAttribute("baseType", "lpstr")
                            , new XElement(xn + "lpstr", "placeholder")
                        )
                    )
                    , new XElement(xn2 + "Company", Environment.UserName)
                    , new XElement(xn2 + "LinksUpToDate", false)
                    , new XElement(xn2 + "SharedDoc", false)
                    , new XElement(xn2 + "HyperlinksChanged", false)
                    , new XElement(xn2 + "AppVersion", "12.0000")
             ));
        }
    }
}