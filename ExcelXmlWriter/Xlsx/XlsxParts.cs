﻿using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ExcelXmlWriter.Properties;
using System.Data;
using System.Xml.Linq;
using System.Xml;

using Ionic.Zip; 

namespace ExcelXmlWriter.Xlsx
{
    public class XlsxParts : IExcelBackend, IDisposable
    {
        ZipFile z;
        Stream p;
        Relationships mainRels;
        Relationships xlRels;
        XlsxWorkbookMetadata workbookXml;
        ContentRelationships workbookXmlPackagePart;
        XlsxSharedStringsXml sharedStrings;
        bool alreadyClosed = false;

        XlsxWorksheet currentWorksheet;
        List<XlsxWorksheet> worksheets = new List<XlsxWorksheet>();

        public long FileSize
        {
            get
            {
                long ret = sharedStrings.OutputStream.Length;
                foreach (var wksht in worksheets)
                {
                    ret += wksht.OutputStream.Length;
                }
                return ret;
            }
        }

        public void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders)
        {
            string shtnm = "worksheets/sheet" + sheetCount.ToString() + "_" + subSheetCount.ToString() + ".xml";
            string filnm = "/xl/" + shtnm;
            var djfk = new ContentRelationships() { PackagePath = shtnm, RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" };
            xlRels.Link(djfk);
            string id = xlRels.Id(djfk).ToString();

            XlsxWorksheet w = new XlsxWorksheet(sheetName, id, resultHeaders, sharedStrings, filnm);
            z.AddEntry(filnm, w.OutputStream);

            worksheets.Add(w);
            currentWorksheet = w;
        }

        public void CloseSheet()
        {
            currentWorksheet.Close();
        }

        public void Close()
        {
            if (alreadyClosed)
                return;

            XlsxAppXml a = new XlsxAppXml();

            a.SetSheetCount(worksheets);
            mainRels.Link(XlsxAppXml.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"));

            workbookXml.SetSheetCount(worksheets);

            z.AddEntry("_rels/.rels", mainRels.Write().ToString());

            z.AddEntry("/docProps/app.xml", a.Write().ToString());

            ContentTypes f = new ContentTypes(worksheets);
            z.AddEntry("/[Content_Types].xml", f.Write());

            z.AddEntry("/xl/workbook.xml", workbookXml.Write());

            z.AddEntry("/xl/_rels/workbook.xml.rels", xlRels.Write());
            sharedStrings.Close();
            z.Save(p);

            foreach (var djfak in worksheets)
            {
                djfak.OutputStream.Close();
                if (File.Exists(djfak.FileAssociatedWithOutputStream))
                    File.Delete(djfak.FileAssociatedWithOutputStream);
            }

            sharedStrings.OutputStream.Close();
            if (File.Exists(sharedStrings.FileAssociateWithOutputStream))
                File.Delete(sharedStrings.FileAssociateWithOutputStream);

            alreadyClosed = true;
        }

        public XlsxParts(Stream path)
        {
            p=path;
        	z = new ZipFile();

            mainRels = new Relationships();
            xlRels = new Relationships();

            #region workbook.xml_Content_Type

            workbookXml = new XlsxWorkbookMetadata();
            workbookXmlPackagePart = XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml");
            mainRels.Link(workbookXmlPackagePart);

            #endregion

            #region core.xml

            XDocument xdjfkd = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkd = XDocument.Parse(Settings.Default.CoreXml.OuterXml, LoadOptions.None);
            mainRels.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"
                ));

            using (StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8))
            {
                var za = new XmlWriterSettings();
                za.Encoding = Encoding.UTF8;
                using (XmlWriter apo = XmlWriter.Create(sb, za))
                {
                    xdjfkd.Save(apo);
                    apo.Close();
                    z.AddEntry("/docProps/core.xml", sb.ToString());
                }
            }

            #endregion

            #region theme

            XDocument xdjfkdd = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkdd = XDocument.Parse(Settings.Default.ThemeXml.OuterXml, LoadOptions.None);
            xlRels.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"
                ));

            using (StringWriterWithEncoding sbs = new StringWriterWithEncoding(Encoding.UTF8))
            {
                var zas = new XmlWriterSettings();
                zas.Encoding = Encoding.UTF8;

                using (XmlWriter apos = XmlWriter.Create(sbs, zas))
                {
                    xdjfkdd.Save(apos);
                    apos.Close();
                    z.AddEntry("/xl/theme/theme1.xml", sbs.ToString());
                }
            }
            

            #endregion

            #region styles

            XDocument xdjfkdsa = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkdsa = XDocument.Parse(Settings.Default.StylesXml.OuterXml, LoadOptions.None);
            xlRels.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"
                 ));

            using (StringWriterWithEncoding sasb =new StringWriterWithEncoding(Encoding.UTF8))
            {
                var zaaa = new XmlWriterSettings();
                zaaa.Encoding = Encoding.UTF8;

                using (XmlWriter apso = XmlWriter.Create(sasb, zaaa))
                {
                    xdjfkdsa.Save(apso);
                    apso.Close();
                    z.AddEntry("/xl/styles.xml", sasb.ToString());
                }
            }

            #endregion

            #region sharedstrings

            string jf = Path.GetTempFileName();
            var strm1 = new FileStream(jf, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            z.AddEntry("/xl/sharedStrings.xml", strm1);
            sharedStrings = new XlsxSharedStringsXml(strm1, jf);

            xlRels.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml"));

            #endregion

        }

        public void WriteRow(IDataReader queryReader)
        {
            currentWorksheet.writerow(queryReader);
        }

        public void WriteRow(IDataReader queryReader, string[] columnValuesToReturn)
        {
            currentWorksheet.writerow(queryReader, columnValuesToReturn);
        }
        
        public string[] ReadKeyValues(IDataReader queryReader, string[] colsToObtainValsFrom)
        {
        	return currentWorksheet.ReadKeyValues(queryReader, colsToObtainValsFrom);
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (z != null)
                    z.Dispose();
                if (sharedStrings != null)
                    sharedStrings.Dispose();

                foreach (var w in worksheets)
                {
                    w.Dispose();
                }

            }
        }

        #endregion



        

    }
}
