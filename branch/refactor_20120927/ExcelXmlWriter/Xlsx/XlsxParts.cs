using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;
using ExcelXmlWriter.Properties;
using System.Data;
using Ionic.Zip;
using ExcelXmlWriter.Xlsx;
using System.Xml.Linq;
using System.Xml;

namespace ExcelXmlWriter
{

    // fixme
    public class FixMeContentTypes
    {
        protected XDocument appXml = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));
        XNamespace xn11 = "http://schemas.openxmlformats.org/package/2006/content-types";
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
        internal FixMeContentTypes(XlsxWorksheetPartCollection worksheets)
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

            foreach (var x in worksheets.aa)
            {
                 appXml.Elements().First(xx => xx.Name.LocalName== "Types"&&xx.Name.Namespace==xn11).Add(new XElement(xn11 + "Override", new XAttribute("PartName", x.filenm)
                        , new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            }
        }
    }



    public class XlsxParts : IExcelBackend
    {

        //Package packageObject;
        ZipFile z;
        Rels r;
        Rels relsforwkshts;
        XlsxWorkbookMetadata workbookXml;
        ZipAAA workbookXmlPackagePart;
        XlsxSharedStringsXml sharedStrings;

        XlsxWorksheet currentWorksheet;
        XlsxWorksheetPartCollection worksheets;

        public void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders)
        {
            //Uri u4 = new Uri(, UriKind.Relative);
            
            string shtnm ="worksheets/sheet" + sheetCount.ToString() + "_" + subSheetCount.ToString() + ".xml";
            string filnm = "/xl/"+shtnm;
            var djfk = new ZipAAA() { path = shtnm, RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" };
            relsforwkshts.Link(djfk);
            string id = relsforwkshts.Id(djfk).ToString();

            //PackagePart pt = packageObject.CreatePart(u4, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
            //string id = workbookXmlPackagePart.CreateRelationship(new Uri("worksheets/sheet" + sheetCount.ToString()
            //    + "_" + subSheetCount.ToString() + ".xml", UriKind.Relative), TargetMode.Internal
            //   , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").Id;

            string jf=Path.GetTempFileName();
            var strm1 = new FileStream(jf, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var strm = z.AddEntry(filnm, strm1).InputStream;
            XlsxWorksheet w = new XlsxWorksheet(strm, sheetCount, subSheetCount, sheetName, id, resultHeaders, sharedStrings,jf);
            w.filenm = filnm;


            

            
            worksheets.Add(w);
            currentWorksheet = w;
        }

        public void CloseSheet()
        {
            currentWorksheet.Close();
            //StaticFunctions.copyStream(currentWorksheet.s, worksheets.retrieveStream(currentWorksheet));
        }

        public void Close()
        {
            XlsxAppXml a = new XlsxAppXml();
            //a.LinkToPackage(packageObject);
            a.SetSheetCount(worksheets.aa);
            r.Link(a.LinkToPackage());

            workbookXml.SetSheetCount(worksheets.aa);

            //sharedStrings.close();

            //packageObject.Close();
            z.AddEntry("_rels/.rels", r.Write().ToString());

            z.AddEntry("/docProps/app.xml", a.Write().ToString());

            FixMeContentTypes f = new FixMeContentTypes(worksheets);
            z.AddEntry("/[Content_Types].xml", f.Write());

            z.AddEntry("/xl/workbook.xml", workbookXml.Write());

            z.AddEntry("/xl/_rels/workbook.xml.rels", relsforwkshts.Write());
            sharedStrings.close();
            z.Save();

            foreach (var djfak in worksheets.aa)
            {
                djfak.sasdf.Close();
                if(File.Exists(djfak.filenmOs))
                File.Delete(djfak.filenmOs);
            }

            
            sharedStrings.s.Close();
            if (File.Exists(sharedStrings.jf))
                File.Delete(sharedStrings.jf);
        }

        public XlsxParts(string path)
        {
            //packageObject = Package.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            if (File.Exists(path))
                File.Delete(path);
            z = new ZipFile(path);

             r = new Rels();
             relsforwkshts = new Rels();

            worksheets = new XlsxWorksheetPartCollection();

            

            #region workbook.xml_Content_Type

            workbookXml = new XlsxWorkbookMetadata();
            workbookXmlPackagePart = workbookXml.LinkToPackage();
            r.Link(workbookXmlPackagePart);
            

            #endregion

            #region core.xml

            XlsxPart core = new XlsxPart(Settings.Default.CoreXml);
            XDocument xdjfkd = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkd = XDocument.Parse(Settings.Default.CoreXml.OuterXml, LoadOptions.None);
            r.Link(core.LinkToPackage("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties","docProps/core.xml"
                ));
            

            StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8);

            var za = new XmlWriterSettings();
            za.Encoding = Encoding.UTF8;

            XmlWriter apo = XmlWriter.Create(sb, za);
            xdjfkd.Save(apo);
            apo.Close();

            z.AddEntry("/docProps/core.xml", sb.ToString());
            //core.close();

            #endregion

            #region theme

            //Uri u7 = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
            //PackagePart p7 = packageObject.CreatePart(u7, "application/vnd.openxmlformats-officedocument.theme+xml", CompressionOption.Normal);
            //Settings.Default.ThemeXml.Save(p7.GetStream());
            //p7.GetStream().Close();
            //workbookXmlPackagePart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");
            XlsxPart theme = new XlsxPart(Settings.Default.ThemeXml);
            XDocument xdjfkdd = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkdd = XDocument.Parse(Settings.Default.ThemeXml.OuterXml, LoadOptions.None);
            relsforwkshts.Link(theme.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme","theme/theme1.xml"
                ));


            StringWriterWithEncoding sbs = new StringWriterWithEncoding(Encoding.UTF8);

            var zas = new XmlWriterSettings();
            zas.Encoding = Encoding.UTF8;

            XmlWriter apos = XmlWriter.Create(sbs, zas);
            xdjfkdd.Save(apos);
            apos.Close();

            z.AddEntry("/xl/theme/theme1.xml", sbs.ToString());

            #endregion

            #region styles

            //Uri u8 = new Uri("/xl/styles.xml", UriKind.Relative);
            //PackagePart p8 = packageObject.CreatePart(u8, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
            //Settings.Default.StylesXml.Save(p8.GetStream());
            //p8.GetStream().Close();
            //workbookXmlPackagePart.CreateRelationship(new Uri("styles.xml", UriKind.Relative), TargetMode.Internal
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            XlsxPart styles = new XlsxPart(Settings.Default.StylesXml);
            XDocument xdjfkdsa = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

            xdjfkdsa = XDocument.Parse(Settings.Default.StylesXml.OuterXml, LoadOptions.None);
           relsforwkshts.Link(styles.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles","styles.xml"
                ));


            StringWriterWithEncoding sasb = new StringWriterWithEncoding(Encoding.UTF8);

            var zaaa = new XmlWriterSettings();
            zaaa.Encoding = Encoding.UTF8;

            XmlWriter apso = XmlWriter.Create(sasb, zaaa);
            xdjfkdsa.Save(apso);
            apso.Close();

            z.AddEntry("/xl/styles.xml", sasb.ToString());

            #endregion

            #region sharedstrings

            
            string jf = Path.GetTempFileName();
            var strm1 = new FileStream(jf, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var t = z.AddEntry("/xl/sharedStrings.xml", strm1);
            sharedStrings = new XlsxSharedStringsXml(strm1, jf);
            //r.Link(sharedStrings.LinkToPackage());
            relsforwkshts.Link(sharedStrings.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings","sharedStrings.xml"));

            #endregion

            

            
        }

        public void WriteRow(IDataReader queryReader)
        {
            
            currentWorksheet.writerow(queryReader, sharedStrings);
        }
    }
}
