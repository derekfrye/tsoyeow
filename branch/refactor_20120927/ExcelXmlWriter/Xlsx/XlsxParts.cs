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

namespace ExcelXmlWriter
{
    public class XlsxParts : IExcelBackend
    {

        //Package packageObject;
        ZipFile z;
        XlsxWorkbookMetadata workbookXml;
        ZipAAA workbookXmlPackagePart;
        XlsxSharedStringsXml sharedStrings;

        XlsxWorksheet currentWorksheet;
        XlsxWorksheetPartCollection worksheets;

        public void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders)
        {
            Uri u4 = new Uri("/xl/worksheets/sheet" + sheetCount.ToString()
                + "_" + subSheetCount.ToString() + ".xml", UriKind.Relative);
            //PackagePart pt = packageObject.CreatePart(u4, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
            //string id = workbookXmlPackagePart.CreateRelationship(new Uri("worksheets/sheet" + sheetCount.ToString()
            //    + "_" + subSheetCount.ToString() + ".xml", UriKind.Relative), TargetMode.Internal
            //   , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").Id;
            //var strm = pt.GetStream();
            //XlsxWorksheet w = new XlsxWorksheet(strm, sheetCount, subSheetCount, sheetName, id, resultHeaders, sharedStrings);

            //worksheets.Add(w, pt);
            //currentWorksheet = w;
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

            workbookXml.SetSheetCount(worksheets.aa);

            sharedStrings.close();

            //packageObject.Close();
        }

        public XlsxParts(string path)
        {
            //packageObject = Package.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            z = new ZipFile(path);

            Rels r = new Rels();

            worksheets = new XlsxWorksheetPartCollection();

            

            #region workbook.xml_Content_Type

            workbookXml = new XlsxWorkbookMetadata();
            workbookXmlPackagePart = workbookXml.LinkToPackage();
            r.Link(workbookXmlPackagePart);

            #endregion

            #region core.xml

            XlsxPart core = new XlsxPart(Settings.Default.CoreXml);
            
            r.Link(core.LinkToPackage(
                ));
            //core.close();

            #endregion

            #region theme

            //Uri u7 = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
            //PackagePart p7 = packageObject.CreatePart(u7, "application/vnd.openxmlformats-officedocument.theme+xml", CompressionOption.Normal);
            //Settings.Default.ThemeXml.Save(p7.GetStream());
            //p7.GetStream().Close();
            //workbookXmlPackagePart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");

            #endregion

            #region styles

            //Uri u8 = new Uri("/xl/styles.xml", UriKind.Relative);
            //PackagePart p8 = packageObject.CreatePart(u8, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
            //Settings.Default.StylesXml.Save(p8.GetStream());
            //p8.GetStream().Close();
            //workbookXmlPackagePart.CreateRelationship(new Uri("styles.xml", UriKind.Relative), TargetMode.Internal
            //    , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");

            #endregion

            #region sharedstrings

            //sharedStrings = new XlsxSharedStringsXml(packageObject, workbookXmlPackagePart);

            #endregion

            z.AddEntry("_rels/.rels", r.Write().ToString());

            z.Save();
        }

        public void WriteRow(IDataReader queryReader)
        {
            if (currentWorksheet.closed)
                throw new Exception("Sorry, worksheet closed!");
            currentWorksheet.writerow(queryReader, sharedStrings);
        }
    }
}
