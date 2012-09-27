using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;
using ExcelXmlWriter.Properties;
using System.Data;

namespace ExcelXmlWriter
{
    class XlsxParts : IExcelBackend
    {

        Package packageObject;
        XlsxWorkbookMetadata workbookXml;
        PackagePart workbookXmlPackagePart;
        XlsxSharedStringsXml sharedStrings;

        XlsxWorksheet currentWorksheet;
        XlsxWorksheetPartCollection worksheets;

        public void CreateSheet(int sheetCount, int subSheetCount, string sheetName, DataRowCollection resultHeaders)
        {
            Uri u4 = new Uri("/xl/worksheets/sheet" + sheetCount.ToString()
                + "_" + subSheetCount.ToString() + ".xml", UriKind.Relative);
            PackagePart pt = packageObject.CreatePart(u4, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Normal);
            string id = workbookXmlPackagePart.CreateRelationship(new Uri("worksheets/sheet" + sheetCount.ToString()
                + "_" + subSheetCount.ToString() + ".xml", UriKind.Relative), TargetMode.Internal
               , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").Id;
            XlsxWorksheet w = new XlsxWorksheet(sheetCount, subSheetCount, sheetName, id, resultHeaders, sharedStrings);

            worksheets.Add(w, pt);
            currentWorksheet = w;
        }

        public void CloseSheet()
        {
            currentWorksheet.Close();
            StaticFunctions.copyStream(currentWorksheet.s, worksheets.retrieveStream(currentWorksheet));
        }

        public void Close()
        {
            XlsxAppXml a = new XlsxAppXml();
            a.LinkToPackage(packageObject);
            a.SetSheetCount(worksheets.aa);

            workbookXml.SetSheetCount(worksheets.aa);

            sharedStrings.close();

            packageObject.Close();
        }

        internal XlsxParts(string path)
        {
            packageObject = Package.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            worksheets = new XlsxWorksheetPartCollection();

            #region application/xml

            // here to start the pkg?
            Uri u1_1 = new Uri("/xl/unused.xml", UriKind.Relative);
            PackagePart p1_1 = packageObject.CreatePart(u1_1, "application/xml", CompressionOption.Normal);

            #endregion

            #region workbook.xml_Content_Type

            workbookXml = new XlsxWorkbookMetadata();
            workbookXmlPackagePart = workbookXml.LinkToPackage(packageObject);

            #endregion

            #region core.xml

            XlsxPart core = new XlsxPart(Settings.Default.CoreXml);
            core.LinkToPackage(packageObject, new Uri("/docProps/core.xml", UriKind.Relative)
                , "application/vnd.openxmlformats-package.core-properties+xml"
                , "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                );
            core.close();

            #endregion

            #region theme

            Uri u7 = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
            PackagePart p7 = packageObject.CreatePart(u7, "application/vnd.openxmlformats-officedocument.theme+xml", CompressionOption.Normal);
            using (MemoryStream fs = new MemoryStream())
            {
                Settings.Default.ThemeXml.Save(fs);
                fs.Flush();
                fs.Seek(0, SeekOrigin.Begin);
                StaticFunctions.copyStream(fs, p7.GetStream());
            }
            workbookXmlPackagePart.CreateRelationship(new Uri("theme/theme1.xml", UriKind.Relative), TargetMode.Internal
                , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme");

            #endregion

            #region styles

            Uri u8 = new Uri("/xl/styles.xml", UriKind.Relative);
            PackagePart p8 = packageObject.CreatePart(u8, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", CompressionOption.Normal);
            using (MemoryStream fs = new MemoryStream())
            {
                Settings.Default.StylesXml.Save(fs);
                fs.Flush();
                fs.Seek(0, SeekOrigin.Begin);
                StaticFunctions.copyStream(fs, p8.GetStream());
            }
            workbookXmlPackagePart.CreateRelationship(new Uri("styles.xml", UriKind.Relative), TargetMode.Internal
                , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");

            #endregion

            #region sharedstrings

            sharedStrings = new XlsxSharedStringsXml(packageObject, workbookXmlPackagePart);

            #endregion
        }

        public void WriteRow(IDataReader queryReader)
        {
            currentWorksheet.writerow(queryReader, sharedStrings);
        }
    }
}
