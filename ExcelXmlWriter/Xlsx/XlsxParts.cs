using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using ExcelXmlWriter.Properties;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter.Xlsx
{
    internal sealed class XlsxParts : IExcelBackend, IDisposable
    {
        readonly Stream output;
        readonly Relationships mainRelationships;
        readonly Relationships workbookRelationships;
        readonly XlsxWorkbookMetadata workbookMetadata;
        readonly XlsxSharedStringsXml sharedStrings;
        readonly List<XlsxWorksheet> worksheets = [];
        readonly string coreXml;
        readonly string themeXml;
        readonly string stylesXml;

        bool closed;
        XlsxWorksheet? currentWorksheet;

        internal XlsxParts(Stream outputStream)
        {
            output = outputStream;
            mainRelationships = new Relationships();
            workbookRelationships = new Relationships();

            workbookMetadata = new XlsxWorkbookMetadata();
            mainRelationships.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "xl/workbook.xml"));

            coreXml = SerializeXmlDocument(Settings.Default.CoreXml.OuterXml);
            themeXml = SerializeXmlDocument(Settings.Default.ThemeXml.OuterXml);
            stylesXml = SerializeXmlDocument(Settings.Default.StylesXml.OuterXml);

            mainRelationships.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"));
            workbookRelationships.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml"));
            workbookRelationships.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml"));

            sharedStrings = new XlsxSharedStringsXml(new TempFileBuffer());
            workbookRelationships.Link(XlsxPart.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "sharedStrings.xml"));
        }

        public long FileSize
        {
            get
            {
                long total = sharedStrings.OutputStream.Length;
                foreach (var worksheet in worksheets)
                {
                    total += worksheet.OutputStream.Length;
                }

                return total;
            }
        }

        public void CreateSheet(int sheetCount, int subSheetCount, string sheetName, ResultSetSchema schema)
        {
            string sheetPath = "worksheets/sheet" + sheetCount.ToString(CultureInfo.InvariantCulture) + "_" + subSheetCount.ToString(CultureInfo.InvariantCulture) + ".xml";
            var relationship = new ContentRelationships
            {
                PackagePath = sheetPath,
                RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            };
            workbookRelationships.Link(relationship);
            string relationshipId = workbookRelationships.Id(relationship);

            currentWorksheet = new XlsxWorksheet(sheetName, relationshipId, schema, sharedStrings, "/xl/" + sheetPath);
            worksheets.Add(currentWorksheet);
        }

        public void WriteRow(IDataRecord data)
        {
            if (currentWorksheet is null)
            {
                throw new InvalidOperationException("A worksheet must be created before rows can be written.");
            }

            currentWorksheet.WriteRow(data);
        }

        public void CloseSheet()
        {
            currentWorksheet?.Close();
            currentWorksheet = null;
        }

        public void Close()
        {
            if (closed)
            {
                return;
            }

            var appXml = new XlsxAppXml();
            appXml.SetSheetCount(worksheets);
            mainRelationships.Link(XlsxAppXml.LinkToPackage("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml"));
            workbookMetadata.SetSheetCount(worksheets);
            sharedStrings.Close();

            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            {
                AddStringEntry(archive, "_rels/.rels", mainRelationships.Write());
                AddStringEntry(archive, "/docProps/core.xml", coreXml);
                AddStringEntry(archive, "/docProps/app.xml", appXml.Write());
                AddStringEntry(archive, "/[Content_Types].xml", new ContentTypes(worksheets).Write());
                AddStringEntry(archive, "/xl/workbook.xml", workbookMetadata.Write());
                AddStringEntry(archive, "/xl/_rels/workbook.xml.rels", workbookRelationships.Write());
                AddStringEntry(archive, "/xl/theme/theme1.xml", themeXml);
                AddStringEntry(archive, "/xl/styles.xml", stylesXml);
                AddStreamEntry(archive, "/xl/sharedStrings.xml", sharedStrings.OutputStream);

                foreach (var worksheet in worksheets)
                {
                    AddStreamEntry(archive, worksheet.PackageFileName, worksheet.OutputStream);
                }
            }

            foreach (var worksheet in worksheets)
            {
                worksheet.Dispose();
            }

            sharedStrings.Dispose();
            closed = true;
        }

        public void Dispose()
        {
            if (!closed)
            {
                Close();
            }
        }

        static string SerializeXmlDocument(string xmlDocument)
        {
            var document = XDocument.Parse(xmlDocument, LoadOptions.None);
            using var stringWriter = new StringWriterWithEncoding(Encoding.UTF8);
            var settings = new XmlWriterSettings { Encoding = Encoding.UTF8 };
            using var xmlWriter = XmlWriter.Create(stringWriter, settings);
            document.Save(xmlWriter);
            xmlWriter.Close();
            return stringWriter.ToString();
        }

        static string NormalizeEntryPath(string entryPath)
        {
            return entryPath.TrimStart('/');
        }

        static void AddStringEntry(ZipArchive archive, string entryPath, string content)
        {
            var entry = archive.CreateEntry(NormalizeEntryPath(entryPath), CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }

        static void AddStreamEntry(ZipArchive archive, string entryPath, Stream source)
        {
            var entry = archive.CreateEntry(NormalizeEntryPath(entryPath), CompressionLevel.Optimal);
            using var destination = entry.Open();
            source.CopyTo(destination);
        }
    }
}
