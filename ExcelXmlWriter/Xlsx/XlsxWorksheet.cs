using System.Data;
using System.Globalization;
using System.Xml;
using ExcelXmlWriter.Workbook;

namespace ExcelXmlWriter.Xlsx
{
    internal sealed class XlsxWorksheet : IDisposable
    {
        const int OverpunchPrecision = 100;

        readonly TempFileBuffer buffer;
        readonly XlsxSharedStringsXml sharedStringReference;
        readonly XmlWriter writer;

        internal XlsxWorksheet(string sheetName, string relationshipId, ResultSetSchema schema, XlsxSharedStringsXml sharedStrings, string packageFileName)
        {
            PackageFileName = packageFileName;

            buffer = new TempFileBuffer();
            sharedStringReference = sharedStrings;
            writer = XmlWriter.Create(OutputStream);

            RelationshipId = relationshipId;
            Sheetname = sheetName;

            writer.WriteStartDocument(true);
            writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            writer.WriteAttributeString("xmlns", "mc", null, "http://schemas.openxmlformats.org/markup-compatibility/2006");
            writer.WriteAttributeString("mc", "Ignorable", null, "x14ac");
            writer.WriteAttributeString("xmlns", "x14ac", null, "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            writer.WriteWhitespace(Environment.NewLine);

            writer.WriteStartElement("dimension");
            writer.WriteAttributeString("ref", "A1");
            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);

            writer.WriteStartElement("sheetViews");
            writer.WriteWhitespace(Environment.NewLine);
            writer.WriteStartElement("sheetView");
            writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteWhitespace(Environment.NewLine);

            writer.WriteStartElement("pane");
            writer.WriteAttributeString("ySplit", "1");
            writer.WriteAttributeString("topLeftCell", "A2");
            writer.WriteAttributeString("activePane", "bottomLeft");
            writer.WriteAttributeString("state", "frozen");
            writer.WriteEndElement();

            writer.WriteStartElement("selection");
            writer.WriteAttributeString("pane", "bottomLeft");
            writer.WriteAttributeString("activeCell", "A2");
            writer.WriteAttributeString("sqref", "A2");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);
            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);

            writer.WriteStartElement("sheetFormatPr");
            writer.WriteAttributeString("defaultRowHeight", "15");
            writer.WriteAttributeString("x14ac", "dyDescent", null, "0.25");
            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);

            writer.WriteStartElement("sheetData");
            writer.WriteWhitespace(Environment.NewLine);

            WriteValues(schema.Rows.Cast<DataRow>().Select(row => row["ColumnName"]?.ToString()));
        }

        internal string Sheetname { get; }

        internal string RelationshipId { get; }

        internal Stream OutputStream => buffer.Stream;

        internal string PackageFileName { get; }

        void WriteValues(IEnumerable<string?> values)
        {
            writer.WriteStartElement("row");
            writer.WriteWhitespace(Environment.NewLine);

            foreach (var value in values)
            {
                WriteCellValue(value ?? string.Empty);
            }

            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);
        }

        void WriteCellValue(string value)
        {
            var excelDataType = StaticFunctions.ResolveDataType(value);

            writer.WriteStartElement("c");
            switch (excelDataType)
            {
                case ExcelDataType.Number:
                case ExcelDataType.Date:
                    if (excelDataType == ExcelDataType.Date)
                    {
                        writer.WriteAttributeString("s", "1");
                    }

                    writer.WriteStartElement("v");
                    writer.WriteString(XlsxData.ConvertToWriteableValue(value, excelDataType));
                    break;
                case ExcelDataType.OverpunchNumber:
                    writer.WriteStartElement("v");
                    var overpunch = StaticFunctions.applyOverPunch(value, OverpunchPrecision);
                    writer.WriteString(XlsxData.ConvertToWriteableValue(overpunch.val.ToString(CultureInfo.InvariantCulture), ExcelDataType.Number));
                    break;
                default:
                    long sharedStringId = sharedStringReference.Write(XlsxData.ConvertToWriteableValue(value, excelDataType));
                    writer.WriteAttributeString("t", "s");
                    writer.WriteStartElement("v");
                    writer.WriteString(sharedStringId.ToString(CultureInfo.InvariantCulture));
                    break;
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        internal void WriteRow(IDataRecord record)
        {
            var values = Enumerable.Range(0, record.FieldCount)
                .Select(index => record.GetValue(index)?.ToString());
            WriteValues(values);
        }

        internal void Close()
        {
            writer.WriteEndElement();
            writer.WriteStartElement("pageMargins");
            writer.WriteAttributeString("left", "0.7");
            writer.WriteAttributeString("right", "0.7");
            writer.WriteAttributeString("top", "0.75");
            writer.WriteAttributeString("bottom", "0.75");
            writer.WriteAttributeString("header", "0.3");
            writer.WriteAttributeString("footer", "0.3");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.Close();
            buffer.Rewind();
        }

        public void Dispose()
        {
            writer.Dispose();
            buffer.Dispose();
        }
    }
}
