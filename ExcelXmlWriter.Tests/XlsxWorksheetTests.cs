using System.Data;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using ExcelXmlWriter.Workbook;
using ExcelXmlWriter.Xlsx;
using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class XlsxWorksheetTests
{
    [Fact]
    public void WriteRow_writes_headers_numbers_dates_and_shared_strings()
    {
        var schemaTable = new DataTable();
        schemaTable.Columns.Add("ColumnName", typeof(string));
        schemaTable.Rows.Add("Id");
        schemaTable.Rows.Add("CreatedAt");
        schemaTable.Rows.Add("Name");
        var schema = ResultSetSchema.Create(schemaTable)!;

        using var sharedStringBuffer = new TempFileBuffer();
        using var sharedStrings = new XlsxSharedStringsXml(sharedStringBuffer);
        using var worksheet = new XlsxWorksheet("Orders_1", "rId1", schema, sharedStrings, "/xl/worksheets/sheet1_1.xml");

        var rowTable = new DataTable();
        rowTable.Columns.Add("Id", typeof(string));
        rowTable.Columns.Add("CreatedAt", typeof(string));
        rowTable.Columns.Add("Name", typeof(string));
        rowTable.Rows.Add("42", "2009-12-31T00:00:00", "alpha");
        using var reader = rowTable.CreateDataReader();
        reader.Read();

        worksheet.WriteRow(reader);
        worksheet.Close();
        sharedStrings.Close();

        using var worksheetReader = new StreamReader(worksheet.OutputStream, Encoding.UTF8, leaveOpen: true);
        var worksheetXml = XDocument.Parse(worksheetReader.ReadToEnd(), LoadOptions.None);
        var rows = worksheetXml.Root!
            .Elements()
            .First(element => element.Name.LocalName == "sheetData")
            .Elements()
            .Where(element => element.Name.LocalName == "row")
            .ToArray();

        rows.Should().HaveCount(2);
        var dataCells = rows[1].Elements().Where(element => element.Name.LocalName == "c").ToArray();
        dataCells[0].Attribute("t").Should().BeNull();
        dataCells[1].Attribute("s")!.Value.Should().Be("1");
        dataCells[2].Attribute("t")!.Value.Should().Be("s");

        var dateValue = dataCells[1].Elements().First(element => element.Name.LocalName == "v").Value;
        dateValue.Should().Be(((Convert.ToDateTime("2009-12-31", CultureInfo.CurrentCulture) - new DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc)).TotalDays).ToString(CultureInfo.InvariantCulture));

        using var sharedStringsReader = new StreamReader(sharedStrings.OutputStream, Encoding.UTF8, leaveOpen: true);
        var sharedStringsXml = XDocument.Parse(sharedStringsReader.ReadToEnd(), LoadOptions.None);
        var sharedValues = sharedStringsXml.Root!
            .Elements()
            .Where(element => element.Name.LocalName == "si")
            .Select(element => element.Descendants().First(node => node.Name.LocalName == "t").Value)
            .ToArray();

        sharedValues.Should().Contain(["Id", "CreatedAt", "Name", "alpha"]);
    }
}
