using System.Data;
using ExcelXmlWriter.Workbook;
using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class ResultSetSchemaTests
{
    [Fact]
    public void Create_returns_null_when_schema_table_is_missing()
    {
        ResultSetSchema.Create(null).Should().BeNull();
    }

    [Fact]
    public void Create_tracks_column_names_case_insensitively()
    {
        var schemaTable = new DataTable();
        schemaTable.Columns.Add("ColumnName", typeof(string));
        schemaTable.Rows.Add("OrderId");
        schemaTable.Rows.Add("CustomerName");

        var schema = ResultSetSchema.Create(schemaTable);

        schema.Should().NotBeNull();
        schema!.ContainsColumn("orderid").Should().BeTrue();
        schema.ContainsColumn("CUSTOMERNAME").Should().BeTrue();
        schema.ContainsColumn("Missing").Should().BeFalse();
    }
}
