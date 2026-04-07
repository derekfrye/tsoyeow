using ExcelXmlWriter.Tests.TestData;
using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class QueryReaderTests
{
    [Fact]
    public void OpenReader_reads_all_rows_across_multiple_tables_from_xml_fixture()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id", "Name"], [["1", "alpha"], ["2", "beta"]]),
            ("Invoices", ["Id", "Name"], [["3", "gamma"]]));

        using var reader = new QueryReader(xmlPath, queryTimeout: 30, queryIsFile: true, connStr: string.Empty);

        reader.OpenReader();

        reader.HasActiveResultSet.Should().BeTrue();
        reader.CurrentResult.Should().Be(1);
        reader.Read().Should().BeTrue();
        reader["Id"].Should().Be("1");
        reader.Read().Should().BeTrue();
        reader.Read().Should().BeFalse();
        reader.TotalRecordsRead.Should().Be(2);

        reader.NextResult().Should().BeTrue();
        reader.CurrentResult.Should().Be(2);
        reader.Read().Should().BeTrue();
        reader["Name"].Should().Be("gamma");
        reader.Read().Should().BeFalse();
        reader.TotalRecordsRead.Should().Be(3);
        reader.NextResult().Should().BeFalse();
    }

    [Fact]
    public void OpenReader_leaves_reader_inactive_for_empty_dataset()
    {
        var xmlPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xml");
        File.WriteAllText(xmlPath, "<?xml version=\"1.0\" encoding=\"utf-8\"?><NewDataSet />");

        using var reader = new QueryReader(xmlPath, queryTimeout: 30, queryIsFile: true, connStr: string.Empty);

        reader.OpenReader();

        reader.HasActiveResultSet.Should().BeFalse();
        reader.IsClosed.Should().BeTrue();
    }

    [Fact]
    public void Members_that_require_an_open_reader_throw_clear_state_errors()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(("Orders", ["Id"], [["1"]]));
        using var reader = new QueryReader(xmlPath, queryTimeout: 30, queryIsFile: true, connStr: string.Empty);

        Action act = () => _ = reader["Id"];

        act.Should().Throw<InvalidOperationException>()
            .WithMessage("The query reader has not been opened.");
    }
}
