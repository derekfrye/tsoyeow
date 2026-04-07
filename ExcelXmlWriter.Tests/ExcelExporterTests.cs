using System.Globalization;
using ExcelXmlWriter.Tests.TestData;
using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class ExcelExporterTests
{
    [Fact]
    public async Task ExportAsync_rewrites_missing_extension_and_emits_expected_sheet_names()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Data.xml");
        var destinationStem = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}");
        var exporter = new ExcelExporter();
        var request = new ExportRequest
        {
            QueryText = fixturePath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
        };
        request.ResultSetNames[1] = "blah blah";
        request.ResultSetNames[2] = "x";

        var result = await exporter.ExportAsync(request, destinationStem);

        result.TotalRecordsRead.Should().Be(7);
        result.Workbooks.Should().ContainSingle();
        result.Workbooks[0].Path.Should().Be(destinationStem + ".xlsx");
        File.Exists(result.Workbooks[0].Path).Should().BeTrue();

        using var archive = new WorkbookArchive(result.Workbooks[0].Path);
        archive.EntryNames.Should().Contain("xl/worksheets/sheet1_1.xml");
        archive.EntryNames.Should().Contain("xl/worksheets/sheet2_1.xml");

        var workbookXml = archive.ReadXml("xl/workbook.xml");
        var sheetNames = workbookXml.Root!
            .Descendants()
            .Where(element => element.Name.LocalName == "sheet")
            .Select(element => element.Attribute("name")!.Value)
            .ToArray();

        sheetNames.Should().ContainInOrder("blah blah_1", "x_1");
        archive.CountRows("xl/worksheets/sheet1_1.xml").Should().Be(3);
        archive.CountRows("xl/worksheets/sheet2_1.xml").Should().Be(6);
    }

    [Fact]
    public async Task ExportAsync_splits_across_multiple_workbooks_when_result_set_limit_is_hit()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id", "Value"], CreateRows(2, index => [(index + 1).ToString(CultureInfo.InvariantCulture), $"order-{index + 1}"])),
            ("Invoices", ["Id", "Value"], CreateRows(2, index => [(index + 1).ToString(CultureInfo.InvariantCulture), $"invoice-{index + 1}"])),
            ("Payments", ["Id", "Value"], CreateRows(1, index => [(index + 1).ToString(CultureInfo.InvariantCulture), $"payment-{index + 1}"])));
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        var exporter = new ExcelExporter();
        var request = new ExportRequest
        {
            QueryText = xmlPath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
            MaxResultSetsPerWorkbook = 2,
        };

        var result = await exporter.ExportAsync(request, destinationPath);

        result.TotalRecordsRead.Should().Be(5);
        result.Workbooks.Select(workbook => workbook.SequenceNumber).Should().Equal(1, 2);
        result.Workbooks.Select(workbook => Path.GetFileName(workbook.Path)).Should().Equal(
            Path.GetFileName(destinationPath),
            $"{Path.GetFileNameWithoutExtension(destinationPath)}_2.xlsx");

        using (var firstWorkbook = new WorkbookArchive(result.Workbooks[0].Path))
        {
            firstWorkbook.EntryNames.Should().Contain("xl/worksheets/sheet1_1.xml");
            firstWorkbook.EntryNames.Should().Contain("xl/worksheets/sheet2_1.xml");
            firstWorkbook.CountRows("xl/worksheets/sheet1_1.xml").Should().Be(3);
            firstWorkbook.CountRows("xl/worksheets/sheet2_1.xml").Should().Be(3);
        }

        using var secondWorkbook = new WorkbookArchive(result.Workbooks[1].Path);
        secondWorkbook.EntryNames.Should().Contain("xl/worksheets/sheet3_1.xml");
        secondWorkbook.CountRows("xl/worksheets/sheet3_1.xml").Should().Be(2);
    }

    [Fact]
    public async Task ExportAsync_splits_rows_across_multiple_worksheets_when_row_limit_is_hit()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id", "Value"], CreateRows(5, index => [(index + 1).ToString(CultureInfo.InvariantCulture), $"row-{index + 1}"])));
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        var exporter = new ExcelExporter();
        var request = new ExportRequest
        {
            QueryText = xmlPath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
            MaxRowsPerWorksheet = 2,
        };

        var result = await exporter.ExportAsync(request, destinationPath);

        using var workbook = new WorkbookArchive(result.Workbooks[0].Path);
        workbook.EntryNames.Should().Contain("xl/worksheets/sheet1_1.xml");
        workbook.EntryNames.Should().Contain("xl/worksheets/sheet1_2.xml");
        workbook.EntryNames.Should().Contain("xl/worksheets/sheet1_3.xml");
        workbook.CountRows("xl/worksheets/sheet1_1.xml").Should().Be(2);
        workbook.CountRows("xl/worksheets/sheet1_2.xml").Should().Be(3);
        workbook.CountRows("xl/worksheets/sheet1_3.xml").Should().Be(3);
    }

    [Fact]
    public async Task ExportAsync_splits_into_multiple_workbooks_when_maximum_size_is_small()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id", "Value"], CreateRows(40, index => [(index + 1).ToString(CultureInfo.InvariantCulture), $"{index:D4}-{new string((char)('a' + (index % 26)), 200)}"])));
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        var exporter = new ExcelExporter();
        var request = new ExportRequest
        {
            QueryText = xmlPath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
            MaxWorkbookBytes = 1_000,
        };

        var result = await exporter.ExportAsync(request, destinationPath);

        result.Workbooks.Count.Should().BeGreaterThan(1);

        var totalRowCount = result.Workbooks.Sum(workbookPath =>
        {
            using var workbook = new WorkbookArchive(workbookPath.Path);
            return workbook.EntryNames
                .Where(entryName => entryName.StartsWith("xl/worksheets/", StringComparison.Ordinal))
                .Sum(worksheet => workbook.CountRows(worksheet) - 1);
        });

        totalRowCount.Should().Be(40);
    }

    [Fact]
    public async Task ExportAsync_reports_query_start_save_and_completed_progress()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id"], CreateRows(1, index => [(index + 1).ToString(CultureInfo.InvariantCulture)])));
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        var exporter = new ExcelExporter();
        var progressEvents = new List<ExportProgress>();
        var request = new ExportRequest
        {
            QueryText = xmlPath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
        };

        await exporter.ExportAsync(request, destinationPath, new Progress<ExportProgress>(progress => progressEvents.Add(progress)));

        progressEvents.Select(progress => progress.Stage).Should().ContainInOrder(
            ExportStage.QueryStarting,
            ExportStage.SavingWorkbook,
            ExportStage.Completed);
        progressEvents.Last().TotalRows.Should().Be(1);
    }

    [Fact]
    public async Task ExportAsync_throws_operation_canceled_when_token_is_already_canceled()
    {
        var xmlPath = TestDataSetFactory.CreateXmlFile(
            ("Orders", ["Id"], CreateRows(1, index => [(index + 1).ToString(CultureInfo.InvariantCulture)])));
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        var exporter = new ExcelExporter();
        var request = new ExportRequest
        {
            QueryText = xmlPath,
            SourceKind = ExportSourceKind.XmlDataSetFile,
        };
        using var cancellationTokenSource = new CancellationTokenSource();
        cancellationTokenSource.Cancel();

        Func<Task> act = async () => await exporter.ExportAsync(request, destinationPath, cancellationToken: cancellationTokenSource.Token);

        await act.Should().ThrowAsync<OperationCanceledException>();
    }

    [Fact]
    public async Task ExportAsync_rejects_invalid_requests()
    {
        var exporter = new ExcelExporter();
        var destinationPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        await AssertFailure(
            new ExportRequest { QueryText = "", SourceKind = ExportSourceKind.XmlDataSetFile },
            typeof(ArgumentException));
        await AssertFailure(
            new ExportRequest { QueryText = "query", SourceKind = ExportSourceKind.SqlQuery, ConnectionString = "" },
            typeof(ArgumentException));
        await AssertFailure(
            new ExportRequest { QueryText = "query", SourceKind = ExportSourceKind.XmlDataSetFile, MaxRowsPerWorksheet = 0 },
            typeof(ArgumentOutOfRangeException));
        await AssertFailure(
            new ExportRequest { QueryText = "query", SourceKind = ExportSourceKind.XmlDataSetFile, MaxWorkbookBytes = 0 },
            typeof(ArgumentOutOfRangeException));
        await AssertFailure(
            new ExportRequest { QueryText = "query", SourceKind = ExportSourceKind.XmlDataSetFile, MaxResultSetsPerWorkbook = 0 },
            typeof(ArgumentOutOfRangeException));

        async Task AssertFailure(ExportRequest request, Type exceptionType)
        {
            Func<Task> act = async () => await exporter.ExportAsync(request, destinationPath);
            var exception = await act.Should().ThrowAsync<Exception>();
            exception.Which.Should().BeOfType(exceptionType);
        }
    }

    static object?[][] CreateRows(int count, Func<int, object?[]> factory)
    {
        return Enumerable.Range(0, count).Select(factory).ToArray();
    }
}
