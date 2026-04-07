namespace ExcelXmlWriter;

public sealed class ExportedWorkbook
{
    public required int SequenceNumber { get; init; }

    public required string Path { get; init; }
}

public sealed class ExportResult
{
    public required int TotalRecordsRead { get; init; }

    public required IReadOnlyList<ExportedWorkbook> Workbooks { get; init; }
}
