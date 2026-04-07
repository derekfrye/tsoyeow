namespace ExcelXmlWriter;

public enum ExportStage
{
    QueryStarting,
    QueryRunning,
    SavingWorkbook,
    Completed,
}

public sealed class ExportProgress
{
    public ExportStage Stage { get; init; }

    public decimal? RowsPerSecond { get; init; }

    public int TotalRows { get; init; }

    public string? Message { get; init; }

    public string? OutputPath { get; init; }
}
