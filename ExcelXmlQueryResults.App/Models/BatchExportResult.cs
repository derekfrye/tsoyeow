using ExcelXmlWriter;

namespace ExcelXmlQueryResults.App.Models;

public sealed class BatchExportResult
{
    public required IReadOnlyList<ExportResult> Exports { get; init; }

    public IReadOnlyList<string> GeneratedFiles =>
        Exports
            .SelectMany(result => result.Workbooks)
            .Select(workbook => workbook.Path)
            .ToArray();
}
