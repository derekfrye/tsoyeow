namespace ExcelXmlWriter;

public enum ExportSourceKind
{
    SqlQuery,
    XmlDataSetFile,
}

public enum OutputWorkbookFormat
{
    Xlsx,
}

public sealed class ExportRequest
{
    public string QueryText { get; set; } = string.Empty;

    public string ConnectionString { get; set; } = string.Empty;

    public ExportSourceKind SourceKind { get; set; } = ExportSourceKind.SqlQuery;

    public OutputWorkbookFormat Format { get; set; } = OutputWorkbookFormat.Xlsx;

    public int QueryTimeoutSeconds { get; set; } = 30;

    public int MaxRowsPerWorksheet { get; set; } = 150_000;

    public bool IncludeHeadersForEmptyResults { get; set; } = true;

    public long MaxWorkbookBytes { get; set; } = 2_500_000_000;

    public bool RewriteOverpunchNumbers { get; set; }

    public int MaxResultSetsPerWorkbook { get; set; } = int.MaxValue;

    public IList<string> WorksheetBoundaryColumns { get; } = [];

    public IDictionary<int, string> ResultSetNames { get; } = new Dictionary<int, string>();

    public IDictionary<int, ExcelDataType> ColumnTypeMappings { get; } = new Dictionary<int, ExcelDataType>();
}
