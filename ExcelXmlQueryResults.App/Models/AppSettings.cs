namespace ExcelXmlQueryResults.App.Models;

public enum ConnectionMode
{
    IntegratedSecurity,
    UsernamePassword,
}

public sealed class AppSettings
{
    public string Server { get; set; } = string.Empty;

    public string Database { get; set; } = string.Empty;

    public ConnectionMode ConnectionMode { get; set; } = ConnectionMode.IntegratedSecurity;

    public string Username { get; set; } = string.Empty;

    public string Password { get; set; } = string.Empty;

    public int QueryTimeoutSeconds { get; set; }

    public int MaxRowsPerWorksheet { get; set; } = 150_000;

    public long MaxWorkbookBytes { get; set; } = 1_073_741_824;

    public int MaxResultSetsPerWorkbook { get; set; } = 99;

    public bool IncludeHeadersForEmptyResults { get; set; } = true;

    public bool RewriteOverpunchNumbers { get; set; } = true;

    public string OutputPath { get; set; } = string.Empty;

    public List<string> WorksheetBoundaryColumns { get; set; } = [];

    public Dictionary<int, string> ResultSetNames { get; set; } = [];
}
