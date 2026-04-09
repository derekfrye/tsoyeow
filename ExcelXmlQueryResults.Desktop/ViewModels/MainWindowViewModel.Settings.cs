using System.Globalization;
using ExcelXmlQueryResults.App.Models;

namespace ExcelXmlQueryResults.Desktop.ViewModels;

public sealed partial class MainWindowViewModel
{
    void ApplySettings(AppSettings settings)
    {
        Server = settings.Server;
        Database = settings.Database;
        UseIntegratedSecurity = settings.ConnectionMode == ConnectionMode.IntegratedSecurity;
        Username = settings.Username;
        Password = settings.Password;
        QueryTimeoutSeconds = settings.QueryTimeoutSeconds.ToString(CultureInfo.InvariantCulture);
        MaxRowsPerWorksheet = settings.MaxRowsPerWorksheet.ToString(CultureInfo.InvariantCulture);
        MaxWorkbookSizeGiB = ((decimal)settings.MaxWorkbookBytes / (1024 * 1024 * 1024)).ToString("0.###", CultureInfo.InvariantCulture);
        MaxResultSetsPerWorkbook = settings.MaxResultSetsPerWorkbook.ToString(CultureInfo.InvariantCulture);
        BoundaryColumns = string.Join(", ", settings.WorksheetBoundaryColumns);
        ResultSetNames = string.Join(", ", settings.ResultSetNames.OrderBy(pair => pair.Key).Select(pair => $"{pair.Key}={pair.Value}"));
        OutputPath = settings.OutputPath;
    }

    AppSettings CreateSettingsSnapshot()
    {
        return new AppSettings
        {
            Server = Server.Trim(),
            Database = Database.Trim(),
            ConnectionMode = UseIntegratedSecurity ? ConnectionMode.IntegratedSecurity : ConnectionMode.UsernamePassword,
            Username = Username.Trim(),
            Password = Password,
            QueryTimeoutSeconds = ParseInt(QueryTimeoutSeconds, nameof(QueryTimeoutSeconds)),
            MaxRowsPerWorksheet = ParseInt(MaxRowsPerWorksheet, nameof(MaxRowsPerWorksheet)),
            MaxWorkbookBytes = ParseWorkbookBytes(MaxWorkbookSizeGiB),
            MaxResultSetsPerWorkbook = ParseInt(MaxResultSetsPerWorkbook, nameof(MaxResultSetsPerWorkbook)),
            IncludeHeadersForEmptyResults = true,
            RewriteOverpunchNumbers = true,
            OutputPath = OutputPath.Trim(),
            WorksheetBoundaryColumns = ParseCsvList(BoundaryColumns),
            ResultSetNames = ParseResultSetNames(ResultSetNames),
        };
    }

    static int ParseInt(string value, string fieldName)
    {
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) || parsed < 0)
        {
            throw new InvalidOperationException($"{fieldName} must be a non-negative whole number.");
        }

        return parsed;
    }

    static long ParseWorkbookBytes(string gibibytes)
    {
        if (!decimal.TryParse(gibibytes, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed) || parsed <= 0)
        {
            throw new InvalidOperationException("Max workbook size must be a positive number of GiB.");
        }

        return decimal.ToInt64(decimal.Round(parsed * 1024 * 1024 * 1024, 0, MidpointRounding.AwayFromZero));
    }

    static List<string> ParseCsvList(string text)
    {
        return text
            .Split([',', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
    }

    static Dictionary<int, string> ParseResultSetNames(string text)
    {
        var result = new Dictionary<int, string>();
        var entries = text.Split([',', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var entry in entries)
        {
            var parts = entry.Split('=', 2, StringSplitOptions.TrimEntries);
            if (parts.Length != 2 || !int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var key))
            {
                throw new InvalidOperationException("Result set names must use the format '1=Sheet Name, 2=Another Name'.");
            }

            if (!string.IsNullOrWhiteSpace(parts[1]))
            {
                result[key] = parts[1];
            }
        }

        return result;
    }
}
