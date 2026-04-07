using System.Xml.Linq;
using ExcelXmlQueryResults.App.Models;

namespace ExcelXmlQueryResults.App.Services;

public sealed class LegacySettingsImporter
{
    public bool TryImport(string path, out AppSettings? settings)
    {
        settings = null;

        if (!File.Exists(path))
        {
            return false;
        }

        var document = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        var root = document.Element("configuration");
        var appSettings = root?.Element("appSettings");
        if (root is null || appSettings is null)
        {
            return false;
        }

        settings = new AppSettings
        {
            Server = ReadValue(appSettings, "Server"),
            Database = ReadValue(appSettings, "Database"),
            ConnectionMode = string.Equals(ReadValue(appSettings, "ConnectionMethod"), "SSPI", StringComparison.OrdinalIgnoreCase)
                ? ConnectionMode.IntegratedSecurity
                : ConnectionMode.UsernamePassword,
            Username = ReadValue(appSettings, "ConnectionUsername"),
            Password = ReadValue(appSettings, "ConnectionPassword"),
            QueryTimeoutSeconds = ParseInt(ReadValue(appSettings, "QueryTimeout"), 0),
            MaxRowsPerWorksheet = ParseInt(ReadValue(appSettings, "MaxRowsPerSheet"), 150_000),
            MaxWorkbookBytes = ParseLong(ReadValue(appSettings, "MaximumWorkbookSizeInBytes"), 1_073_741_824),
            MaxResultSetsPerWorkbook = ParseInt(ReadValue(appSettings, "MaximumResultSetsPerWorkbook"), 99),
            IncludeHeadersForEmptyResults = ParseBool(ReadValue(appSettings, "WriteEmptyResultColumnHeaders"), true),
            RewriteOverpunchNumbers = ParseBool(ReadValue(appSettings, "AutoRewriteOverpunch"), true),
            WorksheetBoundaryColumns = ReadBoundaryColumns(root),
            ResultSetNames = ReadResultSetNames(root),
        };

        return true;
    }

    static string ReadValue(XElement appSettings, string key)
    {
        return appSettings
            .Elements("add")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("key"), key, StringComparison.Ordinal))?
            .Attribute("value")?
            .Value
            ?? string.Empty;
    }

    static List<string> ReadBoundaryColumns(XElement root)
    {
        return root
            .Descendants("ColumnsThatPreventNewWorksheets")
            .Elements("entry")
            .Select(entry => entry.Element("value")?.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Cast<string>()
            .ToList();
    }

    static Dictionary<int, string> ReadResultSetNames(XElement root)
    {
        var result = new Dictionary<int, string>();
        foreach (var entry in root.Descendants("ResultNames").Elements("entry"))
        {
            if (int.TryParse(entry.Element("key")?.Value, out var key))
            {
                var value = entry.Element("value")?.Value;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    result[key] = value;
                }
            }
        }

        return result;
    }

    static int ParseInt(string value, int fallback)
    {
        return int.TryParse(value, out var parsed) ? parsed : fallback;
    }

    static long ParseLong(string value, long fallback)
    {
        return long.TryParse(value, out var parsed) ? parsed : fallback;
    }

    static bool ParseBool(string value, bool fallback)
    {
        return bool.TryParse(value, out var parsed) ? parsed : fallback;
    }
}
