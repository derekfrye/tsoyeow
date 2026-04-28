using System.Text.Json;
using ExcelXmlQueryResults.App.Models;

namespace ExcelXmlQueryResults.App.Services;

public sealed class AppSettingsStore
{
    static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
    };

    readonly LegacySettingsImporter legacyImporter = new();

    public string SettingsPath { get; }

    public AppSettingsStore()
    {
        var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        SettingsPath = Path.Combine(appDataPath, "tsoyeow", "settings.json");
    }

    public async Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        if (File.Exists(SettingsPath))
        {
            await using var stream = File.OpenRead(SettingsPath);
            var settings = await JsonSerializer.DeserializeAsync<AppSettings>(stream, SerializerOptions, cancellationToken);
            return settings ?? new AppSettings();
        }

        var importedSettings = TryImportLegacySettings() ?? new AppSettings();
        await SaveAsync(importedSettings, cancellationToken);
        return importedSettings;
    }

    public async Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(settings);

        var directory = Path.GetDirectoryName(SettingsPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(SettingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, SerializerOptions, cancellationToken);
    }

    public void Save(AppSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        var directory = Path.GetDirectoryName(SettingsPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using var stream = File.Create(SettingsPath);
        JsonSerializer.Serialize(stream, settings, SerializerOptions);
    }

    AppSettings? TryImportLegacySettings()
    {
        foreach (var candidate in GetLegacyCandidates())
        {
            if (legacyImporter.TryImport(candidate, out var imported) && imported is not null)
            {
                return imported;
            }
        }

        return null;
    }

    static IEnumerable<string> GetLegacyCandidates()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "ExcelXmlQueryResults.exe.config");
        yield return Path.Combine(AppContext.BaseDirectory, "ExcelXmlQueryResults.dll.config");
        yield return Path.Combine(AppContext.BaseDirectory, "App.config");

        var repoConfig = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "ExcelXmlQueryResults", "App.config"));
        yield return repoConfig;
    }
}
