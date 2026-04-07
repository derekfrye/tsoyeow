using System.IO.Compression;
using System.Xml.Linq;
using FluentAssertions;

namespace ExcelXmlWriter.Tests.TestData;

internal sealed class WorkbookArchive : IDisposable
{
    readonly ZipArchive archive;

    internal WorkbookArchive(string path)
    {
        archive = ZipFile.OpenRead(path);
    }

    internal IReadOnlyList<string> EntryNames => archive.Entries.Select(entry => entry.FullName).ToArray();

    internal XDocument ReadXml(string entryName)
    {
        var entry = archive.GetEntry(entryName);
        entry.Should().NotBeNull($"entry '{entryName}' should exist in the generated workbook");

        using var stream = entry!.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    internal int CountRows(string sheetEntryName)
    {
        return ReadXml(sheetEntryName)
            .Root!
            .Elements()
            .First(element => element.Name.LocalName == "sheetData")
            .Elements()
            .Count(element => element.Name.LocalName == "row");
    }

    internal string? ReadSharedString(int index)
    {
        var sharedStrings = ReadXml("xl/sharedStrings.xml");
        return sharedStrings.Root!
            .Elements()
            .Where(element => element.Name.LocalName == "si")
            .ElementAt(index)
            .Descendants()
            .First(element => element.Name.LocalName == "t")
            .Value;
    }

    public void Dispose()
    {
        archive.Dispose();
    }
}
