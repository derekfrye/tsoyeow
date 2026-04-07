using System.Text;
using System.Xml.Linq;
using ExcelXmlWriter.Xlsx;
using FluentAssertions;

namespace ExcelXmlWriter.Tests;

public sealed class XlsxSharedStringsXmlTests
{
    [Fact]
    public void Write_deduplicates_strings_and_strips_invalid_xml_characters()
    {
        using var tempBuffer = new TempFileBuffer();
        using var sharedStrings = new XlsxSharedStringsXml(tempBuffer);

        var firstId = sharedStrings.Write("ab\u0001c");
        var secondId = sharedStrings.Write("abc");
        sharedStrings.Close();

        firstId.Should().Be(0);
        secondId.Should().Be(0);

        using var reader = new StreamReader(sharedStrings.OutputStream, Encoding.UTF8, leaveOpen: true);
        var xml = XDocument.Parse(reader.ReadToEnd(), LoadOptions.None);
        var values = xml.Root!
            .Elements()
            .Where(element => element.Name.LocalName == "si")
            .Select(element => element.Descendants().First(node => node.Name.LocalName == "t").Value)
            .ToArray();

        values.Should().Equal("abc");
    }
}
