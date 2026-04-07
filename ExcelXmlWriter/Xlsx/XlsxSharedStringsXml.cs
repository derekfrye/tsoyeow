using System.Xml;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /xl/sharedStrings.xml
    /// </summary>
    internal sealed class XlsxSharedStringsXml : IDisposable
    {
        const int MaxCachedSharedStrings = 100_000;

        readonly TempFileBuffer buffer;
        readonly XmlWriter writer;
        readonly Dictionary<string, long> sharedStringCache = new(StringComparer.Ordinal);

        long currentSharedStringPosition;

        internal XlsxSharedStringsXml(TempFileBuffer tempBuffer)
        {
            buffer = tempBuffer;

            writer = XmlWriter.Create(OutputStream);
            writer.WriteStartDocument(true);
            writer.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteWhitespace(Environment.NewLine);
        }

        internal Stream OutputStream => buffer.Stream;

        internal static ContentRelationships LinkToPackage(string relationshipType, string packagePath)
        {
            return new ContentRelationships { PackagePath = packagePath, RelationshipType = relationshipType };
        }

        internal long Write(string value)
        {
            var sanitizedValue = string.Concat(value.Where(XmlConvert.IsXmlChar));
            if (sharedStringCache.TryGetValue(sanitizedValue, out var position))
            {
                return position;
            }

            writer.WriteStartElement("si");
            writer.WriteStartElement("t");
            writer.WriteString(sanitizedValue);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteWhitespace(Environment.NewLine);

            var writtenPosition = currentSharedStringPosition;
            currentSharedStringPosition++;

            if (sharedStringCache.Count < MaxCachedSharedStrings)
            {
                sharedStringCache[sanitizedValue] = writtenPosition;
            }

            return writtenPosition;
        }

        internal void Close()
        {
            writer.WriteEndElement();
            writer.Close();
            buffer.Rewind();
        }

        public void Dispose()
        {
            writer.Dispose();
            buffer.Dispose();
        }
    }
}
