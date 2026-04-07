using System.Globalization;
using System.Xml.Linq;

namespace ExcelXmlWriter.Xlsx
{
    internal static class RelationshipIdGenerator
    {
        static readonly object SyncRoot = new();
        static int currentId;

        internal static string RetrieveNextId()
        {
            lock (SyncRoot)
            {
                currentId++;
                return "rId" + currentId.ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    internal sealed class ContentRelationships
    {
        internal string PackagePath { get; set; } = string.Empty;

        internal string RelationshipType { get; set; } = string.Empty;
    }

    internal sealed class Relationships
    {
        readonly Dictionary<string, ContentRelationships> relationships = new();

        internal string Id(ContentRelationships relationship)
        {
            return relationships.First(entry => ReferenceEquals(entry.Value, relationship)).Key;
        }

        internal void Link(ContentRelationships relationship)
        {
            relationships.Add(RelationshipIdGenerator.RetrieveNextId(), relationship);
        }

        internal string Write()
        {
            XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

            var root = new XElement(relationshipsNamespace + "Relationships");
            foreach (var relationship in relationships)
            {
                root.Add(
                    new XElement(
                        relationshipsNamespace + "Relationship",
                        new XAttribute("Type", relationship.Value.RelationshipType),
                        new XAttribute("Target", relationship.Value.PackagePath),
                        new XAttribute("Id", relationship.Key)));
            }

            var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
            return XlsxPart.Write(document);
        }
    }
}
