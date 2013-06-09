using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;

namespace ExcelXmlWriter.Xlsx
{
    static class RelationshipIdGenerator
    {
        static readonly object ll = new object();
        static int i = 0;
        internal static string RetrieveNextId()
        {
            lock (ll)
            {
                i = i + 1;

            }
            // Excel must have rId prefixed to the Id.
            return "rId" + i.ToString();
        }
    }

    class ContentRelationships
    {
        internal string PackagePath
        { get; set; }
        internal string RelationshipType
        { get; set; }
    }

    class Relationships
    {
        Dictionary<string, ContentRelationships> rels;

        internal Relationships()
        {
            rels = new Dictionary<string, ContentRelationships>();
        }

        internal string Id(ContentRelationships x)
        {
            return rels.First(xx => xx.Value == x).Key;
        }

        internal void Link(ContentRelationships x)
        {
            var t = RelationshipIdGenerator.RetrieveNextId();
            rels.Add(t, x);
        }

        internal string Write()
        {
            XNamespace xn2 = "http://schemas.openxmlformats.org/package/2006/relationships";

            var t = new XElement(xn2 + "Relationships");
            foreach (var z in rels)
            {
                t.Add(new XElement(xn2 + "Relationship", new XAttribute("Type", z.Value.RelationshipType), new XAttribute("Target", z.Value.PackagePath), new XAttribute("Id", z.Key.ToString())));
            }
            XDocument x = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), t);

            return XlsxPart.Write(x);
        }

    }
}
