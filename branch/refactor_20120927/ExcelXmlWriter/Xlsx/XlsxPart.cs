using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Xml;

namespace ExcelXmlWriter.Xlsx
{
    class XlsxPart
    {
        protected XDocument appXml = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

        internal XlsxPart()
        {
        }

        internal XlsxPart(XmlDocument inputDocument)
        {
            appXml = XDocument.Parse(inputDocument.OuterXml);
        }

        internal static ContentRelationships LinkToPackage(string relationshipType, string packagePath)
        {
            return new ContentRelationships() { PackagePath = packagePath, RelationshipType = relationshipType };
        }

        internal string Write()
        {
            StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8);

            var za = new XmlWriterSettings();
            za.Encoding = Encoding.UTF8;

            XmlWriter apo = XmlWriter.Create(sb, za);
            appXml.Save(apo);
            apo.Close();
            return sb.ToString();
        }
    }
}