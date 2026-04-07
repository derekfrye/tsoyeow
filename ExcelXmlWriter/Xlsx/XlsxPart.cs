using System.Text;
using System.Xml;
using System.Xml.Linq;

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

        internal static string Write(XDocument appXml)
        {
            using (StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8))
            {
                var za = new XmlWriterSettings();
                za.Encoding = Encoding.UTF8;

                using (XmlWriter apo = XmlWriter.Create(sb, za))
                {
                    appXml.Save(apo);
                    apo.Close();
                    return sb.ToString();
                }

            }
        }

        internal string Write()
        {
            return Write(appXml);
        }
    }
}
