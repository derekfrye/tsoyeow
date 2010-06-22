using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.Xml.Linq;
using System.IO;
using System.Xml;

namespace ExcelXmlWriter
{
    class XlsxPart
    {

        protected PackagePart p3;
        protected XDocument appXml = new XDocument();

        internal XlsxPart()
        {
        }

        internal XlsxPart(XmlDocument x)
        {
            appXml = XDocument.Parse(x.OuterXml);
        }

        internal PackagePart LinkToPackage(Package p, Uri u, string contentType, string relationshipType)
        {
            p3 = p.CreatePart(u, contentType, CompressionOption.SuperFast);
            p.CreateRelationship(u, TargetMode.Internal, relationshipType);
            return p3;
        }

        internal void close()
        {
            using (MemoryStream fs = new MemoryStream())
            {
                XmlWriter xxx = XmlWriter.Create(fs);
                appXml.Save(xxx);
                xxx.Flush();
                fs.Flush();
                fs.Seek(0, SeekOrigin.Begin);

                StaticFunctions.copyStream(fs, p3.GetStream());
            }
        }
    }
}
