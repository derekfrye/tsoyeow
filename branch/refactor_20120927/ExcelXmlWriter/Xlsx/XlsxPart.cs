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

        //protected PackagePart p3;
        protected XDocument appXml = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));

        internal XlsxPart()
        {
        }

        internal XlsxPart(XmlDocument x)
        {
            appXml = XDocument.Parse(x.OuterXml);
        }

        internal ZipAAA LinkToPackage(string reltp,string pth)
        {
            //p3 = p.CreatePart(u, contentType, CompressionOption.SuperFast);
            //p.CreateRelationship(u, TargetMode.Internal, relationshipType);
            
                //string contentType= "application/vnd.openxmlformats-package.core-properties+xml";
                //string relationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
                return new ZipAAA() { path = pth, RelType = reltp };
            //return p3;
        }

        //internal void close()
        //{        	
        //    XmlWriter xxx = XmlWriter.Create(p3.GetStream());
        //    appXml.Save(xxx);
        //    xxx.Close();
        //    p3.GetStream().Close();
        //}
    }
}
