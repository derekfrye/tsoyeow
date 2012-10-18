using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;

namespace ExcelXmlWriter.Xlsx
{
    public sealed class StringWriterWithEncoding : StringWriter
    {
        private readonly Encoding encoding;

        public StringWriterWithEncoding(Encoding encoding)
        {
            this.encoding = encoding;
        }

        public override Encoding Encoding
        {
            get { return encoding; }
        }
    }

    static class RelInt
    {
        static readonly object ll = new object();
        static int i = 0;
        public static string A()
        {
            lock (ll)
            {
                i = i + 1;
                
            }
            return "rId"+i.ToString();
        }
    }

    class Rels
    {
        Dictionary<string,ZipAAA> rels;
        //int i = 0;

        public Rels()
        {
            rels = new Dictionary<string, ZipAAA>();
        }

        public string Id(ZipAAA x)
        {
            return rels.First(xx => xx.Value == x).Key;
        }

        public void Link(ZipAAA x)
        {
            //i = i + 1;
            var t = RelInt.A();
            rels.Add(t,x);
        }

        public string Write()
        {
            XNamespace xn2 = "http://schemas.openxmlformats.org/package/2006/relationships";

            //<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            
            var t = new XElement(xn2 + "Relationships");
            foreach (var z in rels)
            {
                t.Add(new XElement(xn2 + "Relationship", new XAttribute("Type", z.Value.RelType), new XAttribute("Target", z.Value.path), new XAttribute("Id", z.Key.ToString())));
            }
            XDocument x = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), t);

            StringWriterWithEncoding sb = new StringWriterWithEncoding(Encoding.UTF8);
            
            var za = new XmlWriterSettings();
            za.Encoding = Encoding.UTF8;
            
            XmlWriter apo = XmlWriter.Create(sb, za);
            x.Save(apo);
            apo.Close();
            return sb.ToString();
        }
    
    }
}
