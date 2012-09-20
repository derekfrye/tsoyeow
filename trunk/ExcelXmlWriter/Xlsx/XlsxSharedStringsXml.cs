using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;
using System.Collections;
using System.Security.Cryptography;

namespace ExcelXmlWriter
{
    class XlsxSharedStringsXml
    {
        PackagePart sharedStringsPt;
        long ss = 0;
        FileStream fs;
        Crc32 c = new Crc32();
        Dictionary<string, long> a = new Dictionary<string, long>();
        readonly byte[] openString = Encoding.UTF8.GetBytes("<si><t>");
        readonly byte[] closeString = Encoding.UTF8.GetBytes("</t></si>" + Environment.NewLine);

        internal XlsxSharedStringsXml(Package p, PackagePart p1)
        {
            fs = new FileStream(Path.GetTempFileName(), FileMode.Create);
            Uri u9 = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
            sharedStringsPt = p.CreatePart(u9, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", CompressionOption.Normal);

            StringBuilder sb = new StringBuilder();
            sb.Append(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            sb.Append(Environment.NewLine);
            sb.Append(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" >");
            sb.Append(Environment.NewLine);

            byte[] b = Encoding.UTF8.GetBytes(sb.ToString());
            fs.Write(b, 0, b.Length);

            p1.CreateRelationship(new Uri("sharedStrings.xml", UriKind.Relative), TargetMode.Internal
                , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
        }

        /// <summary>
        /// Write a string value to the sharedStrings.xml file, and write the sharedStrings array entry in the current worksheet stream.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal long write(string s)
        {
            Crc32 cc = new Crc32();

            byte[] ba = Encoding.UTF8.GetBytes(s);
            string res = string.Empty;
            foreach (byte w in cc.ComputeHash(ba))
                res += w.ToString("x2").ToLower();
            //byte[] b = cc.ComputeHash(ba);

            long f=0;
            a.TryGetValue(res, out f);

            //bool cn = a.Keys.Where(x => x[0] == b[0] && x[1] == b[1] && x[2] == b[2] && x[3] == b[3]).Any();

            if (a.Count == 0 || (f == 0 && !a.ContainsKey(res)))
            {
                //byte[] ba = Encoding.UTF8.GetBytes(s);
                // fixme, xml:space="preserve"
                fs.Write(openString, 0, openString.Length);
                fs.Write(ba, 0, ba.Length);
                fs.Write(closeString, 0, closeString.Length);
                a.Add(res, ss); 
                return ss++;
            }
            else
                return f;
        }

        internal void close()
        {
            byte[] b = Encoding.UTF8.GetBytes("</sst>");
            fs.Write(b, 0, b.Length);

            fs.Flush();
            fs.Seek(0, SeekOrigin.Begin);
            StaticFunctions.copyStream(fs, sharedStringsPt.GetStream());
            fs.Close();
            File.Delete(fs.Name);
        }
    }
}
