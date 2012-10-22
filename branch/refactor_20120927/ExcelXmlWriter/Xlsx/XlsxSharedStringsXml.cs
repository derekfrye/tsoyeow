using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace ExcelXmlWriter.Xlsx
{
    /// <summary>
    /// /xl/sharedStrings.xml
    /// </summary>
    class XlsxSharedStringsXml : IDisposable
    {
        long curentSharedStringPosition = 0;
        XmlWriter xw;
        
        Dictionary<int, Tuple<long, string>> sharedStringDictionary = new Dictionary<int, Tuple<long, string>>();

        internal Stream OutputStream
        { get; private set; }

        internal string FileAssociateWithOutputStream
        { get; private set; }

        internal XlsxSharedStringsXml(Stream outputStream, string fileAssociatedWithOutputStream)
        {
            this.FileAssociateWithOutputStream = fileAssociatedWithOutputStream;

            OutputStream = outputStream;

            xw = XmlWriter.Create(OutputStream);
            xw.WriteStartDocument(true);

            xw.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            xw.WriteWhitespace(Environment.NewLine);
        }


        internal static ContentRelationships LinkToPackage(string reltp, string pth)
        {
            return new ContentRelationships() { PackagePath = pth, RelationshipType = reltp };
        }

        /// <summary>
        /// Write a string value to the sharedStrings.xml file, and write the sharedStrings array entry in the current worksheet stream.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        internal long Write(string value)
        {

            Tuple<long, string> f;

            var p = string.Join(null, value.ToCharArray().Where(x => XmlConvert.IsXmlChar(x)));
            var hdhdsdafd = p.GetHashCode();

            var found = sharedStringDictionary.TryGetValue(hdhdsdafd, out f);

            // if there was a match on hashcode, also determine if the string is identical
            // if so, return the pointer to the correct sharedstirng position
            if (found && string.Equals(f.Item2, value, StringComparison.InvariantCulture))
            {
                return f.Item1;
            }
            // if there isn't a hashcode match OR the string isnt identical, write it to sharedstrings

            else
            {
                xw.WriteStartElement("si");
                xw.WriteStartElement("t");

                xw.WriteString(p);

                xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteWhitespace(Environment.NewLine);

                // if we can't find the value (as opposed to a hashcode collision)
                // we need to write it to sharedstrings, and add it to the lookup array
                // FIXME make this count() test a parameter
                if (!found && sharedStringDictionary.Count < 500000)
                {
                    sharedStringDictionary.Add(hdhdsdafd, new Tuple<long, string>(curentSharedStringPosition, value));
                }
                curentSharedStringPosition++;

                // return the sharedstringposition we wrote
                return curentSharedStringPosition - 1;
            }
        }

        internal void Close()
        {
            xw.WriteEndElement();
            xw.Close();
            OutputStream.Flush();
            OutputStream.Seek(0, SeekOrigin.Begin);
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (xw != null)
                    xw.Close();
            }
        }

        #endregion
    }
}