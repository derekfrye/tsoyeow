using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using System.IO;
using System.Collections;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace ExcelXmlWriter
{
	
	class XlsxSharedStringsXml
	{
		PackagePart sharedStringsPt;
		long curentSharedStringPosition = 0;
		//SHA256Cng h;
		XmlWriter xw;
		
		Dictionary<int, Tuple<long, string>> sharedStringDictionary = new Dictionary<int, Tuple<long, string>>();
		
		internal XlsxSharedStringsXml(Package p, PackagePart p1)
		{
			
			Uri u9 = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
			sharedStringsPt = p.CreatePart(u9, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
			                               , CompressionOption.Normal);
			Stream s=sharedStringsPt.GetStream();
			//h = new SHA256Cng();

			xw = XmlWriter.Create(s);
			xw.WriteStartDocument(true);
			
			xw.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xw.WriteWhitespace(Environment.NewLine);
			p1.CreateRelationship(new Uri("sharedStrings.xml", UriKind.Relative), TargetMode.Internal
			                      , "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
		}

		/// <summary>
		/// Write a string value to the sharedStrings.xml file, and write the sharedStrings array entry in the current worksheet stream.
		/// </summary>
		/// <param name="sa"></param>
		/// <returns></returns>
		internal long write(string sa)
		{

			Tuple<long, string> f;
			
			var p =string.Join(null,sa.ToCharArray().Where(x=> XmlConvert.IsXmlChar(x)));
			var hdhdsdafd=p.GetHashCode();
			
			var found=sharedStringDictionary.TryGetValue(hdhdsdafd, out f);

			// if there was a match on hashcode, also determine if the string is identical
			// if so, return the pointer to the correct sharedstirng position
			if (found && string.Equals(f.Item2, sa, StringComparison.InvariantCulture))
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
					sharedStringDictionary.Add(hdhdsdafd, new Tuple<long,string>(curentSharedStringPosition, sa));
				}
				curentSharedStringPosition++;
				
				// return the sharedstringposition we wrote
				return curentSharedStringPosition-1;
			}
		}

		internal void close()
		{
			xw.WriteEndElement();
			xw.Close();
		}
	}
}
