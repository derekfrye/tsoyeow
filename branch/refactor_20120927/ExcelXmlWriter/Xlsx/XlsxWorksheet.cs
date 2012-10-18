using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelXmlWriter;
using System.IO;
using System.IO.Packaging;
using System.Data;
using System.Collections;
using System.Xml;

namespace ExcelXmlWriter
{
	class XlsxWorksheet
	{

		public string sheetname;
		XmlWriter wx;
		string id;
		
        public bool closed
        { get; private set; }

		internal string Id
		{
			get { return id; }
		}

        internal Stream sasdf
        {
            get;
            private set;
        }

        internal string filenmOs
        { get; private set; }

        /// <summary>
        /// Filename within the zip archive.
        /// </summary>
        internal string filenm
        { get; set; }


		internal XlsxWorksheet(Stream sss,int count, int subcount, string name, string id, DataRowCollection d, XlsxSharedStringsXml s,string jdfk)
		{
            filenmOs = jdfk;
			sasdf=sss;
            closed = false;
			wx = XmlWriter.Create(sasdf);

			this.id = id;
			sheetname = name;
			
			wx.WriteStartDocument(true);
			
			wx.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			
			wx.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			wx.WriteAttributeString("xmlns", "mc", null, "http://schemas.openxmlformats.org/markup-compatibility/2006");
			wx.WriteAttributeString("mc", "Ignorable", null, "x14ac");
			wx.WriteAttributeString("xmlns", "x14ac", null, "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("dimension");
			wx.WriteAttributeString("ref", "A1");
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("sheetViews");
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("sheetView");
			wx.WriteAttributeString("workbookViewId", "0");
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("pane");
			wx.WriteAttributeString("ySplit", "1");
			wx.WriteAttributeString("topLeftCell", "A2");
			wx.WriteAttributeString("activePane", "bottomLeft");
			wx.WriteAttributeString("state", "frozen");
			wx.WriteEndElement();
			
			wx.WriteStartElement("selection");
			wx.WriteAttributeString("pane", "bottomLeft");
			wx.WriteAttributeString("activeCell", "A2");
			wx.WriteAttributeString("sqref", "A2");
			wx.WriteEndElement();
			//</sheetView>
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
			//</sheetViews>
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("sheetFormatPr");
			wx.WriteAttributeString("defaultRowHeight", "15");
			wx.WriteAttributeString("x14ac", "dyDescent", null, "0.25");
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
			
			wx.WriteStartElement("sheetData");
			wx.WriteWhitespace(Environment.NewLine);

			// write row hdr
			
			wx.WriteStartElement("row");
			wx.WriteWhitespace(Environment.NewLine);
			foreach (DataRow rows in d)
			{
				// FIXME don't hardcode 100
				// FIXME the call to overpunch happens twice, could just happen once with appropriate reutnr value
				writeval(rows["ColumnName"].ToString(), StaticFunctions.ResolveDataType(rows["ColumnName"].ToString(), 100), s);
			}
			// write row close
			
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
		}

		void writeval(string p, ExcelDataType excelDataType, XlsxSharedStringsXml s)
		{
			//write(XlsxCell.hdr(excelDataType));
			wx.WriteStartElement("c");
			
			
			switch (excelDataType)
			{
				case ExcelDataType.Number:
				case ExcelDataType.Date:
					if(excelDataType==ExcelDataType.Date)
					{
						wx.WriteAttributeString("s","1");
					}
					wx.WriteStartElement("v");
					wx.WriteString(XlsxData.DataVal(p, excelDataType));
					break;
				case ExcelDataType.OverpunchNumber:
					wx.WriteStartElement("v");
					//fixme don't hardcoe 100
					Overpunch i = StaticFunctions.applyOverPunch(p, 100);
					wx.WriteString(XlsxData.DataVal(i.val.ToString(), ExcelDataType.Number));
					break;
				default:
					var pa=XlsxData.DataVal(p, excelDataType);
					long d=s.write(pa);
					
					wx.WriteAttributeString("t","s");
					wx.WriteStartElement("v");
					wx.WriteString(d.ToString());
					break;
			}
			// </v>
			wx.WriteEndElement();
			// </c>
            // once encountered an error saying stream was disposed..
			wx.WriteEndElement();
		}

		internal void writerow(IDataReader queryReader, XlsxSharedStringsXml s)
		{
			// write row hdr
			wx.WriteStartElement("row");
			wx.WriteWhitespace(Environment.NewLine);
			
			// cycle through the columns, writing the values
			for (int i = 0; i < queryReader.FieldCount; i++)
			{
				// FIXME don't hardcode 100
				// FIXME the call to overpunch happens twice, could just happen once with appropriate reutnr value
				writeval(queryReader[i].ToString(), StaticFunctions.ResolveDataType(queryReader[i].ToString(), 100), s);
			}
			
			// write row close
			wx.WriteEndElement();
			wx.WriteWhitespace(Environment.NewLine);
		}

		
		internal void Close()
		{
			//write(@"</sheetData>");
			wx.WriteEndElement();
			//write(@"<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>");
			wx.WriteStartElement("pageMargins");
			wx.WriteAttributeString("left", "0.7");
			wx.WriteAttributeString("right", "0.7");
			wx.WriteAttributeString("top", "0.75");
			wx.WriteAttributeString("bottom", "0.75");
			wx.WriteAttributeString("header", "0.3");
			wx.WriteAttributeString("footer", "0.3");
			wx.WriteEndElement();
			//write(@"</worksheet>");
			wx.WriteEndElement();
			// close the writer
			wx.Close();
			// close the stream
            sasdf.Flush();
            sasdf.Seek(0, SeekOrigin.Begin);
            closed = true;
		}

	}
}
