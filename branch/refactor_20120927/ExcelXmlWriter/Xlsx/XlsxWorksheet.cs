using System;
using System.Collections.Generic;
using System.Text;
using ExcelXmlWriter;
using System.IO;
using System.IO.Packaging;
using System.Data;
using System.Collections;

namespace ExcelXmlWriter
{
    class XlsxWorksheet
    {

        public string sheetname;
        FileStream fs;
        string id;

        internal string Id
        {
            get { return id; }
        }

        internal FileStream s
        { get { return fs; } }

        internal XlsxWorksheet(int count, int subcount, string name, string id, DataRowCollection d, XlsxSharedStringsXml s)
        {
            fs = new FileStream(Path.GetTempFileName(), FileMode.Create);

            this.id = id;
            sheetname = name;

            write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
            write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
            + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" 
            + " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" 
            + " mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">" 
            + Environment.NewLine);
            write(@"<dimension ref=""A1""/>" + Environment.NewLine);
            write(@"<sheetViews>" + Environment.NewLine);
            write(@"<sheetView workbookViewId=""0"">" + Environment.NewLine);
            write(@"<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen"" />");
            write(@"<selection pane=""bottomLeft"" activeCell=""A2"" sqref=""A2"" />");
            write(@"</sheetView>");
            write(@"</sheetViews>" + Environment.NewLine);
            write(@"<sheetFormatPr defaultRowHeight=""15"" x14ac:dyDescent=""0.25""/>" + Environment.NewLine);
            write(@"<sheetData>" + Environment.NewLine);

            // write row hdr
            write(XlsxRow.hdr);
            foreach (DataRow rows in d)
            {
                // FIXME don't hardcode 100
                // FIXME the call to overpunch happens twice, could just happen once with appropriate reutnr value
                writeval(rows["ColumnName"].ToString(), StaticFunctions.ResolveDataType(rows["ColumnName"].ToString(), 100), s);
            }
            // write row close
            write(XlsxRow.hdrclose);
        }

        void write(string s)
        {
            byte[] b = Encoding.UTF8.GetBytes(s);
            fs.Write(b, 0, b.Length);
            //fileSize += b.Length;
        }

        void writeval(string p, ExcelDataType excelDataType, XlsxSharedStringsXml s)
        {
            write(XlsxCell.hdr(excelDataType));
            switch (excelDataType)
            {
                case ExcelDataType.Number:
                case ExcelDataType.Date:
                    write(XlsxData.DataVal(p, excelDataType));
                    break;
                case ExcelDataType.OverpunchNumber:
                    //fixme don't hardcoe 100
                    Overpunch i = StaticFunctions.applyOverPunch(p, 100);
                    write(XlsxData.DataVal(i.val.ToString(), ExcelDataType.Number));
                    break;
                default:
                    long d=s.write(XlsxData.DataVal(p, excelDataType));
                    write(d.ToString());
                    break;
            }
            write(XlsxCell.hdrclose);
        }

        internal void writerow(IDataReader queryReader, XlsxSharedStringsXml s)
        {
            // write row hdr
            write(XlsxRow.hdr);
            // cycle through the columns, writing the values
            for (int i = 0; i < queryReader.FieldCount; i++)
            {
                // FIXME don't hardcode 100
                // FIXME the call to overpunch happens twice, could just happen once with appropriate reutnr value
                writeval(queryReader[i].ToString(), StaticFunctions.ResolveDataType(queryReader[i].ToString(), 100), s);
            }
            // write row close
            write(XlsxRow.hdrclose);
        }

        internal void Close()
        {
            write(@"</sheetData>");
            write(@"<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>");
            write(@"</worksheet>");
            fs.Flush();
            fs.Seek(0, SeekOrigin.Begin);
        }

        //void Dispose()
        //{
        //    fs.Close();
        //    File.Delete(fs.Name);
        //}

        //void hdr()
        //{
        //    write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>");
        //    write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        //    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        //    + " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
        //    + " mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">"
        //    + Environment.NewLine);
        //    write(@"<dimension ref=""A1""/>" + Environment.NewLine);
        //    write(@"<sheetViews>" + Environment.NewLine);
        //    write(@"<sheetView workbookViewId=""0""/>" + Environment.NewLine);
        //    write(@"</sheetViews>" + Environment.NewLine);
        //    write(@"<sheetFormatPr defaultRowHeight=""15"" x14ac:dyDescent=""0.25""/>" + Environment.NewLine);
        //    write(@"<sheetData>" + Environment.NewLine);
        //}
    }
}
