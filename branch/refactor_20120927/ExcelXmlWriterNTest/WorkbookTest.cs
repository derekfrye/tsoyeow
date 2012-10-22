using ExcelXmlWriter;
using System.IO;
using System;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Data.SqlClient;
using Ionic.Zip;
using NUnit.Framework;
using System.Collections.Generic;

namespace ExcelXmlWriterTest
{
    /// <summary>
    ///This is a test class for WorkbookTest and is intended
    ///to contain all WorkbookTest Unit Tests
    ///</summary>
    [TestFixture()]
    public class WorkbookTest
    {
        /// <summary>
        ///A test for Workbook Constructor
        ///</summary>
        [Test()]
        public void BrokenWorkbookFromFileConstructorTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "Data.xml";

            p.Query = path;
            p.FromFile = true;
            //p.connStr = connStr;
            //p.columnTypeMappings = columnTypeMappings;
            p.MaxRowsPerSheet = 100000;
            //p.resultNames = resultNames;
            
            //p.defaultColumnType = ExcelDataType.General;
            p.QueryTimeout = 0;
            //p.numberFormatCulture = c1;

            Workbook target = new Workbook(p);
            MemoryStream fs = new MemoryStream();

            if (target.RunQuery())
                target.WriteQueryResults(fs);

            fs.Flush();
            fs.Seek(0, SeekOrigin.Begin);

            StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

            fs.Seek(0, SeekOrigin.Begin);
            string[] a = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int len = a.Length;

            fs.Close();
            sr.Close();

            Assert.AreEqual(len, 146);

            XElement x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                .Where(x1 => x1
                    .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet1_1")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

            XElement x3 = x2
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").Last();

            Assert.AreEqual(x3.Value, "2009-12-18T18:19:25");
            Assert.AreEqual(x3.Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Type").Value, "DateTime");
        }

        [Test()]
        public void BrokenWorkbookQueryConstructorTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL.sql";

            StreamReader sr = new StreamReader(path);
            p.Query = sr.ReadToEnd();
            sr.Close();
            p.FromFile = false;

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = @".";
                sb.InitialCatalog = "master";
            sb.IntegratedSecurity = true;
			p.ConnectionString = sb.ConnectionString;
            //p.columnTypeMappings = columnTypeMappings;
            p.MaxRowsPerSheet = 100000;
            //p.resultNames = resultNames;
           //p.defaultColumnType = ExcelDataType.General;
            p.QueryTimeout = 0;
            p.WriteEmptyResultSetColumns = false;
            //p.numberFormatCulture = c1;

            Workbook target = new Workbook(p);
            MemoryStream fs = new MemoryStream();

            if (target.RunQuery())
                target.WriteQueryResults(fs);

            fs.Flush();
            fs.Seek(0, SeekOrigin.Begin);

            sr = new StreamReader(fs, Encoding.UTF8);

            XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

            fs.Seek(0, SeekOrigin.Begin);
            string[] a = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            int len = a.Length;

            fs.Close();
            sr.Close();

            Assert.AreEqual(len, 142);

            XElement x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                .Where(x1 => x1
                    .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet1_1")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

            XElement x3 = x2
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").Last();

            Assert.AreEqual(x3.Value, "2009-12-08T18:19:17");
            Assert.AreEqual(x3.Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Type").Value, "DateTime");

            p.WriteEmptyResultSetColumns = true;
            target = new Workbook(p);
            fs = new MemoryStream();

            if (target.RunQuery())
                target.WriteQueryResults(fs);

            fs.Flush();
            fs.Seek(0, SeekOrigin.Begin);

            sr = new StreamReader(fs, Encoding.UTF8);

            x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

            fs.Seek(0, SeekOrigin.Begin);
            string[] a1 = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            len = a1.Length;

            fs.Close();
            sr.Close();

            Assert.AreEqual(len, 181);

            // last row
            x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                .Where(x1 => x1
                    .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet3_1")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

            // last cell
            x3 = x2
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").Last();

            Assert.AreEqual(x3.Value, "this is also on hte new sheet");

            // only 1 row
            Assert.AreEqual(x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                .Where(x1 => x1
                    .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet2_1")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Count(), 1);
        }

        [Test()]
        public void BrokenWorkbookQueryWriteResultsOverSizeTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit.sql";

            StreamReader sr = new StreamReader(path);
            p.Query = sr.ReadToEnd();
            sr.Close();
            p.FromFile = false;

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = @".";
            sb.InitialCatalog = "master";
            sb.IntegratedSecurity = true;
            p.ConnectionString = sb.ConnectionString;
            //p.columnTypeMappings = columnTypeMappings;
            p.MaxRowsPerSheet = 100000;
            //p.resultNames = resultNames;
            //p.de = ExcelDataType.General;
            p.QueryTimeout = 0;
            // 1 MB
            p.MaxWorkBookSize = 1000000;
            p.WriteEmptyResultSetColumns = false;
            //p.numberFormatCulture = c1;

            Workbook target = new Workbook(p);
            IList<MemoryStream> fs1 = new List<MemoryStream>();

            if (target.RunQuery())
            {
                MemoryStream fs = new MemoryStream();
                fs1.Add(fs);
                WorkBookStatus status = target.WriteQueryResults(fs);
                while (status != WorkBookStatus.Completed)
                {
                    MemoryStream fsa = new MemoryStream();
                    fs1.Add(fsa);
                    status = target.WriteQueryResults(fsa);
                }
            }

            int currentStream = 1;
            foreach (MemoryStream fs in fs1)
            {
                fs.Flush();
                fs.Seek(0, SeekOrigin.Begin);

                sr = new StreamReader(fs, Encoding.UTF8);
                XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

                fs.Seek(0, SeekOrigin.Begin);
                string[] a = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                

                fs.Close();
                sr.Close();

                XElement x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet1_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

                // first element
                XElement x3 = x2
                    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").First();

                int len = a.Length;
                //int aaa = currentStream;
                //int dfjkla = Convert.ToInt32(x3.Value);

                if (currentStream == 1)
                {
                    Assert.AreEqual(len, 14968);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 1063);

                }
                else if (currentStream == 5)
                {
                    Assert.AreEqual(len, 14940);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 5307);

                }
                else if (currentStream == 22)
                {
                    Assert.AreEqual(len, 14926);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 23331);

                }
                else if (currentStream == 31)
                {
                    Assert.AreEqual(len, 13484);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 32768);

                }
                else if (currentStream > 31)
                    Assert.Fail("Too many streams initiated");

                currentStream++;
            }
        }

        [Test()]
        public void BrokenWorkbookQueryWriteResultOverSizeTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit 5 result sets.sql";

            StreamReader sr = new StreamReader(path);
            p.Query = sr.ReadToEnd();
            sr.Close();
            p.FromFile = false;

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = @".";
            sb.InitialCatalog = "master";
            sb.IntegratedSecurity = true;
            p.ConnectionString = sb.ConnectionString;
            //p.columnTypeMappings = columnTypeMappings;
            p.MaxRowsPerSheet = 100000;
            //p.resultNames = resultNames;
            //p.DefaultColumnType = ExcelDataType.General;
            p.QueryTimeout = 0;
            // 100 KB
            p.MaxWorkBookSize = 100000;
            p.WriteEmptyResultSetColumns = false;
            //p.numberFormatCulture = c1;

            Workbook target = new Workbook(p);
            IList<MemoryStream> fs1 = new List<MemoryStream>();

            if (target.RunQuery())
            {
                int currentFile = 1;
                MemoryStream fs = new MemoryStream();
                fs1.Add(fs);
                while (target.NextResult())
                {
                    if (currentFile != 1)
                    {
                        fs = new MemoryStream();
                        fs1.Add(fs);
                    }
                    WorkBookStatus status = target.WriteQueryResult(fs);
                    while (status != WorkBookStatus.Completed)
                    {
                        //fs.Close();
                        currentFile++;
                        fs = new MemoryStream();
                        fs1.Add(fs);
                        status = target.WriteQueryResult(fs);
                    }
                    //fs.Close();
                    currentFile++;
                }
                target.QueryClose();
            }

            int currentStream = 1;
            bool allStreams = false;
            foreach (MemoryStream fs in fs1)
            {
                fs.Flush();
                fs.Seek(0, SeekOrigin.Begin);

                sr = new StreamReader(fs, Encoding.UTF8);
                XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

                fs.Seek(0, SeekOrigin.Begin);
                string[] a = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);


                fs.Close();
                sr.Close();

                string vl = x.Element("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Element("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet").Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value;

                XElement x2 = null;
                    
                    if(currentStream<=3)
                x2= x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet1_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();
                    else if (currentStream<= 6)
                        x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet2_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();
                    else if (currentStream == 7)
                        x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet3_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();
                    else if (currentStream == 8)
                        x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet4_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();
                    else if (currentStream > 8)
                        x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                        .Where(x1 => x1
                            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet5_1")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

                // first element
                XElement x3 = x2
                    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").First();

                int len = a.Length;

                if (currentStream == 1 || currentStream == 4 || currentStream == 9)
                {
                    Assert.AreEqual(len, 1556);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 105);
                }
                else if (currentStream == 2 || currentStream == 5 || currentStream == 10)
                {
                    Assert.AreEqual(len, 1556);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 210);
                }
                else if (currentStream == 3 || currentStream == 6 || currentStream == 11)
                {
                    Assert.AreEqual(len, 730);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 256);
                }
                else if (currentStream == 7)
                {
                    Assert.AreEqual(len, 1346);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 90);
                }
                else if (currentStream == 8)
                {
                    Assert.AreEqual(len, 1528);
                    Assert.AreEqual(Convert.ToInt32(x3.Value), 103);
                }
                else
                    Assert.Fail("Too many streams initiated");

                if (currentStream == 11)
                    allStreams = true;

                currentStream++;
            }

            if (!allStreams)
                Assert.Fail("Not enough streams tested");

        }
        
        [Test()]
        public void IncompleteXlsxFromFileSeparateTabsTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "Data.xml";

            p.Query = path;
            p.FromFile = true;
            //p.connStr = connStr;
            //p.columnTypeMappings = columnTypeMappings;
            p.MaxRowsPerSheet = 100000;
            p.ResultNames = new Dictionary<int, string>();
            p.ResultNames.Add(1, "blah blah");
            p.ResultNames.Add(2, "x");
            
            p.QueryTimeout = 0;
            //p.numberFormatCulture = c1;

            Workbook target = new Workbook(p);

            string path1 = Path.GetTempFileName();

            if (target.RunQuery())
                target.WriteQueryResults(path1);


            string path2 = Environment.CurrentDirectory;
            path2 = Path.GetDirectoryName(path2);
            path2 = Path.GetDirectoryName(path2);
            path2 = Path.GetDirectoryName(path2);
            path2 = path2 + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterTest_vs2008" + Path.DirectorySeparatorChar.ToString()
                + "test.xlsx";

            File.Copy(path1, path2, true);

            //Package pa = Package.Open(path2);
            ZipFile z = new ZipFile(path2);
            var t = z.SelectEntries("xl/worksheets/sheet1_1.xml").First();
            MemoryStream ms = new MemoryStream();
            t.Extract(ms);
            ms.Flush();
            ms.Seek(0, SeekOrigin.Begin);

            //PackagePart strm = pa.GetPart(new Uri("/xl/worksheets/sheet1_1.xml", UriKind.Relative));

            //Stream m = t.InputStream;

            StreamReader sr = new StreamReader(ms, Encoding.UTF8);
            string b = sr.ReadToEnd();
            XDocument x = XDocument.Parse(b, LoadOptions.None);

            ms.Close();

            var asdz =
                // Xml is e.g. <worksheet><sheetData><row>...</row><row><c s="1">12345.2423</c>...
                x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements()
                .First(aaa => aaa.Name.LocalName == "row"
                && aaa.Elements().Where(bbb => bbb.Name.LocalName == "c").Any(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1))
                .Elements().First(ccc => ccc.Name.LocalName == "c" && ccc.Attributes("s").Any() && Convert.ToInt32(ccc.Attribute("s").Value) == 1).Value;

            // ensure correct xl date value
            Assert.AreEqual(Convert.ToDouble(asdz)
            , 40155.7633988426);

            var asdz2 = x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements().Last(aaa => aaa.Name.LocalName == "row")
                .Elements().Where(aaa=>aaa.Name.LocalName=="c").Where(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1).First();

            // ensure correct xl date value
            Assert.AreEqual(Convert.ToDouble(asdz2.Elements().First(aaa=>aaa.Name.LocalName=="v").Value)
            , 40165.7634929051);

            // ensure correct cell counts in 1st row
            Assert.AreEqual(x.Element("worksheet").Element("sheetData").Elements("row").First()
                .Elements("c").Count(), 3);

            // ensure correct cell counts in 2nd (and last row)
            Assert.AreEqual(x.Element("worksheet").Element("sheetData").Elements("row").Last()
                .Elements("c").Count(), 3);

            // ensure correct shared cell refernce in last row
            Assert.AreEqual(Convert.ToInt32(x.Element("worksheet").Element("sheetData").Elements("row").Last()
                .Elements("c").Where(xx => xx.Attributes("t").Any() && xx.Attribute("t").Value == "s").First().Element("v").Value)
                , 1);

            
        }
    }
}
