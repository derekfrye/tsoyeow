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
using ExcelXmlQueryResults;
using ExcelXmlWriter.Workbook;
using ExcelXmlWriterNTest.Properties;

namespace ExcelXmlWriterTest
{
    /// <summary>
    ///This is a test class for WorkbookTest and is intended
    ///to contain all WorkbookTest Unit Tests
    ///</summary>
    //[TestClass()]
    [TestFixture()]
    public class WorkbookTest
    {

        [Test()]
        [TestCase(1,1)]
        public void WorkbookProcessBatchesTest()
        {

            WorkBookParams p = new WorkBookParams();

            string[] tt = new string[2];

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            string tempOutputPath = Path.Combine(path, "ExcelXmlWriterNTest", "bin", "Debug");
            //tempOutputPath = Path.Combine(path, Utility.getIncrFileName(1, System.Reflection.MethodInfo.GetCurrentMethod().Name));

            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit 5 result sets.sql";

            tt[0] = path;

            path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            tempOutputPath = Path.Combine(path, "ExcelXmlWriterNTest", "bin", "Debug");
            //tempOutputPath = Path.Combine(path, Utility.getIncrFileName(1, System.Reflection.MethodInfo.GetCurrentMethod().Name));

            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit 5 result sets.sql";

            tt[1] = path;

            Dictionary<string, FileStream> fs1 = new Dictionary<string, FileStream>();
            int i = 1;

            // re-create rather than make the original static, or have to instantiate a FormEntrance
            foreach (var qry in tt)
            {
                using (StreamReader sr = new StreamReader(new FileStream(qry, FileMode.Open, FileAccess.Read)))
                {
                    p.Query = sr.ReadToEnd();
                    p.FromFile = false;
                    SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
                    sb.DataSource = Settings.Default.datasource;
                    sb.InitialCatalog = Settings.Default.database;
                    sb.IntegratedSecurity = true;
                    p.ConnectionString = sb.ConnectionString;
                    p.MaxRowsPerSheet = 150000;
                    //p.MaxWorkBookSize = 1000000;
                    p.DupeKeysToDelayStartingNewWorksheet = new string[] { "a1" };
                    p.MaximumResultSetsPerWorkbook = 2;

                    Workbook target = new Workbook(p);
                    
                    if (target.RunQuery())
                    {
                        i = i + 1;
                        string pht = Path.Combine(tempOutputPath
                            , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                        FileStream fs = new FileStream(pht
                            , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                        fs1.Add(pht, fs);
                        WorkBookStatus status = target.WriteQueryResults(fs);
                        while (status != WorkBookStatus.Completed)
                        {
                            i = i + 1;
                            pht = Path.Combine(tempOutputPath
                                , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                            FileStream fsa = new FileStream(pht
                            , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                            fs1.Add(pht, fsa);
                            status = target.WriteQueryResults(fsa);
                            fsa.Close();
                        }
                        fs.Close();
                    }

                }
            }            

            Assert.AreEqual(fs1.Count, 6);

            int currentStream = 1;
            int worksheet = 1;
            int resetVariable = 1;
            foreach (var fs in fs1)
            {
                // the first 3 entries in fs1 are from the same query, the next 3 are from the next queyr
                if (resetVariable == 4)
                {
                    currentStream = 1;
                    worksheet = 1;
                }
                resetVariable++;

                fs.Value.Close();
                //fs.Value.Flush();
                //fs.Value.Seek(0, SeekOrigin.Begin);
                for (int ii = 1; ii < 3; ii++)
                {
                    if (worksheet < 6)
                    {
                        ZipFile z = new ZipFile(fs.Key);
                        // broken here
                        var t = z.SelectEntries("xl/worksheets/sheet" + worksheet.ToString() + "_1.xml").First();

                        MemoryStream ms = new MemoryStream();
                        t.Extract(ms);
                        ms.Flush();
                        ms.Seek(0, SeekOrigin.Begin);

                        //PackagePart strm = pa.GetPart(new Uri("/xl/worksheets/sheet1_1.xml", UriKind.Relative));

                        //Stream m = t.InputStream;

                        StreamReader sr1 = new StreamReader(ms, Encoding.UTF8);
                        string b1 = sr1.ReadToEnd();
                        XDocument x = XDocument.Parse(b1, LoadOptions.None);

                        ms.Close();

                        var asdza =
                            // Xml is e.g. <worksheet><sheetData><row>...</row><row><c s="1">12345.2423</c>...
                            x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                            .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements()
                            .Where(aaa => aaa.Name.LocalName == "row").Count();
                        if (currentStream == 1)
                        {
                            if (worksheet == 1 || worksheet == 2)
                            {
                                Assert.AreEqual(asdza, 257);
                            }

                        }
                        else if (currentStream == 2)
                        {
                            if (worksheet == 3)
                            {
                                Assert.AreEqual(asdza, 91);
                            }
                            else if (worksheet == 4)
                            {
                                Assert.AreEqual(asdza, 104);
                            }
                        }

                        else if (currentStream == 3)
                        {
                            Assert.AreEqual(asdza, 257);
                        }

                        worksheet++;
                    }
                }

                currentStream++;
            }
        }
        

        [Test()]
        public void WorkbookQueryWriteResultsOverSizeTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            string tempOutputPath=Path.Combine(path,"ExcelXmlWriterNTest","bin","Debug");
            //tempOutputPath = Path.Combine(path, Utility.getIncrFileName(1, System.Reflection.MethodInfo.GetCurrentMethod().Name));

            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit.sql";

            StreamReader sr = new StreamReader(path);
            p.Query = sr.ReadToEnd();
            sr.Close();
            p.FromFile = false;

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = Settings.Default.datasource;
            sb.InitialCatalog = Settings.Default.database;
            sb.IntegratedSecurity = true;
            p.ConnectionString = sb.ConnectionString;
            //p.columnTypeMappings = columnTypeMappings;

            //p.resultNames = resultNames;
            //p.de = ExcelDataType.General;
            p.QueryTimeout = 0;

            p.WriteEmptyResultSetColumns = false;
            //p.numberFormatCulture = c1;            

            // breaks into 6 workbooks
            // workbook 1 & 2: 4097 rows, all in 1 sheet
            // workbook 3: 8193 rows, all in 1 sheet
            // workbook 4: 16385 rows, all in 1 sheet
            // workbook 5: 32769 rows, all in 1 sheet
            // workbook 6: 65537 rows, all in 1 sheet
            p.MaxRowsPerSheet = 38641;
            p.MaxWorkBookSize = 1000000;
            p.DupeKeysToDelayStartingNewWorksheet = new string[] { "a1" };

            Workbook target = new Workbook(p);
            Dictionary<string, FileStream> fs1 = new Dictionary<string, FileStream>();

            //string tempOutputPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            int i = 1;
            if (target.RunQuery())
            {
                string pht=Path.Combine(tempOutputPath
                    , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                FileStream fs = new FileStream(pht
                    , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                fs1.Add(pht,fs);
                WorkBookStatus status = target.WriteQueryResults(fs);
                while (status != WorkBookStatus.Completed)
                {
                    i = i + 1;
                    pht=Path.Combine(tempOutputPath
                        , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                    FileStream fsa = new FileStream(pht
                    , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                    fs1.Add(pht,fsa);
                    status = target.WriteQueryResults(fsa);
                }
            }

            int currentStream = 1;
            foreach (var fs in fs1)
            {
                fs.Value.Close();
                //fs.Value.Flush();
                //fs.Value.Seek(0, SeekOrigin.Begin);

                ZipFile z = new ZipFile(fs.Key);
                // broken here
                var t = z.SelectEntries("xl/worksheets/sheet1_1.xml").First();
                MemoryStream ms = new MemoryStream();
                t.Extract(ms);
                ms.Flush();
                ms.Seek(0, SeekOrigin.Begin);

                //PackagePart strm = pa.GetPart(new Uri("/xl/worksheets/sheet1_1.xml", UriKind.Relative));

                //Stream m = t.InputStream;

                StreamReader sr1 = new StreamReader(ms, Encoding.UTF8);
                string b1 = sr1.ReadToEnd();
                XDocument x = XDocument.Parse(b1, LoadOptions.None);

                ms.Close();

                var asdza =
                    // Xml is e.g. <worksheet><sheetData><row>...</row><row><c s="1">12345.2423</c>...
                    x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                    .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements()
                    .Where(aaa => aaa.Name.LocalName == "row").Count();
                if (currentStream == 1)
                {                    
                    Assert.AreEqual(asdza, 4097);
                }
                else if (currentStream == 2)
                {
                    Assert.AreEqual(asdza, 4097);
                }
                else if (currentStream == 3)
                {
                    Assert.AreEqual(asdza, 8193);
                }
                else if (currentStream == 4)
                {
                    Assert.AreEqual(asdza, 16385);
                }
                else if (currentStream == 5)
                {
                    Assert.AreEqual(asdza, 32769);
                }
                else if (currentStream == 6)
                {
                    Assert.AreEqual(asdza, 65537);
                }

                var asdz =
                    // Xml is e.g. <worksheet><sheetData><row>...</row><row><c s="1">12345.2423</c>...
                    x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                    .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements()
                    .First(aaa => aaa.Name.LocalName == "row"
                    && aaa.Elements().Where(bbb => bbb.Name.LocalName == "c").Any(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1))
                    .Elements().First(ccc => ccc.Name.LocalName == "c" && ccc.Attributes("s").Any() && Convert.ToInt32(ccc.Attribute("s").Value) == 1).Value;

                // ensure correct xl date value
                //Assert.AreEqual(Convert.ToDouble(asdz)
                //, 40155.7633988426);

                var asdz2 = x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                    .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements().Last(aaa => aaa.Name.LocalName == "row")
                    .Elements().Where(aaa => aaa.Name.LocalName == "c").Where(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1).First();

                //sr = new StreamReader(fs, Encoding.UTF8);
                //XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.PreserveWhitespace);

                //fs.Seek(0, SeekOrigin.Begin);
                //string[] a = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                

                //fs.Close();
                //sr.Close();

                //XElement x2 = x.Elements("{urn:schemas-microsoft-com:office:spreadsheet}Workbook")
                //        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Worksheet")
                //        .Where(x1 => x1
                //            .Attribute("{urn:schemas-microsoft-com:office:spreadsheet}Name").Value == "Sheet1_1")
                //        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Table")
                //        .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Row").Last();

                //// first element
                //XElement x3 = x2
                //    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Cell")
                //    .Elements("{urn:schemas-microsoft-com:office:spreadsheet}Data").First();

                //int len = a.Length;
                ////int aaa = currentStream;
                ////int dfjkla = Convert.ToInt32(x3.Value);

                //if (currentStream == 1)
                //{
                //    Assert.AreEqual(len, 14968);
                //    Assert.AreEqual(Convert.ToInt32(x3.Value), 1063);

                //}
                //else if (currentStream == 5)
                //{
                //    Assert.AreEqual(len, 14940);
                //    Assert.AreEqual(Convert.ToInt32(x3.Value), 5307);

                //}
                //else if (currentStream == 22)
                //{
                //    Assert.AreEqual(len, 14926);
                //    Assert.AreEqual(Convert.ToInt32(x3.Value), 23331);

                //}
                //else if (currentStream == 31)
                //{
                //    Assert.AreEqual(len, 13484);
                //    Assert.AreEqual(Convert.ToInt32(x3.Value), 32768);

                //}
                //else if (currentStream > 31)
                //    Assert.Fail("Too many streams initiated");

                currentStream++;
            }
        }

        [Test()]
        public void WorkbookQueryMaxResultsPerWorkbookTest()
        {
            WorkBookParams p = new WorkBookParams();

            string path = Environment.CurrentDirectory;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            string tempOutputPath = Path.Combine(path, "ExcelXmlWriterNTest", "bin", "Debug");
            //tempOutputPath = Path.Combine(path, Utility.getIncrFileName(1, System.Reflection.MethodInfo.GetCurrentMethod().Name));

            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "SQL exceeds filesize limit 5 result sets.sql";

            StreamReader sr = new StreamReader(path);
            p.Query = sr.ReadToEnd();
            sr.Close();
            p.FromFile = false;

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();
            sb.DataSource = Settings.Default.datasource;
            sb.InitialCatalog = Settings.Default.database;
            sb.IntegratedSecurity = true;
            p.ConnectionString = sb.ConnectionString;
            //p.columnTypeMappings = columnTypeMappings;

            //p.resultNames = resultNames;
            //p.de = ExcelDataType.General;
            p.QueryTimeout = 0;

            p.WriteEmptyResultSetColumns = false;
            //p.numberFormatCulture = c1;            

            // should break into 3 workbooks, 961 rows total, 2 result sets per workbook
            // 256, 256, 90, 103, 256
            
            p.MaxRowsPerSheet = 150000;
            //p.MaxWorkBookSize = 1000000;
            p.DupeKeysToDelayStartingNewWorksheet = new string[] { "a1" };
            p.MaximumResultSetsPerWorkbook = 2;

            Workbook target = new Workbook(p);
            Dictionary<string, FileStream> fs1 = new Dictionary<string, FileStream>();

            //string tempOutputPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            int i = 1;
            if (target.RunQuery())
            {
                string pht = Path.Combine(tempOutputPath
                    , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                FileStream fs = new FileStream(pht
                    , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                fs1.Add(pht, fs);
                WorkBookStatus status = target.WriteQueryResults(fs);
                while (status != WorkBookStatus.Completed)
                {
                    i = i + 1;
                    pht = Path.Combine(tempOutputPath
                        , Utility.getIncrFileName(i, System.Reflection.MethodInfo.GetCurrentMethod().Name) + ".xlsx");
                    FileStream fsa = new FileStream(pht
                    , FileMode.Create, FileAccess.ReadWrite, FileShare.None);
                    fs1.Add(pht, fsa);
                    status = target.WriteQueryResults(fsa);
                    fsa.Close();
                }
                fs.Close();
            }

            Assert.AreEqual(fs1.Count, 3);

            int currentStream = 1;
            int worksheet = 1;
            foreach (var fs in fs1)
            {
                fs.Value.Close();
                //fs.Value.Flush();
                //fs.Value.Seek(0, SeekOrigin.Begin);
                for (int ii = 1; ii < 3; ii++)
                {
                    if (worksheet < 6)
                    {
                        ZipFile z = new ZipFile(fs.Key);
                        // broken here
                        var t = z.SelectEntries("xl/worksheets/sheet" + worksheet.ToString() + "_1.xml").First();

                        MemoryStream ms = new MemoryStream();
                        t.Extract(ms);
                        ms.Flush();
                        ms.Seek(0, SeekOrigin.Begin);

                        //PackagePart strm = pa.GetPart(new Uri("/xl/worksheets/sheet1_1.xml", UriKind.Relative));

                        //Stream m = t.InputStream;

                        StreamReader sr1 = new StreamReader(ms, Encoding.UTF8);
                        string b1 = sr1.ReadToEnd();
                        XDocument x = XDocument.Parse(b1, LoadOptions.None);

                        ms.Close();

                        var asdza =
                            // Xml is e.g. <worksheet><sheetData><row>...</row><row><c s="1">12345.2423</c>...
                            x.Elements().First(aa => aa.Name.LocalName == "worksheet")
                            .Elements().First(aa => aa.Name.LocalName == "sheetData").Elements()
                            .Where(aaa => aaa.Name.LocalName == "row").Count();
                        if (currentStream == 1)
                        {
                            if (worksheet == 1 || worksheet == 2)
                            {
                                Assert.AreEqual(asdza, 257);
                            }

                        }
                        else if (currentStream == 2)
                        {
                            if (worksheet == 3)
                            {
                                Assert.AreEqual(asdza, 91);
                            }
                            else if (worksheet == 4)
                            {
                                Assert.AreEqual(asdza, 104);
                            }
                        }

                        else if (currentStream == 3)
                        {
                            Assert.AreEqual(asdza, 257);
                        }

                        worksheet++;
                    }
                }

                currentStream++;
            }
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


            //string path2 = Environment.CurrentDirectory;
            //path2 = Path.GetDirectoryName(path2);
            //path2 = Path.GetDirectoryName(path2);
            //path2 = Path.GetDirectoryName(path2);
            //path2 = path2 + Path.DirectorySeparatorChar.ToString()
            //    + "ExcelXmlWriterTest_vs2008" + Path.DirectorySeparatorChar.ToString()
            //    + "test.xlsx";

            //File.Copy(path1, path2, true);

            //Package pa = Package.Open(path2);
            ZipFile z = new ZipFile(path1);
            // broken here, looks like due to reading from file is broken in queryreader
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
            

            // ensure correct cell counts in 2nd (and last row)
            

            // ensure correct shared cell refernce in last row
            

            
        }
    }
}
