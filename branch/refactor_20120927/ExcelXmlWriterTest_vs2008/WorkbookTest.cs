using ExcelXmlWriter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
namespace ExcelXmlWriterTest_vs2008
{


    /// <summary>
    ///This is a test class for WorkbookTest and is intended
    ///to contain all WorkbookTest Unit Tests
    ///</summary>
    [TestClass()]
    public class WorkbookTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for Workbook Constructor
        ///</summary>
        [TestMethod()]
        public void XlsxFromFileSeparateTabsTest()
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

            p.query = path;
            p.fromFile = true;
            //p.connStr = connStr;
            //p.columnTypeMappings = columnTypeMappings;
            p.maxRowsPerSheet = 100000;
            p.resultNames = new Dictionary<int, string>();
            p.resultNames.Add(1, "blah blah");
            p.resultNames.Add(2, "x");
            
            p.queryTimeout = 0;
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

            Package pa = Package.Open(path2);

            PackagePart strm = pa.GetPart(new Uri("/xl/worksheets/sheet1_1.xml", UriKind.Relative));

            Stream m = strm.GetStream(FileMode.Open, FileAccess.Read);

            StreamReader sr = new StreamReader(m, Encoding.UTF8);
            XDocument x = XDocument.Parse(sr.ReadToEnd(), LoadOptions.None);

            m.Close();

            // ensure correct xl date value
            Assert.AreEqual(Convert.ToDouble(x.Element("worksheet").Element("sheetData").Elements("row").First()
                .Elements("c").Where(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1).First().Element("v").Value)
            , 40155.7633988426);

            // ensure correct xl date value
            Assert.AreEqual(Convert.ToDouble(x.Element("worksheet").Element("sheetData").Elements("row").Last()
                .Elements("c").Where(xx => xx.Attributes("s").Any() && Convert.ToInt32(xx.Attribute("s").Value) == 1).First().Element("v").Value)
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

            pa.Close();
        }
    }
}