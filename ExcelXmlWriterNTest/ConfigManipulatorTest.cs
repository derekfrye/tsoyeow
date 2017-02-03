using ExcelXmlQueryResults;
using NUnit.Framework;
using System.Xml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelXmlWriterTest
{
    
    /// <summary>
    ///This is a test class for ConfigManipulatorTest and is intended
    ///to contain all ConfigManipulatorTest Unit Tests
    ///</summary>
    [TestFixture()]
    public class ConfigManipulatorTest
    {

        string originalAppConfigPath = Assembly.GetExecutingAssembly().Location;
        string testAppConfigPath = Path.GetTempPath();

        #region Test Setup/Teardown

        [SetUp()]
        public void MyClassInitialize()
        {
            originalAppConfigPath = Path.GetDirectoryName(originalAppConfigPath);
            originalAppConfigPath = Path.GetDirectoryName(originalAppConfigPath);
            originalAppConfigPath = Path.GetDirectoryName(originalAppConfigPath);
            originalAppConfigPath = Path.GetDirectoryName(originalAppConfigPath);
            originalAppConfigPath = originalAppConfigPath + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlQueryResults" + Path.DirectorySeparatorChar.ToString()
                + "App.config";

            testAppConfigPath = Path.GetTempFileName();
            System.IO.File.Copy(originalAppConfigPath, testAppConfigPath, true);

        }

        [TearDown()]
        public void MyClassCleanup()
        {
            File.Delete(testAppConfigPath);
        }

        #endregion

        [Test()]
        public void GetSingularValueTest()
        {
            // pull up the config file the old fashioned way ;)
            XmlDocument x = new XmlDocument();
            x.Load(testAppConfigPath);
            string currentServer = x.SelectSingleNode("configuration/appSettings/add[@key='Server']").Attributes["value"].Value;

            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);

            Assert.AreEqual(currentServer, target.GetValue("Server"));
        }

        [Test()]
        public void SaveWriteAndRetrieveSingularValueTest()
        {
            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);
            string currentServer = target.GetValue("Server");
            target.SaveValue("Server", currentServer + "a");
            target.WriteConfig(testAppConfigPath);

            Assert.AreEqual(target.GetValue("Server"), currentServer + "a");
        }

        [Test()]
        public void SaveAndRetrieveSingularValueTest()
        {
            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);
            string currentServer = target.GetValue("Server");
            target.SaveValue("Server", currentServer + "a");
            Assert.AreEqual(target.GetValue("Server"), currentServer + "a");
        }
        
        [Test()]
        public void SaveAndRetrieveDictionaryValueTest()
        {
            Dictionary<object, object> firstDictionary = new Dictionary<object, object>();
            firstDictionary.Add((int)1, "a longer string");
            firstDictionary.Add((int)2, @"a messud up
string");
            firstDictionary.Add((int)3, @"a string with some icky characters! <a><&b><c;>");
            firstDictionary.Add("asdf", @"a string with some icky characters! <a><&b><<c;>    ");
            firstDictionary.Add(false, @"true");
            firstDictionary.Add(true, @"false");

            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);
            target.SaveValue(firstDictionary, "test1");

            Dictionary<object, object> retrievedDictionary = target.GetDictionary("test1");
            foreach (object o in retrievedDictionary.Keys)
            {
                object a = firstDictionary[o];
                object b = retrievedDictionary[o];
                Assert.AreEqual(firstDictionary[o], retrievedDictionary[o]);
            }

            Dictionary<object, object> secondDictionary = new Dictionary<object, object>();
            secondDictionary.Add("", "a longer string");
            secondDictionary.Add("a longer string", @"a messud up
string");
            secondDictionary.Add("a string with some icky characters! <a><&b><c;>", @"a string with some icky characters! <a><&b><c;>");

            target = new ConfigManipulator(testAppConfigPath);
            target.SaveValue(secondDictionary, "test2");

            retrievedDictionary = target.GetDictionary("test2");
            foreach (object o in retrievedDictionary.Keys)
            {
                object a = secondDictionary[o];
                object b = retrievedDictionary[o];
                Assert.AreEqual(secondDictionary[o], retrievedDictionary[o]);
            }
        }
    }
}
