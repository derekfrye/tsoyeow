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
        // holds standard app.config values from ExcelXmlQueryResults - setup in class init
        string testAppConfigPath = Path.GetTempPath();

        #region Test Setup/Teardown

        [SetUp()]
        public void MyClassInitialize()
        {
            originalAppConfigPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
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

        /// <summary>
        /// Check that, once loaded, ConfigManipulator's value for "server" 
        /// matches the value from app.config of ExcelXmlQueryResults
        /// </summary>
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

        /// <summary>
        /// Check that we can overwrite value for "server" 
        /// Tests ConfigManipulator functions SaveValue, WriteConfig, and GetValue
        /// </summary>
        [Test()]
        public void SaveWriteAndRetrieveSingularValueTest()
        {
            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);
            string currentServer = target.GetValue("Server");
            target.SaveValue("Server", currentServer + "a");
            target.WriteConfig(testAppConfigPath);

            Assert.AreEqual(target.GetValue("Server"), currentServer + "a");
        }

        /// <summary>
        /// Check that we can overwrite value for "server" 
        /// Tests ConfigManipulator functions SaveValue and GetDictionary
        /// </summary>
        [Test()]
        public void SaveAndRetrieveDictionaryValueTest()
        {
            Dictionary<object, object> firstDictionary = new Dictionary<object, object>
            {
                { (int)1, "a longer string" },
                { (int)2, @"a messud up
string" },
                { (int)3, @"a string with some icky characters! <a><&b><c;>" },
                { "asdf", @"a string with some icky characters! <a><&b><<c;>    " },
                { false, @"true" },
                { true, @"false" }
            };

            ConfigManipulator target = new ConfigManipulator(testAppConfigPath);
            target.SaveValue(firstDictionary, "test1");

            Dictionary<object, object> retrievedDictionary = target.GetDictionary("test1");
            foreach (object o in retrievedDictionary.Keys)
            {
                object a = firstDictionary[o];
                object b = retrievedDictionary[o];
                Assert.AreEqual(firstDictionary[o], retrievedDictionary[o]);
            }

            Dictionary<object, object> secondDictionary = new Dictionary<object, object>
            {
                { "", "a longer string" },
                { "a longer string", @"a messud up
string" },
                { "a string with some icky characters! <a><&b><c;>", @"a string with some icky characters! <a><&b><c;>" }
            };

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
