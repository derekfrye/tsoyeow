using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ExcelXmlQueryResults.Properties;

namespace ExcelXmlQueryResults
{

    public class ConfigManipulator
    {
        XDocument xd;

        public ConfigManipulator()
        {
            xd = XDocument.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
                , LoadOptions.PreserveWhitespace);
        }

        public ConfigManipulator(string path)
        {
            xd = XDocument.Load(path
                , LoadOptions.PreserveWhitespace);
        }

        public string GetValue(string key)
        {
            string a = string.Empty;
            try
            {
                a = xd.Elements().DescendantsAndSelf("add").Attributes("key")
                    .Where(x => String.Equals(x.Value, key)).First().Parent.Attribute("value").Value;
            }
            catch (InvalidOperationException)
            {
                Exceptions.ConfigFileBroken c = new Exceptions.ConfigFileBroken(Resources.AppconfigBroken + " Missing key: " + key);
                c.Data.Add("key", key);
                throw c;
            }
            return a;
        }

        public Dictionary<object, object> GetDictionary(string name, Type KeyType, Type ValueType)
        {
            return GetDictionary(name, KeyType, true, ValueType, true);
        }

        public Dictionary<object, object> GetDictionary(string name)
        {
            return GetDictionary(name, null, false, null, false);
        }

        Dictionary<object, object> GetDictionary(string name, Type KeyType, bool onlyKeyTypeKeys, Type ValueType, bool onlyValueTypeValues)
        {
            Dictionary<object, object> a = new Dictionary<object, object>();
            try
            {
                XElement x= xd.Element("configuration").Element(Resources.StateSettingsConfigSection).Descendants(name).First();
                foreach (XElement x1 in x.Elements())
                {
                    Type t1 = Type.GetType(x1.Element("key").Attribute("keytype").Value);
                    Type t2 = Type.GetType(x1.Element("value").Attribute("valuetype").Value);
                    if (!a.ContainsKey(Convert.ChangeType(x1.Element("key").Value, t1)))
                        if (!((onlyKeyTypeKeys && t1 != KeyType) || (onlyValueTypeValues && t2 != ValueType)))
                            a.Add(Convert.ChangeType(x1.Element("key").Value, t1)
                                , Convert.ChangeType(x1.Element("value").Value, t2));
                }
            }
            catch (Exception e)
            {
                if (e is InvalidOperationException || e is ArgumentNullException)
                {
                    Exceptions.ConfigFileBroken c = new Exceptions.ConfigFileBroken(Resources.AppconfigBroken 
                        + " Bad/missing value in " + name);
                    c.Data.Add("name", name);
                    throw c;
                }
                else
                    throw;
            }
            return a;
        }

        public void SaveValue(string key, string value)
        {
            XAttribute b;
            XElement c;

            b = xd.Element("configuration").Element("appSettings").Descendants("add").Attributes("key")
                    .Where(x => String.Equals(key, x.Value)).First();
            c = b.Parent;
            c.Attribute("value").SetValue(value);

        }

        public void SaveValue(Dictionary<object, object> c1, string name)
        {
            XElement c;

            if (!xd.Element("configuration").Element(Resources.StateSettingsConfigSection).Elements().Any(x => x.Name == name))
                xd.Element("configuration").Element(Resources.StateSettingsConfigSection).LastNode.AddAfterSelf(new XElement(name));

            c = xd.Element("configuration").Element(Resources.StateSettingsConfigSection).Descendants(name).First();
            c.RemoveNodes();

            foreach (object o in c1.Keys)
                c.Add(
                    new XElement("entry"
                        , new XElement("key"
                            , new XAttribute("keytype", o.GetType().ToString())
                            , o.ToString()
                        )
                        , new XElement("value"
                            , new XAttribute("valuetype", c1[o].GetType().ToString())
                            , c1[o].ToString()
                        )
                    )
                );
        }

        /// <summary>
        /// Persist changes of app.config to disk.
        /// </summary>
        public void WriteConfig()
        {
            xd.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, SaveOptions.None);
            xd = XDocument.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile
                , LoadOptions.PreserveWhitespace);
        }

        public void WriteConfig(string path)
        {
            xd.Save(path, SaveOptions.None);
            xd = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        }
    }
}
