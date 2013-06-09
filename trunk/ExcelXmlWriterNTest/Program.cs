using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace ExcelXmlWriterNTest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string[] a = new[] { Assembly.GetExecutingAssembly().Location };
            NUnit.Gui.AppEntry.Main(a);
        }
    }
}
