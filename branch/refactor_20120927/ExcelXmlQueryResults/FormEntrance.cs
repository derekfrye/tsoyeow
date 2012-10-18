using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using ExcelXmlQueryResults.Properties;
using System.Threading;
using System.Xml.Linq;
using System.Globalization;
using ExcelXmlWriter;

namespace ExcelXmlQueryResults
{
    public partial class FormEntrance : Form
    {
        string fileName;
        string connStr;

        ExcelXmlQueryResultsParams p;
        DateTime workbookStart;

        public FormEntrance()
        {
            InitializeComponent();

            saveFileDialog1.Filter = Resources.SaveDialogFilter;
            this.Text = Resources.Version;

            toolStripStatusLabel2.Text = Resources.Waiting;

            toolTip1.SetToolTip(label3, Resources.TooltipInitialFilename);
            toolTip1.AutoPopDelay = Settings1.Default.toolTipDelayBeforeFade;
            toolTip1.SetToolTip(textBox1, Resources.TooltipInitialFilename);

            ConfigManipulator c = new ConfigManipulator();
            try
            {
                p = FormOptions.LoadOpts();
            }
            catch (Exception e)
            {
                if (e is Exceptions.ConfigFileBroken || e is ArgumentNullException || e is ArgumentException || e is InvalidOperationException)
                {
                    MessageBox.Show(e.Message);
                    if (e is Exceptions.ConfigFileBroken)
                    {
                        if (e.Data.Contains("key"))
                            toolStripStatusLabel2.Text = Resources.AppconfigBroken + " Missing key: " + (string)e.Data["key"];
                        else
                            toolStripStatusLabel2.Text = Resources.AppconfigBroken;
                    }
                    LockUnlockGUIControls(true);
                }
                else
                {
                    MessageBox.Show(e.Message);
                    throw;
                }
            }

            string debugFilePath = Application.ExecutablePath;
            debugFilePath = Path.GetDirectoryName(debugFilePath);
            userFileName.Text = debugFilePath + Path.DirectorySeparatorChar.ToString();
            textBox1.Text = "a.xlsx";
            fileName = userFileName.Text + textBox1.Text;

#if DEBUG

            string path = Environment.CurrentDirectory;
            if(!string.IsNullOrEmpty(path))
            path = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(path))
            path = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(path))
            path = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(path))
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "Sql exceeds filesize limit.sql";
            if (File.Exists(path))
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                StreamReader sr = new StreamReader(fs);
                richTextBox1.Text = sr.ReadToEnd();
                sr.Close();
                fs.Close();
            }
#endif
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text) || !Path.IsPathRooted(userFileName.Text))
            {
                MessageBox.Show(Resources.NoApprFile);
                return;
            }
            if (!userFileName.Text.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                userFileName.Text = userFileName.Text + Path.DirectorySeparatorChar.ToString();
            }

            try
            {
                // create a temp file
                string temp = Path.GetTempFileName();
                // try to copy it to the desired output directory
                string test = Path.GetDirectoryName(userFileName.Text) + Path.DirectorySeparatorChar.ToString()
                    + Path.GetFileName(temp);
                while (true)
                {
                    // if it already exists, make a new temp file and try again
                    if (File.Exists(test))
                    {
                        if (File.Exists(temp))
                            File.Delete(temp);
                        temp = Path.GetTempFileName();
                        test = Path.GetDirectoryName(userFileName.Text) + Path.DirectorySeparatorChar.ToString()
                            + Path.GetFileName(temp);
                    }
                    else
                    {
                        File.Move(temp, test);
                        break;
                    }
                }
                if (File.Exists(temp))
                    File.Delete(temp);
                if (File.Exists(test))
                    File.Delete(test);
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
                return;
            }

            fileName = userFileName.Text + textBox1.Text;
            buildConnStr();
            writeSpreadsheets(); 
        }

        void writeSpreadsheets()
        {

#if DEBUGFROMFILE
            string path = Application.ExecutablePath;
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = Path.GetDirectoryName(path);
            path = path + Path.DirectorySeparatorChar.ToString()
                + "ExcelXmlWriterNTest" + Path.DirectorySeparatorChar.ToString()
                + "Resources" + Path.DirectorySeparatorChar.ToString()
                + "Data.xml";

            p.p.query = path;
            p.p.fromFile = true;
#else
            p.p.Query = richTextBox1.Text;
            p.p.FromFile = false;
#endif
            p.p.ConnectionString = connStr;

            // prevent changes while writing data
            LockUnlockGUIControls(true);
            workbookStart = DateTime.Now;

#if DEBUGNOTHREAD
            if (String.Equals(p.newResultSetMethod, Resources.NewResultSetWorksheet))
                work(new object[] { p, fileName });
            else
                work1(new object[] { p, fileName });
#else
            Thread t;
            if (String.Equals(p.newResultSetMethod, Resources.NewResultSetWorksheet))
                t = new Thread(WriteResultsToSeparateTabs);
            else
                t = new Thread(WriteResultsToSeparateFiles);
            t.IsBackground = true;
            t.Start(new object[] { p, fileName });
#endif
        }

        void WriteResultsToSeparateTabs(object p)
        {
            object[] p2 = (object[])p;
            ExcelXmlQueryResultsParams p1 = (ExcelXmlQueryResultsParams)p2[0];
            string p3 = (string)p2[1];

            // open wb
            Workbook wb = new Workbook(p1.p);

            // subscribe to progress events
            wb.ReaderFinished += new EventHandler<ReaderFinishedEvents>(wb_ReaderFinished);
            wb.QueryStarted += new EventHandler<EventArgs>(wb_QueryStarted);
            wb.QueryException += new EventHandler<QueryExceptionEvents>(wb_QueryError);
            wb.QueryRowsOverTime += new EventHandler<QueryRowsOverTimeEvents>(wb_QueryRowsOverTime);

            if (wb.RunQuery())
            {
                int currentFile = 1;
                
                WorkBookStatus status = wb.WriteQueryResults(p3);
                while (status != WorkBookStatus.Completed)
                {
                    currentFile++;
                    string a = getIncrFileName(currentFile, p3);
                    status = wb.WriteQueryResults(a);
                }
            }
        }

        void WriteResultsToSeparateFiles(object p)
        {
            object[] p2 = (object[])p;
            ExcelXmlQueryResultsParams p1 = (ExcelXmlQueryResultsParams)p2[0];
            string filename = (string)p2[1];
            string orig_filename = (string)p2[1];

            // open wb
            Workbook wb = new Workbook(p1.p);

            // subscribe to progress events
            wb.ReaderFinished += new EventHandler<ReaderFinishedEvents>(wb_ReaderFinished);
            wb.QueryStarted += new EventHandler<EventArgs>(wb_QueryStarted);
            wb.QueryException += new EventHandler<QueryExceptionEvents>(wb_QueryError);
            wb.QueryRowsOverTime += new EventHandler<QueryRowsOverTimeEvents>(wb_QueryRowsOverTime);

            if (wb.RunQuery())
            {
                int currentFile = 1;
                // if we have a name for this file, retrieve it
                if (p1.p.ResultNames.ContainsKey(currentFile))
                    filename = changeFileNameBaseName(filename, p1.p.ResultNames[currentFile]);
                
                while (wb.NextResult())
                {
                    if (currentFile != 1)
                    {
                        // if we have a name for this file, retrieve it
                        if (p1.p.ResultNames.ContainsKey(currentFile))
                            filename = changeFileNameBaseName(filename, p1.p.ResultNames[currentFile]);
                        // otherwise, get the next filename in sequence
                        else
                            filename = getIncrFileName(currentFile, orig_filename);
                    }
                    // write the results
                    WorkBookStatus status = wb.WriteQueryResult(filename);
                    int currentResultSet = 1;
                    // if not all of the results were written (because over max-file size), make a new file and continue writing
                    while (status != WorkBookStatus.Completed)
                    {
                        currentResultSet++;
                        filename = getIncrFileName(currentResultSet, filename);
                        status = wb.WriteQueryResult(filename);
                    }
                    currentFile++;
                }
                wb.QueryClose();
            }
        }

        static string getIncrFileName(int i, string p3)
        {
            return Path.GetDirectoryName(p3)
                + Path.DirectorySeparatorChar.ToString()
                + Path.GetFileNameWithoutExtension(p3)
                + "_"+i.ToString()
                + Path.GetExtension(p3);
        }

        /// <summary>
        /// Turn a fully-qualified filename like C:\a.xml into C:\newfile.xml
        /// </summary>
        /// <param name="currentBaseName"></param>
        /// <param name="newBaseName"></param>
        /// <returns></returns>
        static string changeFileNameBaseName(string currentBaseName, string newBaseName)
        {
            return Path.GetDirectoryName(currentBaseName)
                + Path.DirectorySeparatorChar.ToString()
                + newBaseName
                + Path.GetExtension(currentBaseName);
        }

        void wb_QueryRowsOverTime(object sender, QueryRowsOverTimeEvents e)
        {
            this.Invoke(new IncrSomeResultsD(this.IncrSomeResults), new object[] { e.rowsPerSecond, e.total });
        }

        void wb_QueryError(object sender, QueryExceptionEvents e)
        {
            this.Invoke(new CatchQueryExceptionD(this.CatchQueryException), e);
        }

        void wb_QueryStarted(object sender, EventArgs e)
        {
            this.Invoke(new IncrFinishedD(this.IncrFinished), new object[] { QueryState.Running, null });
        }

        void wb_ReaderFinished(object sender, ReaderFinishedEvents a)
        {
            this.Invoke(new IncrFinishedD(this.IncrFinished), new object[] { QueryState.Finished, a.totalRecordsRead });
        }

        delegate void IncrFinishedD(QueryState b, int total);
        delegate void IncrSomeResultsD(decimal a, int rowsProcessed);
        delegate void CatchQueryExceptionD(QueryExceptionEvents e);

        /// <summary>
        /// Prevent GUI from updates during execution.
        /// </summary>
        /// <param name="lockControls">True to lock controls, false to unlock.</param>
        void LockUnlockGUIControls(bool lockControls)
        {
            splitContainer1.Enabled = !lockControls;
            menuStrip1.Enabled = !lockControls;
            richTextBox1.ReadOnly = lockControls;
        }

        enum QueryState { Running, Finished }

        void IncrFinished(QueryState b, int totalRows)
        {
            if (b == QueryState.Finished)
            {
                DateTime d = DateTime.Now;
                TimeSpan t = d - workbookStart;
                toolStripStatusLabel2.Text = Resources.Finished + " Wrote " + totalRows.ToString("N")
                    + " total rows in " + t.TotalSeconds.ToString() + " seconds.";
                LockUnlockGUIControls(false);
            }
            else
                toolStripStatusLabel2.Text = Resources.RunningNoResults;
        }

        void IncrSomeResults(decimal a, int total)
        {
            if (a > (decimal)0)
            {
                toolStripStatusLabel2.Text = Resources.RunningSomeResults +
                " " + Math.Round(a, 2).ToString("N")
                    + " rows received per second,"
                + " " + total.ToString("N")
                    + " total rows written...";
            }
        }

        void CatchQueryException(QueryExceptionEvents e)
        {
            MessageBox.Show(e.e.Message, "Error");
            LockUnlockGUIControls(false);
            toolStripStatusLabel2.Text = Resources.Waiting;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;
                userFileName.Text = Path.GetDirectoryName(saveFileDialog1.FileName);
                if (!userFileName.Text.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    userFileName.Text += Path.DirectorySeparatorChar.ToString();
                textBox1.Text = Path.GetFileName(fileName);
            }
        }

        private void queryOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormOptions q = new FormOptions();
            if (q.ShowDialog() == DialogResult.OK)
            {
                p = q.ExcelXmlQueryResultsParams;
            }
            else if (q.c1 != null)
            {
                if (q.c1.Data.Contains("key"))
                    toolStripStatusLabel2.Text = Resources.AppconfigBroken + " Missing key: " + (string)q.c1.Data["key"];
                else
                    toolStripStatusLabel2.Text = Resources.AppconfigBroken;
                LockUnlockGUIControls(true);
            }
        }

        private void buildConnStr()
        {
            ConfigManipulator c = new ConfigManipulator();

            SqlConnectionStringBuilder sb = new SqlConnectionStringBuilder();

            try
            {
                sb.DataSource = c.GetValue("Server");
                sb.InitialCatalog = c.GetValue("Database");
                if (String.Equals(c.GetValue("ConnectionMethod"), Resources.SSPIConnection))
                    sb.IntegratedSecurity = true;
                else
                {
                    sb.IntegratedSecurity = false;
                    sb.UserID = c.GetValue("ConnectionUsername");
                    sb.Password = c.GetValue("ConnectionPassword");
                }
            }
            catch
            {
                MessageBox.Show(Resources.AppconfigBroken);
            }

            connStr = sb.ConnectionString;
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void openConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}