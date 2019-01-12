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
using ExcelXmlWriter.Workbook;
using System.Text.RegularExpressions;

namespace ExcelXmlQueryResults
{
    public partial class FormEntrance : Form
    {
        string fileName;
        string connStr;

        WorkBookParams p;
        DateTime workbookStart;

        delegate void IncrFinishedD(QueryState b, int total, string msg);
        delegate void IncrSomeResultsD(decimal a, int rowsProcessed);
        delegate void CatchQueryExceptionD(QueryExceptionEvents e);
        delegate void DSaveFileBegan(QueryState b, int total, string msg);

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

        /// <summary>
        /// Make sure the filename specified is valid and we can write to that directory.
        /// </summary>
        void ValidateOutputFile()
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
        }

        private void runQueriesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ValidateOutputFile();
            buildConnStr();
            var tt=OpenFiles();
            if (tt.Length > 0)
            {
                LockUnlockGUIControls(true);
                Thread t;
                t = new Thread(delegate() { this.processBatches(tt); });
                t.IsBackground = true;
                t.Start();
            }
        }

       internal void processBatches(string[] tt)
        {
            int currentFileCount = 1;
            foreach (var qry in tt)
            {
                using (StreamReader sr = new StreamReader(new FileStream(qry, FileMode.Open, FileAccess.Read)))
                {
                    p.Query = sr.ReadToEnd();
                    p.FromFile = false;
                    p.ConnectionString = connStr;
                    string a = Utility.getIncrFileName(currentFileCount, fileName);
                    WriteResultsToSeparateTabs(new ExcelXmlQueryResultsParams() { e = p, filenm = a });
                    currentFileCount = currentFileCount + 1;                    
                }
            }

        }

        /// <summary>
        /// Load an array with string paths to all the files a user has selected.
        /// </summary>
        /// <returns>The files.</returns>
        string[] OpenFiles()
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                return openFileDialog1.FileNames;
            }
            else
                return new string[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {

            ValidateOutputFile();            
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

            p.Query = path;
            p.FromFile = true;
#else
            p.Query = richTextBox1.Text;
            p.FromFile = false;
#endif
            p.ConnectionString = connStr;

            // prevent changes while writing data
            LockUnlockGUIControls(true);
            workbookStart = DateTime.Now;

#if DEBUGNOTHREAD

                WriteResultsToSeparateTabs(new ExcelXmlQueryResultsParams() { e=p, filenm=fileName });
#else
            Thread t;
            t = new Thread(delegate() { WriteResultsToSeparateTabs(new ExcelXmlQueryResultsParams() { e = p, filenm = fileName }); });
            t.IsBackground = true;
            t.Start();
#endif
        }


        /// <summary>
        /// Write results to separate tabs. 
        /// If MaximumResultSetsPerWorkbook > count of select statements, it'll create auto-numbered output files.
        /// </summary>
        /// <param name="p"></param>
        void WriteResultsToSeparateTabs(ExcelXmlQueryResultsParams p)
        {

            WorkBookParams workbookParams = p.e;
            string destinationFileName = p.filenm;

            // open wb
            Workbook wb = new Workbook(workbookParams);

            // subscribe to progress events
            wb.ReaderFinished += new EventHandler<ReaderFinishedEvents>(wb_ReaderFinished);
            wb.QueryStarted += new EventHandler<EventArgs>(wb_QueryStarted);
            wb.QueryException += new EventHandler<QueryExceptionEvents>(wb_QueryError);
            wb.QueryRowsOverTime += new EventHandler<QueryRowsOverTimeEvents>(wb_QueryRowsOverTime);
            wb.SaveFile += new EventHandler<SaveFileEvent>(wb_SaveBegan);

            if (wb.RunQuery())
            {
                int currentFileCount = 0;

                var r = new Regex(@"(select)\s+", RegexOptions.IgnoreCase);
                var m = r.Match(workbookParams.Query);
                var countOfResultSets = r.Matches(workbookParams.Query).Count;

                // append "_000..." to the filename if we're writing more than workbook
                // if we're writing < 10 results, we'll rename them _0, _1, ... up to _9
                // if we're writing > 10 but < 100 results, we'll rename them _00, _01, ... up to _99
                // etc.
                var padleft = 0;
                var modifiableFileName = destinationFileName;
                if (workbookParams.MaximumResultSetsPerWorkbook < countOfResultSets)
                {
                    padleft = countOfResultSets.ToString().Length + 1;
                    currentFileCount = currentFileCount + 1;
                    modifiableFileName = Utility.getIncrPaddedFileName(currentFileCount, destinationFileName, padleft);
                }

                WorkBookStatus status = wb.WriteQueryResults(modifiableFileName);
                while (status != WorkBookStatus.Completed)
                {
                    currentFileCount = currentFileCount + 1;

                    if (workbookParams.MaximumResultSetsPerWorkbook < countOfResultSets)
                        modifiableFileName = Utility.getIncrPaddedFileName(currentFileCount, destinationFileName, padleft);
                    else
                        modifiableFileName = Utility.getIncrFileName(currentFileCount, destinationFileName);

                    status = wb.WriteQueryResults(modifiableFileName);
                }
            }
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
            this.Invoke(new IncrFinishedD(this.IncrFinished), new object[] { QueryState.Running, null, null });
        }

        void wb_ReaderFinished(object sender, ReaderFinishedEvents a)
        {
            this.Invoke(new IncrFinishedD(this.IncrFinished), new object[] { QueryState.Finished, a.totalRecordsRead, null });
        }

        void wb_SaveBegan(object sender, SaveFileEvent s)
        {
            this.Invoke(new DSaveFileBegan(this.IncrFinished), new object[] { QueryState.Saving, 0, s.Message });
        }

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

        void IncrFinished(QueryState b, int totalRows, string msg)
        {
            if (b == QueryState.Finished)
            {
                DateTime d = DateTime.Now;
                TimeSpan t = d - workbookStart;
                toolStripStatusLabel2.Text = Resources.Finished + " Wrote " + totalRows.ToString("N")
                    + " total rows in " + Math.Round(t.TotalSeconds, 3, MidpointRounding.AwayFromZero).ToString() + " seconds.";
                LockUnlockGUIControls(false);
            }
            else if (b == QueryState.Saving && !string.IsNullOrEmpty(msg))
            {
                toolStripStatusLabel2.Text = msg;
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

        
    }

    
}