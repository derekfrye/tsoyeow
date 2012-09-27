using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using ExcelXmlQueryResults.Properties;
using ExcelXmlWriter;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ExcelXmlQueryResults
{

    struct ExcelXmlQueryResultsParams
    {
        public WorkBookParams p;
        public string newResultSetMethod;
    }

    internal partial class FormOptions : Form
    {

        internal Exceptions.ConfigFileBroken c1;
        ExcelXmlQueryResultsParams p;

        internal FormOptions()
        {
            InitializeComponent();

            tabControl1.Visible = false;

            foreach (TreeNode n in treeView1.Nodes)
            {
                n.Expand();
            }

            // populate login method opts
            comboBox1.Items.Add(Resources.SSPIConnection);
            comboBox1.Items.Add(Resources.UsernamePasswordConnection);
            
            foreach (DataGridViewColumn c1 in resultSetNamesGrid.Columns)
            {
                c1.HeaderCell.ToolTipText = Resources.TooltipResultNames;
            }

            toolStripStatusLabel1.Text = string.Empty;

            ConfigManipulator c = new ConfigManipulator();

            try
            {
                p = LoadOpts();

                textBox1.Text = c.GetValue("Server");
                textBox2.Text = c.GetValue("Database");
                textBox3.Text = c.GetValue("ConnectionUsername");
                textBox4.Text = c.GetValue("ConnectionPassword");

                checkBox1.Checked = p.p.writeEmptyResultSetColumns;
                checkBox2.Checked = p.p.AutoRewriteOverpunch;

                textBox5.Text = p.p.queryTimeout.ToString();
                textBox6.Text = p.p.maxRowsPerSheet.ToString();

                var p1 = p.p.resultNames;
                int count = 0;
                foreach (object o in p1.Keys)
                {
                    DataGridViewRow r = new DataGridViewRow();
                    resultSetNamesGrid.Rows.Add();
                    resultSetNamesGrid.Rows[count].Cells[0].Value = o;
                    resultSetNamesGrid.Rows[count].Cells[1].Value = p1[(int)o];
                    count++;
                }

                if (String.Equals(c.GetValue("ExcelFileType"), Resources.FileTypeXml))
                    comboBox3.SelectedIndex = 1;
                else
                    comboBox3.SelectedIndex = 0;

                if (String.Equals( c.GetValue("ConnectionMethod"), Resources.SSPIConnection))
                    comboBox1.SelectedItem = Resources.SSPIConnection;
                else
                    comboBox1.SelectedItem = Resources.UsernamePasswordConnection;

                if (String.Equals(c.GetValue("NewResultSet"), Resources.NewResultSetWorksheet))
                {
                    radioButton1.Checked = true;
                    radioButton2.Checked = false;
                }
                else
                {
                    radioButton1.Checked = false;
                    radioButton2.Checked = true;
                }
            }
            catch (Exceptions.ConfigFileBroken e)
            {
                MessageBox.Show(e.Message);
                if (e.Data.Contains("key"))
                    toolStripStatusLabel1.Text = Resources.AppconfigBroken + " Missing key: " + (string)e.Data["key"];
                else
                    toolStripStatusLabel1.Text = Resources.AppconfigBroken;
                panel7.Enabled = false;
                c1 = e;
            }
        }

        internal static ExcelXmlQueryResultsParams LoadOpts()
        {
            ConfigManipulator c = new ConfigManipulator();
            ExcelXmlQueryResultsParams a = new ExcelXmlQueryResultsParams();
            a.p = new WorkBookParams();
            a.newResultSetMethod = c.GetValue("NewResultSet");
            a.p.writeEmptyResultSetColumns = Convert.ToBoolean(c.GetValue("WriteEmptyResultColumnHeaders"));
            a.p.AutoRewriteOverpunch = Convert.ToBoolean(c.GetValue("AutoRewriteOverpunch"));
            a.p.backendMethod = Enum.GetValues(typeof(ExcelBackend))
                            .Cast<ExcelBackend>()
                            .Where(x => String.Equals(x.ToString(), c.GetValue("ExcelFileType"))).First();

            int res = 0;
            if (!Int32.TryParse(c.GetValue("MaxRowsPerSheet"), out res))
                a.p.maxRowsPerSheet = Convert.ToInt32(Resources.DefaultMaxRowsPerSheet);
            else
                a.p.maxRowsPerSheet = Convert.ToInt32(c.GetValue("MaxRowsPerSheet"));
            if (Int32.TryParse(c.GetValue("QueryTimeout"), out res))
                a.p.queryTimeout = Convert.ToInt32(c.GetValue("QueryTimeout"));
            
            var p1 = c.GetDictionary("ResultNames", typeof(int), typeof(string));
            foreach (object o in p1.Keys)
            {
                a.p.resultNames.Add(Convert.ToInt32(o)
                    , p1[o].ToString());
            }

            return a;
        }

        #region Properties

        internal ExcelXmlQueryResultsParams ExcelXmlQueryResultsParams
        {
            get { return p; }
        }

        #endregion

        /// <summary>
        /// Save values to form state and app.config.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            ConfigManipulator c = new ConfigManipulator();

            Dictionary<string, string> h = new Dictionary<string, string>();
            h.Add("Server", textBox1.Text);
            h.Add("Database", textBox2.Text);
            h.Add("ConnectionMethod", String.Equals(comboBox1.SelectedItem.ToString(), Resources.SSPIConnection)
                ? Resources.SSPIConnection : Resources.UsernamePasswordConnection);
            h.Add("NewResultSet", radioButton1.Checked
                ? Resources.NewResultSetWorksheet : Resources.NewResultSetWorkbook);
            h.Add("ConnectionUsername", textBox3.Text);
            h.Add("ConnectionPassword", textBox4.Text);
            h.Add("QueryTimeout", textBox5.Text);
            h.Add("MaxRowsPerSheet", textBox6.Text);
            
            if (Regex.IsMatch(comboBox3.SelectedItem.ToString(), Resources.FileTypeXml, RegexOptions.IgnoreCase))
                h.Add("ExcelFileType", Resources.FileTypeXml);
            else
                h.Add("ExcelFileType", Resources.FileTypeXlsx);
            h.Add("WriteEmptyResultColumnHeaders", checkBox1.Checked.ToString());
            h.Add("AutoRewriteOverpunch", checkBox2.Checked.ToString());

            Dictionary<object, object> d1 = new Dictionary<object, object>();

            Dictionary<object, object> d2 = new Dictionary<object, object>();
            foreach (DataGridViewRow d in resultSetNamesGrid.Rows.Cast<DataGridViewRow>().Where(x => x.Cells[0].Value != null))
            {
                d2.Add(Convert.ToInt32(d.Cells[0].Value.ToString()), d.Cells[1].Value.ToString());
            }
            c.SaveValue(d2, "ResultNames");

            foreach (string i in h.Keys)
            {
                c.SaveValue(i, h[i]);
            }

            c.WriteConfig();
            p = LoadOpts();
            this.DialogResult = DialogResult.OK;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (String.Equals(comboBox1.SelectedItem.ToString(), Resources.SSPIConnection))
            {
                textBox3.Enabled = false;
                textBox4.Enabled = false;
            }
            else
            {
                textBox3.Enabled = true;
                textBox4.Enabled = true;
            }
        }

        /// <summary>
        /// Auto-number result sets.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // if the first cell is null and the next cell has a value, autonumber the first cell
            if (resultSetNamesGrid.Rows[e.RowIndex].Cells[0].Value == null && resultSetNamesGrid.Rows[e.RowIndex].Cells[1].Value != null)
            {
                int maxVal = nextAutoVal(resultSetNamesGrid);
                resultSetNamesGrid.Rows[e.RowIndex].Cells[0].Value = maxVal;
            }
        }

        /// <summary>
        /// First gap in the numbers starting from 1. For example, for 1, 2, 4, 5, returns 3
        /// </summary>
        /// <param name="dataGridView2"></param>
        /// <returns></returns>
        static int nextAutoVal(DataGridView dataGridView2)
        {
            int toss=0;
            var p = dataGridView2.Rows.Cast<DataGridViewRow>()
                .Where(x => x.Cells[0].Value != null && Int32.TryParse(x.Cells[0].Value.ToString(), out toss) && toss > 0)
                .OrderBy(x => Convert.ToInt32(x.Cells[0].Value));

            int maxVal = 1;
            if (p.Count() > 0 && Convert.ToInt32(p.First().Cells[0].Value) == 1)
            {
                int prev = Convert.ToInt32(p.First().Cells[0].Value);
                for (int i = 0; i < p.Count(); i++)
                {
                    int current = Convert.ToInt32(p.ElementAt(i).Cells[0].Value);
                    if (current - prev > 1)
                    {
                        maxVal = prev + 1;
                        break;
                    }
                    else
                    {
                        maxVal = current + 1;
                        prev = current;
                    }
                }
            }
            else
                maxVal = 1;

            return maxVal;

        }

        /// <summary>
        /// Make sure first column of result set names is numeric.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == 0 && !String.IsNullOrEmpty(e.FormattedValue.ToString()))
            {
                int res = 0;
                if (!Int32.TryParse(e.FormattedValue.ToString(), out res))
                {
                    MessageBox.Show("Non-integer: " + e.FormattedValue.ToString());
                    e.Cancel = true;
                }
            }
            else if (e.ColumnIndex == 1 && e.FormattedValue.ToString().Length>25)
            {
               MessageBox.Show("Sheet names cannot exceed 25 characters: " + e.FormattedValue.ToString() + " is "+e.FormattedValue.ToString().Length.ToString()+" chars.");
                    e.Cancel = true;
                
            }
        }

        private void textBoxValidateInt(object sender, CancelEventArgs e)
        {
            int res = 0;
            TextBox b = (TextBox)sender;
            if (!Int32.TryParse(b.Text, out res))
            {
                MessageBox.Show("Non-integer:" + b.Text);
                e.Cancel = true;
            }
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                // find current row index
                int row = resultSetNamesGrid.SelectedCells[0].OwningRow.Index;
                // find current row's autonumber
                object selectedRowsAutoNumber = resultSetNamesGrid.Rows[row].Cells[0].Value;

                // read the clipboard values
                var vals = Clipboard.GetText(TextDataFormat.Text).Split(Environment.NewLine.ToCharArray()
                    , StringSplitOptions.RemoveEmptyEntries).Where(x => x.Split("\t".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Length > 0);

                // wipe out the autonumbers for each row after the selected row
                for (int i = row; i < Math.Min(resultSetNamesGrid.Rows.Count, row + vals.Count()); i++)
                {
                    resultSetNamesGrid.Rows[i].Cells[0].Value = null;
                }

                // add the total new rows we need to accomodate the clipboard length
                int newRowsNeeded = resultSetNamesGrid.Rows.Count - row;
                if (resultSetNamesGrid.Rows.Count - row < vals.Count())
                    for (int i = 0; i <= vals.Count() - newRowsNeeded; i++)
                        resultSetNamesGrid.Rows.Add();

                for (int i = 0; i < vals.Count(); i++)
                {
                    string[] clipboardContentsArray = vals.ElementAt(i).Split("\t".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                    // starting at the selected row, add in each clipboard item
                    int rowToModify = i + row;

                    if (!string.IsNullOrEmpty(clipboardContentsArray[0]))
                    {
                        string val = Regex.Replace(clipboardContentsArray[0], @"[\u0000-\u001F,\u007F,\u0080-\u009F]", string.Empty);
                        if (clipboardContentsArray[0].Length > 25)
                            resultSetNamesGrid.Rows[rowToModify].Cells[1].Value = val.Substring(0, 25);
                        else
                            resultSetNamesGrid.Rows[rowToModify].Cells[1].Value = val;
                    }
                    if (resultSetNamesGrid.Rows[rowToModify].Cells[0].Value == null)
                    {
                        int maxVal;
                        if (selectedRowsAutoNumber == null)
                            maxVal = nextAutoVal(resultSetNamesGrid);
                        else
                        {
                            maxVal = Convert.ToInt32(selectedRowsAutoNumber); 
                            selectedRowsAutoNumber = null;
                        }
                        resultSetNamesGrid.Rows[rowToModify].Cells[0].Value = maxVal;
                    }
                }
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(sender, null);
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

            Control p = new Control();

            switch (e.Node.Name)
            {
               
                case "Node1":
                    p = tabnamepanel;
                    break;
                case "Node5":
                    p = sqloptionspanel;
                    break;
                case "Node3":
                case "Node0":
                default:
                    p = excelformatpanel;
                    break;
            }

            p.Parent = tabControl1.Parent;
            p.Location = tabControl1.Location;
            p.Visible = true;

            foreach (Control a in new Control[] { excelformatpanel, tabnamepanel, sqloptionspanel })
            {
                if (a != p)
                {
                    a.Visible = false;
                }
            }
        }
    }
}