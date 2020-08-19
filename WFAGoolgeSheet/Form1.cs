using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Media;
using System.IO;
using System.Threading;
using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using System.Text.RegularExpressions;

using Microsoft.Azure.Management.Monitor.Fluent.ActionGroup.Definition;

namespace WFAGoolgeSheet
{
    public partial class Form1 : Form

    {
        public DataGridView dataGridView { get; set; }
        public class sTabName
        {
            public string tabname { get; set; }
            public string range { get; set; }
            public string keyfield { get; set; }
            public override string ToString() { return this.tabname; }
        }

        public Form1()
        {
            InitializeComponent();

            comboBox1.Items.Add(new sTabName { tabname = "Imported Names", range = "!A1:G", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "Field Service", range = "!A1:I", keyfield = "Field Service" });
            comboBox1.Items.Add(new sTabName { tabname = "Confirmed English", range = "!A1:H", keyfield = "Telephone" });
            comboBox1.Items.Add(new sTabName { tabname = "Contacted 5 times letters", range = "!A1:H", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "Only Spanish", range = "!A1:H", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "List of Last names to work", range = "!A1:G", keyfield = "Name" });
            //comboBox1.Items.Add(new sTabName { tabname = "Common First Names", range = "!A1:G", keyfield = "TELEPHONE" });
            comboBox1.SelectedIndex = 1;
        }

        public DataGridView myDG { get { return dataGridView1; } }

        //---------------------------------------------------------------------------//
        // global flags and data
        //
        bool updateinprogress = false;
        bool DataChanged = false;
        bool isProcessRunning = false;
        int dataLoadForSheet = -1;
        bool waiting = false;
        int rowOffset = 0;
        int firstrow = 0;
        int skiprow = 0;
        int progress = 0;
        int rcount = -1;
        public static string myVar = null;  // passing data between forms
        public static object lockMethod2 = new object();
        DialogResult dr = new DialogResult();
        bool done = true;
        public int totalRow = 0;
        public int SecondFromTop;
        public int SecondFormLeft;
        public System.Drawing.Size SecondFormSize = System.Drawing.Size.Empty;


        //
        // list of all datagridview changes
        //
        List<List<String>> cellch = new List<List<String>>(); //Creates new nested List

        //
        // Google parameters
        static UserCredential credential;               // Sheet Service
        static string ApplicationName = null;
        static String spreadsheetId = null;             // Spreadsheet ID




        //---------------------------------------------------------------------------
        //
        // close application and exit
        //
        private void button1_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous 3 dialogs.
                //
                if (result1 == DialogResult.No || result1 == DialogResult.Retry)
                    return;
            }
            Application.Exit();
        }

        //---------------------------------------------------------------------------
        //
        // run application 
        //
        private void button2_Click(object sender, EventArgs e)
        {
            //if (dataLoadForSheet == comboBox1.SelectedIndex) goto LB1;
            int NumofRec = 0;
            int r1 = -1;
            if (string.IsNullOrEmpty(textBox3.Text)) firstrow = 0;  // setup x of n grouping
            else firstrow = Convert.ToInt32(textBox3.Text);
            if (string.IsNullOrEmpty(textBox5.Text)) skiprow = 0;
            else skiprow = Convert.ToInt32(textBox5.Text);
            firstrow = skiprow + 1 - firstrow;                       // invert it!

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show(" No Sheet Tab selected");
                return;
            }
            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                MessageBox.Show(" No Sheet [test/live] selected");
                return;
            }
            if (comboBox1.SelectedIndex == 1)
            {
                if (firstrow == 0 || skiprow == 0 || firstrow > skiprow)
                {
                    MessageBox.Show(" group range is bad ");
                    return;
                }

                bool found = false;
                for (int k = 0; k < checkedListBox1.Items.Count; k++)
                    if (checkedListBox1.GetItemChecked(k))
                    {
                        found = true;
                        break;
                    }
                if (!found)
                {
                    MessageBox.Show(" no filters selected ");
                    return;
                }
            }
            if (DataChanged)
            {
                DialogResult result1 = MessageBox.Show(" There are unsaved changes \n Do you want to Re-Run and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous dialog.
                //
                if (result1 == DialogResult.No || result1 == DialogResult.Retry)
                    return;
            }

            DataChanged = false;
            updateinprogress = true;
            startPB(System.Drawing.Color.Green);
            textBox1.Text = " ..logging in ";
            textBox1.Update();
            this.Controls.Add(dataGridView1);
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // If modifying these scopes, delete your previously saved credentials
            // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            ApplicationName = "Google Sheets API .NET Quickstart";

            //UserCredential credential;

            //Form1 child = new Form1();
            DataBindings.Clear();

            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = false;
            //dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.GridColor = System.Drawing.Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Navy;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font =
                new Font(dataGridView1.Font, FontStyle.Bold);
            dataGridView1.Name = "dataGridView1";
            //dataGridView1.Location = new Point(8, 8);
            //dataGridView1.Size = new Size(500, 250);
            dataGridView1.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


            Cursor.Current = Cursors.WaitCursor;
            int v = dataGridView1.VerticalScrollingOffset;
            int h = dataGridView1.HorizontalScrollingOffset;

            using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            textBox1.Text = "..connecting Sheet ";
            textBox1.Update();
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            //spreadsheetId = "1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI";
            string spreadsheetId1 = "1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok";    // test sheet
            string spreadsheetId2 = "1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc";    // live sheet  /edit#gid=2145664999"

            //
            // process RadioButtons on Form1
            //var checkedRadio = new[] { groupBox1 }.SelectMany(g => g.Controls.OfType<RadioButton>()
            //                    .Where(r => r.Checked));
            if (radioButton1.Checked) spreadsheetId = spreadsheetId1;
            if (radioButton2.Checked) spreadsheetId = spreadsheetId2;

            int selectIndex = comboBox1.SelectedIndex;
            sTabName selectCar = (sTabName)comboBox1.SelectedItem;
            //String range = "List of Last Names to work!A1:G";
            String range = selectCar.tabname + selectCar.range;

            Form1.ActiveForm.Text = String.Format("Working with sheet '{0}', columns {1} ", spreadsheetId, range);
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            // https://docs.google.com/spreadsheets/d/1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI/edit#gid=1072334158
            // https://docs.google.com/spreadsheets/d/1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok/edit#gid=2145664999
            // https://docs.google.com/spreadsheets/d/1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc/edit#gid=2145664999
            ValueRange response = request.Execute();
            textBox1.Text = ".. reading data";
            textBox1.Update();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                // find the header row and how many columns the sheet has
                rcount = 0;
                int a = 0;
                foreach (var row in values)
                {
                    if (row.Count > rcount)
                    {
                        rcount = row.Count;
                        rowOffset = a + 1;
                    }
                    if (a++ > 10) break;
                }
                NumofRec = 0;

                //
                //  Build Header Columns
                //
                foreach (var row in values)
                {
                    if (NumofRec == 0)
                    {
                        //for (int z = 0; z < row.Count; z++)
                        for (int z = 0; z < rcount; z++)
                        {
                            if (row.Count < rcount) break;
                            dataGridView1.Columns.Add(row[z].ToString(), row[z].ToString());
                        }
                        if (row.Count < rcount) continue;

                        dataGridView1.Visible = true;
                        dataGridView1.Update();
                        NumofRec++;
                        continue;
                    }

                    // now build the data
                    dataGridView1.Visible = false;
                    int x = dataGridView1.Rows.Add();

                    // show progress
                    int remainder;
                    Math.DivRem(NumofRec, values.Count, out remainder);
                    if (values.Count > 100)
                        progress = remainder / (values.Count / 100);
                    else
                        progress = remainder * (100 - 1 / values.Count);
                    if (progress < 100) progressBar1.Value = progress;
                    if (remainder % 139 == 1)
                    {
                        textBox1.Text = "adding .. " + NumofRec.ToString();
                        textBox1.Update();
                        progressBar1.Update();
                    }

                    NumofRec++;
                    for (int z = 0; z < row.Count; z++)
                    {
                        dataGridView1.CurrentCell = dataGridView1[z, x];
                        dataGridView1.CurrentCell.Value = row[z];
                        dataGridView1.Rows[x].Cells[z].Selected = true;
                    }

                    if (r1 == -1) r1 = firstrow;
                    if (comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 0)
                    {
                        dataGridView1.CurrentRow.Visible = false;
                        string testStr = dataGridView1.Rows[x].Cells[4].Value?.ToString();
                        foreach (string kind in checkedListBox1.CheckedItems)
                        {
                            //if ((r1 % skiprow) == 0)
                            //{
                            if ((testStr == kind || (string.IsNullOrEmpty(testStr) && kind == "blank") ||
                            (string.IsNullOrWhiteSpace(testStr) && kind == "blank")
                            && ((r1++ % skiprow) == 0)))
                            {
                                dataGridView1.Rows[x].Visible = true;
                                break;
                            }
                            //}
                            //else dataGridView1.Rows[x].Visible = false;
                        }
                    }
                }
            }
            else
            {
                textBox1.Text = "No data found.";
                textBox1.Update();
                DataChanged = false;
                updateinprogress = true;
                dataGridView1.Visible = true;
                return;
            }
            //

            for (int y = 0; y < dataGridView1.Rows.Count; y++)
                if (dataGridView1.Rows[y].Visible == true)
                {
                    dataGridView1.Rows[y].Cells[0].Selected = true;
                    break;
                }

            dataGridView1.HorizontalScrollingOffset = h;
            Cursor.Current = Cursors.Default;
            //dataGridView1.CurrentCell = null; ;
            dataGridView1.Visible = true;
            dataGridView1.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.SortMode = DataGridViewColumnSortMode.NotSortable);
            totalRow = dataGridView1.Rows.Count - 1;
            textBox1.Text = (totalRow).ToString() + " records";
            dataGridView1.Update();
            dataGridView1.Show();
            Thread.Sleep(25);
            updateinprogress = false;
            DataChanged = false;
            cellch.Clear();
            progressBar1.Value = 100;

            if (comboBox1.SelectedIndex == 0)
            {
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                button8.Visible = true;
                textBox6.Visible = true;
                textBox7.Visible = true;
                textBox8.Visible = true;
                button9.Visible = true;
            }

            if (comboBox1.SelectedIndex == 1)
            {
                button9.Visible = false;
            }

            if (comboBox1.SelectedIndex == 2)
            {
                button9.Visible = false;
            }
            if (comboBox1.SelectedIndex == 3)
            {
                button9.Visible = false;
            }
            if (comboBox1.SelectedIndex == 4)
            {
                button9.Visible = false;
            }
            if (comboBox1.SelectedIndex == 5)
            {
                button9.Visible = false;
            }
        }


        //---------------------------------------------------------------------------
        //
        // Clear Data 
        //
        private void button3_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = MessageBox.Show(" There are unsaved changes \n Do you want to Clear and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous 3 dialogs.
                //
                if (result1 == DialogResult.No || result1 == DialogResult.Retry)
                    return;
            }
            updateinprogress = true;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            textBox1.Text = (dataGridView1.Rows.Count).ToString() + " records";
            cellch.Clear();
            //checkedListBox1.Visible = false;
            checkedListBox1.ResetText();
            textBox2.Text = " none selected";
            textBox2.Update();
            DataChanged = false;
            dataGridView1.Refresh();
            updateinprogress = false;
            Form1.ActiveForm.Text = "No Active Sheet";
            isProcessRunning = false;
            progressBar1.Value = 0;
        }

        //---------------------------------------------------------------------------
        //
        // Log out
        //
        private void button5_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous 3 dialogs.
                //
                if (result1 == DialogResult.Yes)
                    Application.Exit();
            }
            cellch.Clear();
            credential.RevokeTokenAsync(new CancellationToken());
        }

        //---------------------------------------------------------------------------
        //
        // mouse Click - logic
        //   if in Cloumn 0 - change to row select mode
        //          does not allow row deletes
        //   if in any other column - change to cell select mode
        //
        private void DataGridView1_Cell_Mouse_Click(object sender, DataGridViewCellEventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                return;
            }

            int intIndex = e.RowIndex;
            //int i = 0;
            dataGridView1.MultiSelect = false;

            while (dr == DialogResult.OK || dr == DialogResult.None || dr == DialogResult.No || dr == DialogResult.Retry)
            {
                done = false;
                textBox2.ForeColor = System.Drawing.Color.Black;

                if (dataGridView1.CurrentCell.ColumnIndex == 0)
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    dataGridView1.EditMode = DataGridViewEditMode.EditOnF2;

                    // Sheet Tab selection
                    int selectIndex = comboBox1.SelectedIndex;
                    sTabName selectCar = (sTabName)comboBox1.SelectedItem;
                    String range = selectCar.tabname + selectCar.range;

                    string cellResult = null;
                    string cellNote = null;

                    int nRow = dataGridView1.CurrentCell.RowIndex;
                    if (selectIndex == 1)
                    {
                        if (dr == DialogResult.OK || dr == DialogResult.No)
                        {
                            if (nRow < dataGridView1.RowCount)
                                intIndex = nRow;

                        }
                        if (dr == DialogResult.Retry)
                        {
                            intIndex = intIndex - 1;
                        }
                        else
                        {
                            if (intIndex + 1 >= dataGridView1.RowCount)
                            {
                                textBox2.ForeColor = System.Drawing.Color.Red;
                                textBox2.Text = "last row reached";
                                textBox2.Update();
                                break;
                            }
                        }
                        if (selectIndex == 1)           // is Field Service selected
                        {
                            using (var Form2 = new Form2())
                            {
                                Form2 form2 = new Form2();
                                form2.StartPosition = FormStartPosition.Manual;
                                form2.Left = SecondFormLeft;
                                form2.Top = SecondFromTop;
                                if (SecondFormSize == System.Drawing.Size.Empty)
                                    SecondFormSize = form2.Size;
                                form2.Size = SecondFormSize;

                                short b = 0;
                                string testStr = "";

                                DateTime today = DateTime.Today;

                                form2.textBox1.Text = dataGridView1.CurrentRow.Cells[0].Value?.ToString();
                                form2.textBox2.Text = dataGridView1.CurrentRow.Cells[1].Value?.ToString();
                                form2.textBox3.Text = dataGridView1.CurrentRow.Cells[2].Value?.ToString();
                                form2.textBox4.Text = dataGridView1.CurrentRow.Cells[5].Value?.ToString();
                                string temp1 = textBox3.Text + " of " + textBox5.Text;
                                form2.textBox8.Text = temp1;
                                form2.textBox9.Text = textBox2.Text;

                                form2.textBox7.Text = dataGridView1.CurrentRow.Cells[3].Value?.ToString();
                                string tmp = "";
                                string tmp0 = "";
                                if (string.IsNullOrEmpty(form2.textBox1.Text)) tmp = " no phone ";
                                else tmp = form2.textBox1.Text;
                                if (string.IsNullOrEmpty(form2.textBox7.Text)) tmp0 = " ";
                                else tmp0 = " - city " + form2.textBox7.Text;
                                Clipboard.SetText(tmp + tmp0 + Environment.NewLine);

                                testStr = "0";
                                if (!Int16.TryParse(dataGridView1.CurrentRow.Cells[6].Value?.ToString(), out b)) form2.textBox5.Text = testStr;
                                else
                                {
                                    b++;
                                    form2.textBox5.Text = string.Format("{0}", b);
                                }

                                testStr = dataGridView1.CurrentRow.Cells[7].Value?.ToString();
                                if (testStr == null) form2.textBox6.Text = " ";
                                else form2.textBox6.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                                bool firstpos = false;

                                string temp2 = form2.textBox6.Text; ;
                                string text = temp2.Replace(Environment.NewLine, "^"); // a random token
                                string[] lines = text.Split('^');
                                for (int j = 0; j < lines.Count(); j++)
                                {
                                    string temp = form2.textBox6.Lines[j];
                                    firstpos = temp.Contains(today.ToString("yyyy-MM-dd"));
                                    if (firstpos) break;
                                }

                                if (!firstpos) form2.textBox6.Text = form2.textBox6.Text + Environment.NewLine + today.ToString("yyyy-MM-dd ddd tt") + ": ";
                                else form2.textBox6.Text = form2.textBox6.Text + Environment.NewLine;
                                switch (dataGridView1.CurrentRow.Cells[4].Value?.ToString())
                                {
                                    case "N/A":
                                        form2.radioButton1.Checked = true;
                                        break;
                                    case "B":
                                        form2.radioButton2.Checked = true;
                                        break;
                                    case "I":
                                        form2.radioButton3.Checked = true;
                                        break;
                                    case "DNC":
                                        form2.radioButton4.Checked = true;
                                        break;
                                    case "SP":
                                        form2.radioButton5.Checked = true;
                                        break;
                                    case "E":
                                        form2.radioButton6.Checked = true;
                                        break;
                                    default:
                                        break;
                                }

                                form2.textBox6.SelectionStart = form2.textBox6.Text.Length + 4;
                                form2.textBox6.SelectionLength = 0;

                                dr = form2.ShowDialog();        // bring up the form

                                SecondFormLeft = form2.Left;    // save current position for next time
                                SecondFromTop = form2.Top;
                                SecondFormSize = form2.Size;

                                if (dr == DialogResult.OK)
                                {
                                    //
                                    // process RadioButtons on Form2
                                    var checkedRadio = new[] { form2.groupBox1 }.SelectMany(g => g.Controls.OfType<RadioButton>()
                                                        .Where(r => r.Checked));
                                    // Print name
                                    foreach (var c in checkedRadio)
                                    {
                                        //System.Diagnostics.Debug.Print(c.Name);
                                        switch (c.Text)
                                        {
                                            case "No Answer":
                                                cellResult = "N/A";
                                                break;

                                            case "Business":
                                                cellResult = "B";
                                                break;

                                            case "Inoperative":
                                                cellResult = "I";
                                                break;

                                            case "Do Not Call":
                                                cellResult = "DNC";
                                                break;

                                            case "Spanish":
                                                cellResult = "SP";
                                                break;

                                            case "Native English":
                                                cellResult = "E";
                                                break;

                                            default:
                                                break;
                                        }
                                        cellNote = form2.textBox6.Text;
                                        Console.WriteLine(cellResult);
                                        Console.WriteLine(cellNote);
                                        myVar = null;
                                    }
                                }


                                if (dr == DialogResult.OK)
                                {
                                    //DateTime today = DateTime.Today
                                    dataGridView1.CurrentRow.Cells[5].Value = today.ToString("yyyy-MM-dd");
                                    // bump up attemps

                                    string numberSt = dataGridView1.CurrentRow.Cells[6].Value?.ToString();
                                    if (string.IsNullOrEmpty(numberSt)) numberSt = "0";
                                    int number = -1;
                                    number = Convert.ToInt32(numberSt);
                                    ++number;
                                    dataGridView1.CurrentRow.Cells[6].Value = number.ToString();

                                    if (!string.IsNullOrEmpty(cellResult)) dataGridView1.CurrentRow.Cells[4].Value = cellResult.ToString();    // Result

                                    if (!string.IsNullOrEmpty(cellNote)) cellNote = cellNote + " ";
                                    else cellNote = " ";
                                    dataGridView1.CurrentRow.Cells[7].Value = cellNote.ToString();      // Notes

                                    //
                                    // unselect current row and find the next visable
                                    //
                                    dataGridView1.Rows[nRow].Selected = false;
                                    while (nRow < dataGridView1.RowCount - 1)
                                        if (nRow < dataGridView1.RowCount)
                                        {
                                            if (dataGridView1.Rows[++nRow].Visible == false)
                                                continue;
                                            dataGridView1.Rows[nRow].Selected = true;
                                            int currentRow = nRow;
                                            dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                            break;
                                        }
                                }

                                //
                                // skip to next visable row processing
                                //
                                if (dr == DialogResult.No)
                                {
                                    dataGridView1.Rows[nRow].Selected = false;
                                    while (nRow < dataGridView1.RowCount - 1)
                                    //if (nRow < dataGridView1.RowCount)
                                    {
                                        if (dataGridView1.Rows[++nRow].Visible == false)
                                            continue;
                                        dataGridView1.Rows[nRow].Selected = true;
                                        int currentRow = nRow;
                                        dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                        break;
                                    }
                                    dataGridView1.Refresh();
                                    continue;
                                }

                                //
                                // backup to previous visable row 
                                //
                                if (dr == DialogResult.Retry)
                                {
                                    dataGridView1.Rows[nRow].Selected = false;
                                    while (nRow > 0)
                                    {
                                        if (dataGridView1.Rows[--nRow].Visible == false)
                                            continue;
                                        dataGridView1.Rows[nRow].Selected = true;
                                        int currentRow = nRow;
                                        dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                        break;
                                    }
                                    if (nRow <= 0)
                                    {
                                        textBox2.ForeColor = System.Drawing.Color.Red;
                                        textBox2.Update();
                                        textBox2.Text = " top row reached ";
                                    }
                                    dataGridView1.Refresh();
                                    continue;
                                }
                            }
                        }
                        else break;
                    }
                    else
                    {
                        dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
                        dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
                    }
                    if (selectIndex == 1)
                    {
                        if (!dataGridView1.Rows[nRow].Visible) continue;
                    }
                    else return;
                    //dataGridView1.CurrentCell = dataGridView1[0, nRow];
                    //dataGridView1.Refresh();
                }
                else break;
            }
            dr = DialogResult.None;
            done = true;
        }

        //
        // detect click in header row
        //
        private void dataGridView1_ColumnHeaderMouseClick(
    object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dataGridView1.Columns[e.ColumnIndex];
            //DataGridViewColumn oldColumn = dataGridView1.SortedColumn;
            //ListSortDirection direction;
        }


        //---------------------------------------------------------------------------
        //
        // Cell Changes are tracked  and exit is prevented until saved
        //
        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dr != DialogResult.OK || done) return;
            if (e.ColumnIndex != 0)
                DataGridView1_CellEndEdit(sender, e);
        }

        //---------------------------------------------------------------------------
        //
        // row removal is only permitted in automation
        //
        private void DataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

            UpdateSheet();
        }

        //---------------------------------------------------------------------------
        //
        // user row deletion is not permitted
        //
        private void DataGridView1_UserDeletingRow(object sender,
    DataGridViewRowCancelEventArgs e)
        {
            MessageBox.Show("Cannot delete a record!");
            e.Cancel = true;
        }

        //---------------------------------------------------------------------------
        //
        // UpdateSheet - writes back all changed to the Sheet by calling SaveSheetChanges
        //
        private void UpdateSheet()
        {
            if (!updateinprogress && checkBox1.Checked)
            {
                //Console.WriteLine("UpdateSheet called");
                Cursor.Current = Cursors.WaitCursor;
                SaveSheetChanges();
                Cursor.Current = Cursors.Default;
            }
            else
            {
                if (!updateinprogress) dataGridView1.Refresh();
            }
        }




        //---------------------------------------------------------------------------
        //
        // SaveSheetChanges as found in celch array
        //
        DateTime begTimInc = DateTime.MinValue;
        DateTime nowTimInc = DateTime.MinValue;
        int maxChgs = 20;
        int perSec = 20;
        int totalChgs = 0;
        private void SaveSheetChanges()
        {

            string[] Scopes = { SheetsService.Scope.Spreadsheets };

            startPB(System.Drawing.Color.Yellow);                          // startPB a new bar
            //UserCredential credential;
            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            progressBar1.Value = 2;
            progressBar1.Update();
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            progressBar1.Value = 4;
            progressBar1.Update();

            string sCol = null;
            string sRow = null;
            string sValue = null;
            int l = 0;
            //
            // loop through all changes and send update
            foreach (var ListItem in cellch)
            {
                if (!checkBox1.Checked)
                {
                    nowTimInc = DateTime.Now;
                    if (begTimInc == DateTime.MinValue) begTimInc = nowTimInc;
                    TimeSpan duration = nowTimInc - begTimInc;
                    int d = (int)duration.TotalSeconds;

                    if (d >= perSec || totalChgs >= maxChgs)
                    {
                        //for(int z=0; z < perSec; z++)
                        for (int z = 0; z < (totalChgs > d ? totalChgs : d); z++)
                        {
                            if (checkBox1.Checked && d > perSec)
                            {
                                d = 0;
                                break;
                            }
                            textBox2.ForeColor = System.Drawing.Color.Red;
                            textBox2.Text = string.Format(" waiting {0} sec.", (int)duration.TotalSeconds - z);
                            textBox2.Update();
                            Thread.Sleep(1000);
                        }

                        begTimInc = DateTime.MinValue;
                        totalChgs = 0;
                        textBox2.ForeColor = System.Drawing.Color.Black;
                        textBox2.Text = " continuing";
                        textBox2.Update();
                    }
                }

                foreach (string SubListItem in ListItem)
                {
                    if (sValue == null)
                    {
                        sCol = Regex.Match(SubListItem, @"\d+").Value;
                        sRow = Regex.Match(SubListItem, @"(\d+)(?!.*\d)").Value;
                        sValue = "data";
                        continue;
                    }
                    sValue = SubListItem.ToString();

                    if (!String.IsNullOrEmpty(sRow))
                    {
                        int number = Convert.ToInt32(sRow);
                        number = number + rowOffset + 1;
                        sRow = number.ToString();
                    }
                    if (!String.IsNullOrEmpty(sCol))
                        sCol = Program.ColumnAdress(Int32.Parse(sCol));
                    Console.WriteLine(sCol + ":" + sRow);
                    textBox2.Text = sRow;
                    //textBox2.Update();

                    String spreadsheetId2 = spreadsheetId; ;
                    int selectIndex = comboBox1.SelectedIndex;
                    sTabName selectCar = (sTabName)comboBox1.SelectedItem;
                    //
                    //String range = selectCar.tabname + selectCar.range;
                    String range2 = selectCar.tabname + "!" + sCol + sRow + ":" + sCol + sRow;  // cell to update on Tab 
                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "COLUMNS";// "ROWS";//COLUMNS

                    var oblist = new List<object>() { sValue };
                    valueRange.Values = new List<IList<object>> { oblist };

                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result2 = update.Execute();
                    //totalChgs++;
                    sValue = null;
                }
                int remainder;
                l++;
                Math.DivRem(l, cellch.Count, out remainder);
                if (cellch.Count > 100)
                    progress = (remainder / (cellch.Count / 100));
                else
                    progress = remainder * ((100 - 1) / cellch.Count);
                if (progress < 100) progressBar1.Value = progress;
                progressBar1.Update();
            }
            //if(!checkBox1.Checked)
            //{
            //    begTimInc = DateTime.MinValue;
            //    totalChgs = 1;
            //}
            Console.WriteLine("done!");
            DataChanged = false;
            cellch.Clear();
            textBox1.Text = " done ";
            progressBar1.Value = 100;
            return;
        }

        //
        //
        //
        private void button4_Click(object sender, EventArgs e)
        {
            SaveSheetChanges();
            cellch.Clear();
        }

        private void startPB(System.Drawing.Color color)
        {

            if (isProcessRunning) return;
            isProcessRunning = true;

            progressBar1.Name = "progressBar1";
            progressBar1.TabIndex = 0;
            progressBar1.Maximum = 100;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            progressBar1.ForeColor = color;
        }



        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (updateinprogress) return;
            int? rowIdx = e?.RowIndex;
            int? colIdx = e?.ColumnIndex;
            if (rowIdx.HasValue && colIdx.HasValue)
            {
                var dgv = (DataGridView)sender;
                var cell = dgv?.Rows?[rowIdx.Value]?.Cells?[colIdx.Value]?.Value;
                if (!string.IsNullOrEmpty(cell?.ToString()) && e.ColumnIndex != 0)
                {
                    DataGridViewCell oldcurcell = dgv.CurrentCell;
                    DataChanged = true;
                    int i = cellch.Count;
                    dgv.CurrentCell = dataGridView1[colIdx.Value, rowIdx.Value];

                    cellch.Add(new List<String>()); //Adds new sub List
                    cellch[i].Add(dgv.CurrentCellAddress.ToString()); //Add values to the sub List at index 0
                    cellch[i].Add(cell.ToString());
                    foreach (var ListItem in cellch)
                    {
                        foreach (string SubListItem in ListItem)
                        {
                            if (!Program.Check(SubListItem)) continue;
                            string sCol = Regex.Match(SubListItem, @"\d+").Value;
                            string sRow = Regex.Match(SubListItem, @"(\d+)(?!.*\d)").Value;

                            if (!String.IsNullOrEmpty(sRow))
                            {
                                int number = Convert.ToInt32(sRow);
                                number++;
                                sRow = number.ToString();
                            }
                            if (!String.IsNullOrEmpty(sCol))
                                sCol = Program.ColumnAdress(Int32.Parse(sCol));
                            Console.WriteLine(sCol + ":" + sRow);
                        }
                    }
                    dgv.CurrentCell = oldcurcell;
                    Console.WriteLine("----------------");
                    textBox1.Text = cellch.Count.ToString() + " changes ";
                    textBox1.Update();
                    if (checkBox1.Checked) SaveSheetChanges();
                }
                UpdateSheet();
            }
            //string title = matrix[0][1]; //Retrieve value at index 1 from sub List at index 0
        }

        private void dataGridView1_SelectionChanged(object sender, DataGridViewRowEventArgs e)
        {
            textBox2.Text = e.Row.ToString();
            textBox2.Update();
        }
        private void Textbox_DoulbClickEvent(object sender, EventArgs e)
        {
            string changeSummary = null;
            int i = 0;
            foreach (var ListItem in cellch)
            {
                if (i++ > 100) break;
                foreach (string SubListItem in ListItem)
                {
                    changeSummary = changeSummary + SubListItem + " ";
                    changeSummary = changeSummary + ("\n ");
                }
            }
            changeSummary = changeSummary + "----------------";
            MyMessageBox.Show(changeSummary);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            waiting = false;
        }

        //
        // sheet tab selection
        //
        private void combox1_selectedIndexChange(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();
            if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 2)
            {
                string[] checklist = new string[] { "N/A", "B", "I", "DNC", "SP", "E", "blank", "pE" };
                for (int i = 0; i < checklist.Length; i++)
                {
                    checkedListBox1.Items.Add(checklist[i]);
                }
                checkedListBox1.Visible = true;
                textBox3.Visible = true;
                textBox5.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                label1.Visible = true;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                button8.Visible = false;
                textBox6.Visible = false;
                textBox7.Visible = false;
                textBox8.Visible = false;
                button9.Visible = false;
                checkedListBox1.CheckOnClick = true;
            }
            if (comboBox1.SelectedIndex == 0)
            {
                checkedListBox1.Visible = true;
                textBox3.Visible = true;
                textBox5.Visible = true;
                label5.Visible = true;
                label4.Visible = true;
                button9.Visible = false;
                checkedListBox1.CheckOnClick = true;
            }

            if (comboBox1.SelectedIndex == 2)
            {
                button9.Visible = false;
            }

            if (comboBox1.SelectedIndex == 3)
            {
                button9.Visible = false;
            }

            if (comboBox1.SelectedIndex == 4)
            {
                button9.Visible = false;
            }

            if (comboBox1.SelectedIndex == 5)
            {
                button9.Visible = false;
            }


        }

        //
        // update current row textBox
        int curRow = -1;
        int curRowUpdate = -1;
        private Control txtConsole;
        private readonly int visibleColumnIndex;

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (!updateinprogress)
            {
                textBox2.ForeColor = System.Drawing.Color.Black;
                textBox2.Text = dataGridView1.CurrentRow.Index.ToString();
                textBox2.Update();
            }
        }

        //
        // search processing
        //
        private void Program_SearchDataGrid(object sender, EventArgs e)
        {
            int rowIndex = -1;
            string searchString = textBox4.Text;
            if (searchString == "") return;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value.ToString().Equals(searchString))
                {
                    if (row.Cells[0].Visible == false) continue;
                    rowIndex = row.Index;
                    break;
                }
            }
            if (rowIndex >= 0)
            {
                dataGridView1.CurrentCell = dataGridView1[visibleColumnIndex, rowIndex];
            }
            else return;
            //
            //
            //

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //
        // search datagrid for value
        private void button6_Click_1(object sender, EventArgs e)
        {
            string searchValue = "";
            searchValue = textBox4.Text;
            int foundCnt = 0;

            dataGridView1.ClearSelection();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;

            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Visible == false) continue;
                    for (int y = 0; y < rcount; y++)
                    {
                        if (string.IsNullOrEmpty(row.Cells[y].Value?.ToString())) continue;
                        if ((row.Cells[y].Value?.ToString().ToUpper()).ToUpper().Contains(searchValue))
                        {
                            dataGridView1.Rows[row.Index].Selected = true;
                            foundCnt++;
                            break;
                        }
                    }
                    continue;
                }
                MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            dataGridView1.ClearSelection();
        }

        //
        //
        // RunImported Names
        //

        List<List<String>> names2chk = new List<List<String>>();
        private void RunImportedNames(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            using (var UserControl1 = new UserControl1())
            {

                cellch.Clear();
                textBox1.Text = " ..logging in ";
                textBox1.Update();

                // If modifying these scopes, delete your previously saved credentials
                // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
                string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                string ApplicationName = "Google Sheets API .NET Quickstart";

                //UserCredential credential;

                Cursor.Current = Cursors.WaitCursor;

                using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }
                textBox1.Text = "..connecting Sheet ";
                textBox1.Update();
                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define request parameters.
                //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
                //spreadsheetId = "1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI";
                string spreadsheetId1 = "1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok";    // test sheet
                string spreadsheetId2 = "1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc";    // live sheet  /edit#gid=2145664999"

                //
                // process RadioButtons on Form1
                //var checkedRadio = new[] { groupBox1 }.SelectMany(g => g.Controls.OfType<RadioButton>()
                //                    .Where(r => r.Checked));
                //*if (form1.radioButton1.Checked) */
                if (radioButton2.Checked) spreadsheetId = spreadsheetId1;
                if (radioButton2.Checked) spreadsheetId = spreadsheetId2;

                //int selectIndex = form1.comboBox1.SelectedIndex;
                //selectCar = "List of Names";
                String range = "Common First Names!A1:A";
                //String range = selectCar.tabname + selectCar.range;

                //UserControl..ActiveForm.Text = String.Format("Working with sheet '{0}', columns {1} ", spreadsheetId, range);
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

                // Prints the names and majors of students in a sample spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                // https://docs.google.com/spreadsheets/d/1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI/edit#gid=1072334158
                // https://docs.google.com/spreadsheets/d/1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok/edit#gid=2145664999
                // https://docs.google.com/spreadsheets/d/1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc/edit#gid=2145664999
                ValueRange response = request.Execute();
                textBox1.Text = ".. reading data";
                textBox1.Update();
                names2chk.Clear();
                DataChanged = true; ;
                IList<IList<Object>> values = response.Values;
                int i;
                foreach (var row in values)
                {
                    i = names2chk.Count;
                    //if (string.IsNullOrWhiteSpace(values[][0].ToString()))
                    //    continue;

                    names2chk.Add(new List<String>()); //Adds new sub List
                    names2chk[i].Add(values[i][0].ToString()); //Add values to the sub List at index 0
                    names2chk[i].Add(" ".ToString());
                }
                dataGridView1.Visible = false;
                int numOfSP = 0;
                int numOfEN = 0;
                int numOfSkip = 0;

                int nRow = 0;
            LB1: while (nRow < dataGridView1.RowCount - 1)
                {
                    if (dataGridView1.Rows[++nRow].Visible == false)
                        continue;

                    dataGridView1.Rows[nRow].Selected = true;
                    int currentRow = nRow;
                    dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                    //dataGridView1.CurrentCell.ro = nRow;
                    dataGridView1.CurrentRow.Selected = true;
                    break;
                }
                if (nRow >= dataGridView1.RowCount - 1)
                {
                    updateinprogress = false;
                    dataGridView1.Visible = true;
                    return;
                }
                int l = 0;
                //int J = names2chk.Count;
                bool found = false;
                foreach (var row in values)
                {
                    string names = dataGridView1.CurrentRow.Cells[1].Value?.ToString();
                    string phone = dataGridView1.CurrentRow.Cells[0].Value?.ToString();
                    if (string.IsNullOrEmpty(names) || string.IsNullOrEmpty(phone))
                    {
                        numOfSkip++;
                        goto LB1;
                    }
                    string[] nameList = names.Split(' ');
                    found = false;
                    foreach (string name in nameList)
                    {
                        for (l = 0; l < names2chk.Count; l++)
                        {
                            if (names2chk[l][0].ToString().ToUpper() == name.ToUpper())
                            {
                                found = true;
                                break;
                            }
                            if (found) break;
                        }

                    }
                    if (found)
                    {
                        dataGridView1.CurrentRow.Cells[4].Selected = true;
                        dataGridView1.CurrentRow.Cells[4].Value = "pE";
                        numOfEN++;
                        textBox6.Text = string.Format("found {0}", numOfEN);
                        textBox6.Update();
                        int c = cellch.Count;
                        cellch.Add(new List<String>());
                        cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[c].Add("pE");

                        dataGridView1.CurrentRow.Cells[5].Value = today.ToString("yyyy-MM-dd");
                        dataGridView1.CurrentRow.Cells[5].Selected = true;
                        c = cellch.Count;
                        cellch.Add(new List<String>());
                        cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[c].Add(today.ToString("yyyy-MM-dd"));
                    }
                    else
                    {
                        dataGridView1.CurrentRow.Cells[4].Selected = true;
                        dataGridView1.CurrentRow.Cells[4].Value = "SP";
                        numOfSP++;
                        textBox7.Text = string.Format("found {0}", numOfSP);
                        textBox7.Update();
                        int c = cellch.Count;
                        cellch.Add(new List<String>());
                        cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[c].Add("SP");

                        dataGridView1.CurrentRow.Cells[5].Value = today.ToString("yyyy-MM-dd");
                        dataGridView1.CurrentRow.Cells[5].Selected = true;
                        c = cellch.Count;
                        cellch.Add(new List<String>());
                        cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[c].Add(today.ToString("yyyy-MM-dd"));
                    }
                    textBox8.Text = string.Format("skipped {0}", numOfSkip);
                    textBox8.Update();
                    textBox1.Text = cellch.Count.ToString() + " changes";
                    textBox1.Update();
                    if (checkBox1.Checked) SaveSheetChanges();
                    goto LB1;
                }
            }
            updateinprogress = false;
            dataGridView1.Visible = true;
            return;
        }
        private void copyrange(object sender, EventArgs e/*string sht1, string rng1, string ht2*/)
        {
            //var ss = SpreadsheetApp.getActiveSpreadsheet();
            //var sheet = ss.getSheetByName('Sheet1'); //source sheet
            //var testrange = sheet.getRange('H:H'); //range to check
            //var testvalue = (testrange.getValues());
            //var csh = ss.getSheetByName('Sheet2'); //destination sheet
            //var data = [];
            //var j =[];

            DateTime today = DateTime.Today;
            using (var UserControl1 = new UserControl1())
            {

                cellch.Clear();
                textBox1.Text = " ..logging in ";
                textBox1.Update();

                // If modifying these scopes, delete your previously saved credentials
                // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
                string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                string ApplicationName = "Google Sheets API .NET Quickstart";

                //UserCredential credential;

                Cursor.Current = Cursors.WaitCursor;

                using (var stream =
                        new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }
                textBox1.Text = "..connecting Sheet ";
                textBox1.Update();

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define request parameters.
                //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
                //spreadsheetId = "1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI";
                string spreadsheetId1 = "1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok";    // test sheet
                string spreadsheetId2 = "1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc";    // live sheet  /edit#gid=2145664999"

                //
                // process RadioButtons on Form1
                //var checkedRadio = new[] { groupBox1 }.SelectMany(g => g.Controls.OfType<RadioButton>()
                //                    .Where(r => r.Checked));
                //*if (form1.radioButton1.Checked) */
                if (radioButton1.Checked) spreadsheetId = spreadsheetId1;
                if (radioButton2.Checked) spreadsheetId = spreadsheetId2;

                //-----------------------------------------------------------
                //
                // get phone numbers in Only Spanish
                //
                string range = "Only Spanish!A5:A";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

                // get the contects of selected spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                // https://docs.google.com/spreadsheets/d/1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI/edit#gid=1072334158
                // https://docs.google.com/spreadsheets/d/1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok/edit#gid=2145664999
                // https://docs.google.com/spreadsheets/d/1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc/edit#gid=2145664999
                ValueRange response = request.Execute();
                textBox1.Text = ".. reading data";
                textBox1.Update();

                names2chk.Clear();              // clear phone array
                DataChanged = true; ;

                IList<IList<Object>> values = response.Values; // list of phone numbers and source

                int i;
                int j = 0;
                string forchk = "";
                foreach (var row in values)
                {
                    i = names2chk.Count;

                    names2chk.Add(new List<String>());      //Adds new sub List
                    forchk = values[i][0].ToString();
                    names2chk[i].Add(forchk);               //Add values to the sub List at index 0
                    names2chk[i].Add("Only Spanish".ToString());
                    if (names2chk.Distinct().Count() != names2chk.Count())
                    {
                        MessageBox.Show("duplicate {0}", forchk.ToString());
                    }
                }

                //------------------------------------------------------------
                //
                // get phone numbers in Files Service
                //
                range = "Field Service!A5:A";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                textBox1.Text = ".. reading data";
                textBox1.Update();

                forchk = "";
                foreach (var row in values)
                {
                    i = names2chk.Count;

                    if (values[j++].Count == 0) continue;
                    else forchk = values[j - 1][0].ToString();
                    names2chk.Add(new List<String>());
                    names2chk[i].Add(forchk); //Add values to the sub List at index 0
                    names2chk[i].Add("Field Service".ToString());
                    if (names2chk.Distinct().Count() != names2chk.Count())
                    {
                        MessageBox.Show("duplicate {0}", forchk.ToString());
                    }

                }
                dataGridView1.Visible = false;

                int numOfSP = 0;
                int numOfEN = 0;
                int numOfSkip = 0;
                int nRow = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)  // find first viable row
                {
                    if (row.Cells[0].Visible == false) continue;
                    nRow = row.Index;
                    break;
                }

                List<string> svalue = new List<string>();
                sTabName selectFrom = (sTabName)comboBox1.SelectedItem;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Visible == false) continue;
                    string tmp = row.Cells[row.Index].Value?.ToString();
                    if (tmp != "pE") continue;
                    else row.Cells[4].Value = "E";
                    if (string.IsNullOrEmpty(tmp)) row.Cells[row.Index].Value = " ";
                    dataGridView1.Update();
                    for (int w = 0; w < row.Cells.Count; w++)
                    {
                        if (string.IsNullOrEmpty(row.Cells[w].Value?.ToString())) svalue.Add(" ");
                        else if (w != 4) svalue.Add(row.Cells[w].Value?.ToString());
                        else svalue.Add("E");
                    }

                    string[] Scope = { SheetsService.Scope.Spreadsheets };

                    startPB(System.Drawing.Color.Yellow);                          // startPB a new bar
                                                                                   //UserCredential credential;
                    using (var stream =
                        new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                    {
                        string credPath = System.Environment.GetFolderPath(
                            System.Environment.SpecialFolder.Personal);

                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scope,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }
                    progressBar1.Value = 2;
                    progressBar1.Update();
                    // Create Google Sheets API service.
                    service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    int selectIndex = comboBox1.SelectedIndex;          // select it
                    spreadsheetId2 = spreadsheetId;
                    //int selectIndex = comboBox1.SelectedIndex;
                    //sTabName selectCar = (sTabName)comboBox1.SelectedItem;
                    //
                    //String range = selectCar.tabname + selectCar.range;
                    String range2 = "Field Service!A:H";  // cell to update on Tab 
                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = /*"COLUMNS";*/ "ROWS";//COLUMNS

                    var oblist = new List<object>() { svalue };
                    valueRange.Values = new List<IList<object>> { oblist };
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result2 = update.Execute();
                    //totalChgs++;
                    svalue = null;
                }

            }
        }
    }
}
          



