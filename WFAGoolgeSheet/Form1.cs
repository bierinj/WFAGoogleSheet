using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;  
using System.IO;
using System.Linq;
using System.Globalization;
using System.Threading;
using System.Resources;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows;
using System.ComponentModel;

namespace WFAGoolgeSheet
{
    public partial class Form1 : Form
    {
        public string language = Properties.Settings.Default.Language;
        public DataGridView dataGridView { get; set; }
        //private  FormMediator _formMediator;
        public IList<ResourceDictionary> MergedDictionaries { get; }
        public System.Uri Source { get; set; }
        public class sTabName
        {

            public string tabname { get; set; }
            public string range { get; set; }
            public string keyfield { get; set; }
            public override string ToString() { return this.tabname; }
        }

        public Form1()
        {
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(language);
            InitializeComponent();

            comboBox1.Items.Add(new sTabName { tabname = "Imported Names", range = "!A1:K", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "Field Service", range = "!A1:K", keyfield = "Field Service" });
            comboBox1.Items.Add(new sTabName { tabname = "Confirmed English", range = "!A1:K", keyfield = "Telephone" });
            comboBox1.Items.Add(new sTabName { tabname = "Contacted 5 times letters", range = "!A1:K", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "Only Spanish", range = "!A1:K", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "Other", range = "!A1:K", keyfield = "TELEPHONE" });
            comboBox1.Items.Add(new sTabName { tabname = "List of Last names to work", range = "!A1:G", keyfield = "Name" });
            //comboBox1.Items.Add(new sTabName { tabname = "Common First Names", range = "!A1:G", keyfield = "TELEPHONE" });
            comboBox1.SelectedIndex = 1;
        }

        public DataGridView myDG { get { return dataGridView1; } }

        public object Resources { get; private set; }

        //---------------------------------------------------------------------------//
        // global flags and data
        //
        bool updateinprogress = false;
        bool DataChanged = false;
        bool isProcessRunning = false;
        //int dataLoadForSheet = -1;
        //bool waiting = false;
        //bool isCancelled = false;
        string Tabfocus = null;
        bool gEODhit = false;
        int rowOffset = 0;
        int firstrow = 0;
        int skiprow = 0;
        int lastFSrow = 0;
        //int lastINrow = 0;
        int lastSProw = 0;
        int lastCErow = 0;
        int lastC5row = 0;
        int progress = 0;
        int rcount = -1;
        //int oldIndex = 0;
        int chgCount = 0;
        int foundCnt = 0;
        int firstFound = 0;
        int curFound = 0;
        public static string myVar = null;  // passing data between forms
        public static object lockMethod2 = new object();
        public DialogResult dr = new DialogResult();
        List<String> sSaveRow4Del = new List<String>();
        public float notfoundlat = float.Parse(" -1.4487420320510864");
        public float notfoundlon= float.Parse("-78.376708984375");
        bool done = true;
        int totalRow = 0;
        int totalChgs = 0;
        int numOfSP = 0;
        int numOfpE = 0;
        int numOfSkip = 0;
        public bool setMinDate = false;
        public bool _setPM = false;
        public bool _setAM = false;
        public string checkedRadio = "";
        public string resul = "";
        public string attempt = "";
        public string notes = "";
        public bool adjGPS = false;
        public string confid = "";
        public string lat = "";
        public string lon = "";
        DateTime lastDate = new DateTime();
        DateTime pastTime = new DateTime();
        List<string> lst = new List<string>();    // for days of the week
        public int SecondFromTop;
        public int SecondFormLeft;
        public System.Drawing.Size SecondFormSize = System.Drawing.Size.Empty;
        public int MapFormTop;
        public int MapFormLeft;
        public System.Drawing.Size MapFormSize;
        string[] vs = null;
        //
        // list for sheets row moves
        //
        List<List<String>> moves = new List<List<String>>();

        //
        // list of all datagridview changes
        //
        public List<List<String>> cellch = new List<List<String>>(); //Creates new nested List

        //
        // global phone number list
        //
        IList<IList<object>> values = new List<IList<object>>();

        //
        // Google parameters
        public UserCredential credential;               // Sheet Service
        public string ApplicationName = null;
        public String spreadsheetId = null;             // Spreadsheet ID

        //
        // Map form create
        //
        //public void SomeMethodThatOpensTheSubForm()
        //{
        //    Form4 subForm = new Form4();

        //    _formMediator = new FormMediator(this, subForm);

        //    subForm.Show(this);
        //}
        //---------------------------------------------------------------------------
        //
        // close application and exit
        //
        private void button1_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous 3 dialogs.
                //
                if (result1 == DialogResult.No || result1 == DialogResult.Retry)
                    return;
            }
            System.Windows.Forms.Application.Exit();
        }

        //---------------------------------------------------------------------------
        //
        // run application 
        //
        private void button2_Click(object sender, EventArgs e)
        {
            //isCancelled = false;

            int NumofRec = 0;
            int r1 = -1;
            if (string.IsNullOrEmpty(textBox3.Text)) firstrow = 0;  // setup x of n grouping
            else firstrow = Convert.ToInt32(textBox3.Text);
            if (string.IsNullOrEmpty(textBox5.Text)) skiprow = 0;
            else skiprow = Convert.ToInt32(textBox5.Text);
            firstrow = skiprow + 1 - firstrow;                       // invert it!

            if (checkedListBox1.CheckedItems.Count < 1 && comboBox1.SelectedIndex < 2)
            {
                System.Windows.Forms.MessageBox.Show(" You have no Filters selected");
                return;
            }
            if (comboBox1.SelectedIndex == -1)
            {
                System.Windows.Forms.MessageBox.Show(" No Sheet Tab selected");
                return;
            }

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                System.Windows.Forms.MessageBox.Show(" No Sheet [test/live] selected");
                return;
            }
            if (comboBox1.SelectedIndex == 1)
            {
                if (firstrow == 0 || skiprow == 0 || firstrow > skiprow)
                {
                    System.Windows.Forms.MessageBox.Show(" group range is bad ");
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
                    System.Windows.Forms.MessageBox.Show(" no filters selected ");
                    return;
                }
            }
            if (DataChanged)
            {
                DialogResult result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to Re-Run and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous dialog.
                //
                if (result1 == DialogResult.No || result1 == DialogResult.Retry)
                    return;
            }

            button2.BackColor = System.Drawing.Color.Coral;
            button2.Update();
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
            dataGridView1.GridColor = System.Drawing.Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.Navy;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font =
                new Font(dataGridView1.Font, System.Drawing.FontStyle.Bold);
            dataGridView1.Name = "dataGridView1";

            dataGridView1.AutoSizeRowsMode =
                DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.AllowUserToResizeColumns = true;
            dataGridView1.AllowUserToResizeRows = true;

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
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
            string spreadsheetId1 = Properties.Settings.Default.TestSheet;    // test sheet
            string spreadsheetId2 = Properties.Settings.Default.ProdSheet;    // live sheet  /edit#gid=2145664999"

            if (radioButton1.Checked) spreadsheetId = spreadsheetId1;
            if (radioButton2.Checked) spreadsheetId = spreadsheetId2;

            int selectIndex = comboBox1.SelectedIndex;
            sTabName selectCar = (sTabName)comboBox1.SelectedItem;
            String range = selectCar.tabname + selectCar.range;

            Form1.ActiveForm.Text = String.Format("Working with sheet '{0}', columns {1} ", spreadsheetId, range);

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);
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
                //
                // find column 0 backcolors
                //
                List<string> bcolor = new List<string>();
                SpreadsheetsResource.GetRequest request2 = service.Spreadsheets.Get(spreadsheetId);
                request2.Ranges = string.Format("{0}!A{1}:A{2}", selectCar.tabname, rowOffset + 1, values.Count);
                request2.IncludeGridData = true;
                Google.Apis.Sheets.v4.Data.Spreadsheet response2 = request2.Execute();

                for (int b = 0; b < values.Count - rowOffset; b++)
                {
                    var bval = (float)0.0;
                    var rval = (float)0.0;
                    var gval = (float)0.0;
#if SheetColors == true
                    if (response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Blue == null) bval = (float)0.0;
                    else bval = (float)response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Blue;
                    if (response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Red == null) rval = (float)0.0;
                    else rval = (float)response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Red;
                    if (response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Green == null) gval = (float)0.0;
                    else gval = (float)response2.Sheets[0].Data[0].RowData[b].Values[0].EffectiveFormat.BackgroundColor.Green;
                    System.Drawing.Color tcol = System.Drawing.Color.FromArgb(150, (Int32)(rval * 255), (Int32)(gval * 255), (Int32)(bval * 255));
                    bcolor.Add(tcol.Name);
#endif
                }

                NumofRec = 0;

                //
                //  Build Header Columns
                //
                foreach (var row in values)
                {
                    if (NumofRec == 0)
                    {
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
                    System.Drawing.Color ccolor;

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
#if SheetColors == true
                        if (z == 0)
                        {
                            ccolor = Program.HexStringToColor(bcolor[x]);
                            Console.WriteLine(ccolor.ToString());
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#fffbbc04")) dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Contacted 5 times letters\"";
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ffea9999")) dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Only Spanish\"";
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ffea4335")) dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"This Sheet\"";
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ff8e7cc3")) dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Confirmed English\"";
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ff7f6000"))
                            {
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Imported Names\"";
                                ccolor = System.Drawing.Color.SandyBrown; ;
                            }
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#fff3f3f3"))
                            {
                                ccolor = System.Drawing.ColorTranslator.FromHtml("#ffffffff");
                            }
                                dataGridView1.Rows[x].Cells[0].Style.BackColor = ccolor;
                        }
#endif
                    }

                    if (r1 == -1) r1 = firstrow;
                    if (comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 0)
                    {
                        dataGridView1.Rows[x].Visible = false;
                        string testStr = dataGridView1.Rows[x].Cells[4].Value?.ToString();
                        foreach (string kind in checkedListBox1.CheckedItems)
                        {
                            if ((testStr == kind || (string.IsNullOrEmpty(testStr) && kind == "blank") ||
                            (string.IsNullOrWhiteSpace(testStr) && kind == "blank"))
                            && (r1++ % skiprow) == 0)
                            {
                                dataGridView1.Rows[x].Visible = true;
                                break;
                            }
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
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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

            if (comboBox1.SelectedIndex == 0)               // Imported Names
            {
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                button8.Visible = true;
                textBox9.Visible = true;
                textBox10.Visible = true;
                label10.Visible = true;
                textBox6.Visible = true;
                textBox7.Visible = true;
                textBox8.Visible = true;
                button9.Visible = true;
                label9.Visible = false;
                button9.Visible = false;
                checkBox2.Visible = false;
                checkBox3.Visible = false;
                textBox9.Visible = true;
                textBox10.Visible = true;
                label10.Visible = true;
                textBox9.Text = dataGridView1.RowCount.ToString();
                textBox9.Update();
                textBox10.Text = dataGridView1.RowCount.ToString();
                textBox10.Update();
                button8.BackColor = System.Drawing.Color.LightGreen;  // indicate "find names" is next

            }

            if (comboBox1.SelectedIndex == 1)           // Field Service
            {
                button8.Visible = false;
                button9.Visible = true;
                label9.Visible = false;
                checkBox2.Visible = false;
                checkBox3.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                label10.Visible = false;
                if (comboBox1.SelectedIndex == 1) checkBox4.Visible = true;
                if (gEODhit)
                {
                    gEODhit = false;
                    button9.BackColor = System.Drawing.Color.LightGreen;
                }

            }

            if (comboBox1.SelectedIndex == 2)
            {
                button9.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;

            }
            if (comboBox1.SelectedIndex == 3)
            {
                button9.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;

            }
            if (comboBox1.SelectedIndex == 4)
            {
                button9.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;

            }
            if (comboBox1.SelectedIndex == 5)
            {
                button9.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
            }
            button2.BackColor = System.Drawing.Color.LightGray;
        }



        //---------------------------------------------------------------------------
        //
        // Clear Data 
        //
        private void button3_Click(object sender, EventArgs e)
        {

            if (DataChanged)
            {
                DialogResult result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to Clear and loose these changes?",
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
            textBox1.Update();
            cellch.Clear();
            button4.BackColor = System.Drawing.Color.LightGray;
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

            if (sender.GetType().Name != "Button")
                button2.BackColor = System.Drawing.Color.LightGreen;
            else
                button2.BackColor = System.Drawing.Color.LightGray;
            button2.Update();
        }

        //---------------------------------------------------------------------------
        //
        // Log out
        //
        private void button5_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                //
                // Test the results of the previous 3 dialogs.
                //
                if (result1 == DialogResult.Yes)
                    System.Windows.Forms.Application.Exit();
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
            if (comboBox1.SelectedIndex == 0 || updateinprogress || DataChanged)
            {
                return;
            }

            int intIndex = e.RowIndex;
            //int i = 0;
            dataGridView1.MultiSelect = false;
            bool localAutoGPS = false;
            bool localAutoMap = false;

            while (dr == DialogResult.OK || dr == DialogResult.None || dr == DialogResult.No || dr == DialogResult.Retry)
            {
                done = false;
                textBox2.ForeColor = System.Drawing.Color.Black;

                if (dataGridView1.CurrentCell.ColumnIndex == 0)
                {
                    dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect; // use to be FullRowSelect
                    dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;          // use to be EditOnF2;
                    dataGridView1.ReadOnly = false;

                    // Sheet Tab selection
                    int selectIndex = comboBox1.SelectedIndex;
                    sTabName selectCar = (sTabName)comboBox1.SelectedItem;
                    String range = selectCar.tabname + selectCar.range;

                    string cellResult = null;
                    string cellNote = null;

                    int nRow = dataGridView1.CurrentCell.RowIndex;
                    seeVisibleRow(dataGridView1, nRow);
                    if (selectIndex == 1)               // Field Service
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
                                //form2.textBox9.Text = textBox2.Text;
                                form2.textBox9.Text = dataGridView1.CurrentRow.Index.ToString();

                                form2.textBox7.Text = dataGridView1.CurrentRow.Cells[3].Value?.ToString();
                                string tmp = "";
                                string tmp0 = "";
                                if (string.IsNullOrEmpty(form2.textBox1.Text)) tmp = " no phone ";
                                else
                                {
                                    ToolTip myToolTip = new ToolTip();
                                    tmp = form2.textBox1.Text;
                                    form2.textBox1.BackColor= dataGridView1.Rows[nRow].Cells[0].Style.BackColor;
                                    myToolTip.SetToolTip(form2.textBox1, dataGridView1.CurrentRow.Cells[0].ToolTipText);
                                }
                                form2.textBox1.Update();

                                if (string.IsNullOrEmpty(form2.textBox7.Text)) tmp0 = " ";
                                else tmp0 = " - city " + form2.textBox7.Text;
                                try
                                {
                                    System.Windows.Clipboard.SetText(tmp + tmp0 + Environment.NewLine);
                                }
                                catch { }

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
                                try
                                {
                                    string temp2 = form2.textBox6.Text;
                                    string text = temp2.Replace(Environment.NewLine, "^"); // a random token
                                    string[] lines = text.Split('^');
                                    for (int j = 0; j < lines.Count(); j++)
                                    {
                                        string temp = form2.textBox6.Lines[j];
                                        firstpos = temp.Contains(today.ToString("yyyy-MM-dd"));
                                        if (firstpos) break;
                                    }

                                    if (!firstpos) form2.textBox6.Text = form2.textBox6.Text + Environment.NewLine + DateTime.Now.ToString("yyyy-MM-dd ddd hh:mm tt") + ": ";
                                    else form2.textBox6.Text = form2.textBox6.Text + Environment.NewLine;
                                }
                                catch
                                {
                                    form2.textBox6.Text = form2.textBox6.Text + Environment.NewLine + DateTime.Now.ToString("yyyy-MM-dd ddd hh:mm tt") + ": ";
                                }

                                switch (dataGridView1.CurrentRow.Cells[4].Value?.ToString())
                                {
                                    case "N/A":
                                        form2.radioButton1.Checked = true;
                                        break;
                                    case "Bus":
                                        form2.radioButton2.Checked = true;
                                        break;
                                    case "F":
                                        form2.radioButton3.Checked = true;
                                        break;
                                    case "DNC":
                                        form2.radioButton4.Checked = true;
                                        break;
                                    case "dF":
                                        form2.radioButton5.Checked = true;
                                        break;
                                    case "dN":
                                        form2.radioButton6.Checked = true;
                                        break;
                                    case "In":
                                        form2.radioButton7.Checked = true;
                                        break;
                                    case "No":
                                        form2.radioButton8.Checked = true;
                                        break;
                                    case "O":
                                        form2.radioButton9.Checked = true;
                                        break;
                                    case "X":
                                        form2.radioButton10.Checked = true;
                                        break;
                                    default:
                                        break;
                                }

                                form2.textBox6.SelectionStart = form2.textBox6.Text.Length + 4;
                                form2.textBox6.SelectionLength = 0;

                                form2.checkBox1.Checked = checkBox1.Checked;
                                form2.checkBox2.Checked = checkBox4.Checked;
                                form2.checkBox3.Checked = localAutoGPS;
                                if (form2.checkBox3.Checked)
                                    form2.button5.PerformClick();
                                form2.checkBox4.Checked = localAutoMap;
                                dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
                                adjGPS = false;
                                //  
                                // Bring up Form2
                                //  
                                dr = form2.ShowDialog();        // bring up the form

                                SecondFormLeft = form2.Left;    // save current position for next time
                                SecondFromTop = form2.Top;
                                SecondFormSize = form2.Size;
                                checkBox1.Checked = form2.checkBox1.Checked;

                                checkBox4.Checked = form2.checkBox2.Checked;
                                localAutoGPS = form2.checkBox3.Checked;
                                localAutoMap = form2.checkBox4.Checked;

                                textBox1.Text = cellch.Count.ToString() + " changes ";

                                if (dr == DialogResult.OK)
                                {

                                    saveFSdata(nRow);
                                    
                                    //
                                    // unselect current row and find the next visable
                                    //
                                    try
                                    {
                                        dataGridView1.Rows[nRow].Selected = false;
                                        while (nRow < dataGridView1.RowCount - 1)
                                            if (nRow < dataGridView1.RowCount)
                                            {
                                                if (dataGridView1.Rows[++nRow].Visible == false)
                                                    continue;
                                                if (checkBox4.Checked && (dataGridView1.Rows[nRow].DefaultCellStyle.BackColor != System.Drawing.Color.Aquamarine))
                                                    continue;
                                                dataGridView1.Rows[nRow].Selected = true;
                                                int currentRow = nRow;
                                                dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                                break;
                                            }
                                    }
                                    catch { }
                                }

                                //
                                // skip to next visable row processing
                                //
                                if (dr == DialogResult.No)
                                {

                                    //saveFSdata(nRow);

                                    try
                                    {
                                        dataGridView1.Rows[nRow].Selected = false;
                                        while (nRow < dataGridView1.RowCount - 1)
                                        {
                                            if (dataGridView1.Rows[++nRow].Visible == false)
                                                continue;
                                            if (checkBox4.Checked && (dataGridView1.Rows[nRow].DefaultCellStyle.BackColor != System.Drawing.Color.Aquamarine))
                                                continue;
                                            dataGridView1.Rows[nRow].Selected = true;
                                            int currentRow = nRow;
                                            dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                            break;
                                        }
                                    }
                                    catch { }

                                    dataGridView1.Refresh();
                                    continue;
                                }

                                //
                                // backup to previous visable row 
                                //
                                if (dr == DialogResult.Retry)
                                {

                                    //saveFSdata(nRow);

                                    try
                                    {
                                        dataGridView1.Rows[nRow].Selected = false;
                                        while (nRow > 0)
                                        {
                                            if (dataGridView1.Rows[--nRow].Visible == false)
                                                continue;
                                            if (checkBox4.Checked && (dataGridView1.Rows[nRow].DefaultCellStyle.BackColor != System.Drawing.Color.Aquamarine))
                                                continue;
                                            dataGridView1.Rows[nRow].Selected = true;
                                            int currentRow = nRow;
                                            dataGridView1.Rows[currentRow].Cells[0].Selected = true;
                                            break;
                                        }
                                    }
                                    catch { }

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
                }
                else break;
            }
            dr = DialogResult.None;

            done = true;
        }

        //
        // if autosave - save Field Service Data
        //
        private void saveFSdata(int nRow)
        {
            Form2 form2 = new Form2();
            string cellResult = "";

            dataGridView1.Rows[nRow].Selected = true;
            dataGridView1.Rows[nRow].Cells[0].Selected = true;

            //
            // process RadioButtons on Form2

                switch (checkedRadio)
                {
                    case "No Ans.":
                        cellResult = "N/A";
                        break;

                    case "Business":
                        cellResult = "Bus";
                        break;

                    case "Phone Fault":
                        cellResult = "F";
                        break;

                    case "In Terr.":
                        cellResult = "In";
                        break;

                    case "Out Terr.":
                        cellResult = "O";
                        break;

                    case "No Map":
                        cellResult = "X";
                        break;

                    case "Do Not Call":
                        cellResult = "DNC";
                        break;

                    case "Deaf in family":
                        cellResult = "dF";
                        break;

                    case "Deaf Neighbor":
                        cellResult = "dN";
                        break;

                    case "No Deaf":
                        cellResult = "No";
                        break;

                    default:
                        break;
                }
            //}
            DateTime today = DateTime.Today;

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Selected = true;
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = cellResult.ToString();

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Selected = true;
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value = today.ToString("yyyy-MM-dd");

            int number = 0;
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Selected = true;
            number = Convert.ToInt16(attempt);
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Value = number.ToString();

            notes = notes.Replace(",", "");   // remove all comma's
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Selected = true;
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Value = notes;

            if(adjGPS)
            {
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1,"Confidence")].Selected = true;
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = confid;

                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Selected = true;
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Value = lat;

                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Selected = true;
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Value = lon;

                adjGPS = false;
            }
        }
        //
        // detect click in header row
        //
        private void dataGridView1_ColumnHeaderMouseClick(
    object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dataGridView1.Columns[e.ColumnIndex];
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
            System.Windows.Forms.MessageBox.Show("Cannot delete a record!");
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
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                SaveSheetChanges(null);
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
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

        int maxChgs = Properties.Settings.Default.maxChg;
        int perSec = Properties.Settings.Default.perSec;
        int rows2del = 0;

        private void SaveSheetChanges(string Tabname)
        {
            if (Tabname == null && Tabfocus != null) Tabname = Tabfocus;
            updateinprogress = true;
            DataChanged = true;
            begTimInc = DateTime.MinValue;
            //
            // check login credentials
            //
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

            //
            // initialize process variables
            //
            progressBar1.Value = 4;
            progressBar1.Update();
            button4.BackColor = System.Drawing.Color.Coral;
            string sCol = null;
            string sRow = null;
            string oCol = null;
            string oRow = null;
            string fRow = null;
            string fCol = null;
            //string sSaveRow4Del = null;

            bool iterating = false;
            var sValue = new List<object>();

            int l = 0;
            int h = 0;

            int newr = -1;
            int oldr = -1;
            int newc = -1;
            int oldc = -1;
            int index2 = 0;
            //var dPoint = "";
            int dpCount = -1;
            string DataRow = "";

            string temp0 = "";
            string temp1 = "";
            string temp2 = "";
            string temp3 = "";
            string temp4 = "";
            bool toggle = false;
            bool modstr = false;
            bool conseq = false;

            //Copy items from list1 to list2
            //List<List<String>> cellch = new List<List<String>>();
            List<List<String>> cellch2 = new List<List<String>>();

            cellch2.Clear();
            // loop over tuples with the item and its index
            //foreach (var (dPoint, index) in cellch.Select((value, i) => (value, i))) 
            foreach(var dPoint in cellch)
            {
                if (dPoint == null) continue;      // skip blanks
                temp0 = dPoint[0].ToString();        // co-ordinates
                temp1 = dPoint[1].ToString();       //  its data
            Lb1: newc = Convert.ToInt32(Regex.Match(temp0, @"\d+").Value);  //Col from temp0
                newr = Convert.ToInt32(Regex.Match(temp0, @"(\d+)(?!.*\d)").Value); //Row form temp0
                if (!toggle)
                {
                    temp3 = temp0;      //save starting address
                    temp4 = temp0;
                    temp2 = temp1;      //save inital string
                    toggle = true;
                    modstr = true;
                    conseq = true;
                }

                if (newr != oldr)       // row change
                {
                    oldr = newr;

                    if (conseq)
                    {
                        modstr = true;
                        //conseq = false;
                    }
                    else
                    {
                        oldc = newc - 1;
                        modstr = false;
                    }


                }
                else conseq = true;

                if (conseq && (newc == oldc + 1 || oldc == -1))   // consecutive column
                {
                    if(oldc != -1)temp2 = temp2 + "," + temp1;
                    oldc = newc;
                    modstr = true;
                }
                else
                {
                    temp3 = temp4;
                    //conseq = false;
                    if (conseq)
                    {
                        modstr = false;
                        conseq = false;
                    }
                    else
                    {
                        modstr = false;
                        conseq = false;
                    }
                    oldc = newc;
                }

                if (!modstr)
                {
                    Console.WriteLine("non consecutive column");
                    int i = cellch2.Count;
                    cellch2.Add(new List<String>());
                    cellch2[i].Add(temp3);
                    cellch2[i].Add(temp2);

                    temp2 = "";
                    temp3 = temp0;
                    oldc = -1;

                    toggle = false;
                    modstr = true;
                    goto Lb1;
                    conseq = true;


                }
                continue;
            }
            if (!string.IsNullOrEmpty(temp2))
            {
                Console.WriteLine("purge out any leftover changes");
                int i = cellch2.Count;
                cellch2.Add(new List<String>());
                cellch2[i].Add(temp3);
                cellch2[i].Add(temp2);
            }
            toggle = false;
            conseq = false;
            conseq = false;
            oldc = -1;
            oldr = -1;
            cellch.Clear();     
            //
            // check total changes and processing time
            //   wait appropriately
            //
            foreach (var ListItem in cellch2)
            {
                fCol = null;
                if (true)
                {
                    if (checkBox1.Checked)
                    {
                        totalChgs = 1;
                        begTimInc = DateTime.MinValue;
                    }
                    nowTimInc = DateTime.Now;
                    if (begTimInc == DateTime.MinValue) begTimInc = nowTimInc;
                    TimeSpan duration = nowTimInc - begTimInc;
                    int d = (int)duration.TotalSeconds;
                    if (d > (2 * perSec))
                    {
                        begTimInc = nowTimInc;
                        d = (perSec);
                    }
                    if ((d >= perSec || totalChgs >= maxChgs))
                    {
                        for (int z = 0; z < (totalChgs > d ? totalChgs : d); z++)
                        {
                            if (checkBox1.Checked && d > perSec)
                            {
                                d = 0;
                                break;
                            }
                            textBox1.ForeColor = System.Drawing.Color.Red;
                            textBox1.Text = string.Format(" waiting {0} sec.", d - z);
                            textBox1.Update();
                            Thread.Sleep(1000);
                        }

                        begTimInc = DateTime.MinValue;
                        totalChgs = 0;
                        textBox1.ForeColor = System.Drawing.Color.Black;
                        textBox1.Text = " continuing";
                        textBox1.Update();
                    }
                }

                //
                // iterate through all sublist items converting to An:Bn format
                // and prepare to send
                //
                h = 0;
                int o = 0;
                int s = 0;
                int gCol = -1;
                oCol = null;
                chgCount = 0;
                bool dataready = false;

                foreach (string SubListItem in ListItem)        // calculate the whole row
                {
                    if (dataready == false)
                    {
                        if (SubListItem == null) continue;      // skip blanks
                        sCol = Regex.Match(SubListItem, @"\d+").Value;
                        sRow = Regex.Match(SubListItem, @"(\d+)(?!.*\d)").Value;

                        dataready = true;
                        h++;
                        o = 0;
                        if (!string.IsNullOrEmpty(oCol)) o = Convert.ToInt32(oCol) + 1;
                        if (!string.IsNullOrEmpty(sCol)) s = Convert.ToInt32(sCol);
                        if (s >= o)
                        {
                            if (o == Convert.ToInt32(sCol) || fCol == null)     // is it the next col?
                            {
                                if (fCol == null)                               // is it the first col?
                                    fCol = sCol;
                                fRow = sRow;
                                if (Convert.ToInt32(sCol) >= 0)
                                {
                                    oCol = sCol;                                // yes - keep going
                                    oRow = sRow;
                                }
                                iterating = true;
                                continue;
                            }
                            else
                            {
                                iterating = false;              // end of current row consecutive changes
                                dataready = false;
                            }
                            continue;
                        }
                        else dataready = false;
                        //continue;
                    }
                    if (dataready)
                    {
                        if (SubListItem == "C5") sValue.Add("N/A");
                        else sValue.Add(SubListItem.ToString());         // here is the associated data               
                    }

                    if (moves.Count > 0 && checkBox2.Checked)
                    {
                        foreach (var Item in moves)
                        {
                            string key = null;
                            foreach (string SubItem in Item)
                            {
                                if (key != null) Tabname = SubItem;
                                if (SubItem == SubListItem)
                                {
                                    if (SubItem != null)
                                    {
                                        key = SubItem;
                                        if (key == "skip") continue;
                                        else Tabname = SubItem;
                                    }
                                    if (SubItem == "pS" || SubItem == "SP")
                                    {
                                        numOfSP++;
                                        rows2del++;
                                        textBox7.Text = string.Format("{0} updated", numOfSP);
                                        textBox7.Update();
                                    }
                                    if (SubItem == "pE" || SubItem == "E")
                                    {
                                        numOfpE++;
                                        rows2del++;
                                        textBox6.Text = string.Format("{0} updated", numOfpE);
                                        textBox6.Update();
                                    }
                                    if (
                                        SubItem == "N/A" ||
                                        SubItem == "B" ||
                                        SubItem == "I" ||
                                        SubItem == "blank" ||
                                        SubItem == "DNC")
                                    {
                                        numOfSkip++;
                                        textBox8.Text = string.Format("{0} updated", numOfSkip);
                                        textBox8.Update();
                                    }
                                }
                            }
                        }
                    }
                    dataready = false;                          // prepare for next row
                                                                //break;
                }
                //}
                if (!String.IsNullOrEmpty(oRow))                // increment to next row
                {
                    int number = Convert.ToInt32(oRow);
                    number = number + rowOffset + 1;
                    sRow = number.ToString();
                }
                string stmp = "";
                if (Tabname == "Field Service") stmp = "FS";
                if (Tabname == "Only Spanish") stmp = "OT";
                if (Tabname == "Imported Names") stmp = "IM";
                //if (Tabname == "Other") stmp = "O";
                //if (Tabname == "Contacted 5 times letters") stmp = "5X";
                //if (Tabname == "Confirmed English") stmp = "EN";
                textBox2.Text = string.Format(" "+ stmp + " row {0}", sRow);
                textBox2.Update();

                if (!String.IsNullOrEmpty(oCol))                // check col is good
                    sCol = Program.ColumnAdress(Int32.Parse(oCol));

                    int commas = sValue.Last().ToString().Count(c => c == ',');
                    fCol = Program.ColumnAdress(Int32.Parse(fCol) + commas);

                //
                //
                // prepare to send batch update to Googlesheet
                //
                String spreadsheetId2 = spreadsheetId;
                sTabName selectCar = null;
                if (Tabname == null)                      // determine destination tab
                {
                    int selectIndex = comboBox1.SelectedIndex;
                    selectCar = (sTabName)comboBox1.SelectedItem;
                    Tabname = selectCar.tabname;
                }

                //
                // run our request
                //
                if (Tabname != "delete")
                {
                    IList<IList<object>> oblist = new List<IList<object>>();

                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "ROWS";// "ROWS";//COLUMNS

                    string rValue = sValue[0].ToString();
                    List<string> tValue = rValue.Split(',').ToList();
                    oblist.Add(new List<object>());
                    for (int z = 0; z < tValue.Count(); z++)
                    { 
                        oblist[0].Add(tValue[z]);
                    }

                    valueRange.Values = oblist;

                    //
                    // execute request
                    //
                    String range2 = Tabname + "!" + sCol + sRow + ":" + fCol + sRow;  // cells to update on Tab 
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result2 = update.Execute();
                    sValue.Remove(sValue[0]);
                }

                //
                // delete processed rows in datagrid and Google imported name sheet
                //
                if (checkBox3.Checked)       /*(checkBox3.Checked*/
                {
                    List<Request> deleteRequestsList = new List<Request>();
                    BatchUpdateSpreadsheetRequest _batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                    Request _deleteRequest = new Request();
                    _deleteRequest.DeleteDimension = new DeleteDimensionRequest();
                    _deleteRequest.DeleteDimension.Range = new DimensionRange();

                    if (Tabfocus == null)
                    {
                        int selectIndex = comboBox1.SelectedIndex;
                        selectCar = (sTabName)comboBox1.SelectedItem;
                        Tabfocus = selectCar.tabname;
                    }
                    if (Tabfocus == "Field Service")
                        _deleteRequest.DeleteDimension.Range.SheetId = Properties.Settings.Default.FSsheetID;
                    if (Tabfocus == "Imported Names")
                        _deleteRequest.DeleteDimension.Range.SheetId = Properties.Settings.Default.INsheetID;
                    _deleteRequest.DeleteDimension.Range.Dimension = "ROWS";
                    if (checkBox2.Checked)
                    {
                        int rowIndex = Convert.ToInt32(sRow);

                        if (sSaveRow4Del.Count > 0)
                        {
                            rowIndex = Convert.ToInt32(sSaveRow4Del[chgCount]);
                            rowIndex = rowIndex - chgCount++;

                            _deleteRequest.DeleteDimension.Range.StartIndex = rowIndex + rowOffset;
                            _deleteRequest.DeleteDimension.Range.EndIndex = rowIndex + rowOffset + 1;


                            deleteRequestsList.Add(_deleteRequest);
                            _batchUpdateSpreadsheetRequest.Requests = deleteRequestsList;
                            service.Spreadsheets.BatchUpdate(_batchUpdateSpreadsheetRequest, spreadsheetId).Execute();

                            for (int j = 0; j < dataGridView1.Rows.Count; j++)
                            {
                                if (dataGridView1.Rows[j].Visible == false) continue;
                                dataGridView1.Rows[j].Visible = false;
                                for (int k = j; k < dataGridView1.Rows.Count; k++)
                                {
                                    if (dataGridView1.Rows[k].Visible == false) continue;
                                    dataGridView1.CurrentCell = dataGridView1[0, k];
                                    break;
                                }
                                dataGridView1.Update();
                                textBox1.Text = string.Format("DG row {0}", j);
                                textBox1.Update();
                                break;
                            }
                            dataGridView1.Refresh();
                        }
                        dataGridView1.Update();
                        sValue.Clear();
                        continue;
                    }

                }

                dataGridView1.Update();
                dataGridView1.Visible = true;
                totalChgs++;
                sValue.Clear();
                //
                // update progress bar
                //
                int remainder;
                l++;
                Math.DivRem(l, cellch2.Count, out remainder);
                if (cellch2.Count > 100)
                    progress = (remainder / (cellch2.Count / 100));
                else
                    progress = remainder * ((100 - 1) / cellch2.Count);
                if (progress < 100) progressBar1.Value = progress;
                progressBar1.Update();
            }
            //
            // Clean up and finish
            //
            updateinprogress = false;
            DataChanged = false;
            cellch2.Clear();
            textBox1.Text = " done ";
            textBox1.Update();
            progressBar1.Value = 100;
            button4.BackColor = System.Drawing.Color.LightGray;
            if (comboBox1.SelectedIndex == 0)
            {
                checkBox2.Checked = true;
                checkBox3.Checked = true;
            }
            return;
        }
        //
        //
        //
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
                SaveSheetChanges(Tabfocus);
            else
                SaveSheetChanges(null);
            cellch.Clear();
        }

        public void startPB(System.Drawing.Color color)
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
            checkBox2.Checked = false;
            checkBox2.Checked = false;

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

                    cellch.Add(new List<String>());                     //Adds new sub List
                    cellch[i].Add(dgv.CurrentCellAddress.ToString());   //Add values to the sub List at index 0
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

                        }
                    }
                    dgv.CurrentCell = oldcurcell;
                    textBox1.Text = cellch.Count.ToString() + " changes ";
                    textBox1.Update();
                    textBox2.Text = dgv.CurrentCell.ToString();
                    textBox2.Update();
                    if (checkBox1.Checked) SaveSheetChanges(null);
                }
                UpdateSheet();
            }
        }

        private void dataGridView1_SelectionChanged(object sender, DataGridViewRowEventArgs e)
        {
            textBox2.Text = e.Row.ToString();
            textBox2.Update();
        }

        private void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            button2.BackColor = System.Drawing.Color.LightGreen;
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
                    changeSummary = changeSummary + SubListItem + " \n";
                }
            }
            changeSummary = changeSummary + "\n----------------\n";
            // Create a new instance of the Form2 class
            Form0 changesForm = new Form0();
            changesForm.textBox2.Text = cellch.Count.ToString();
            textBox2.Update();
            changesForm.textBox1.Text = changeSummary;
            textBox1.Update();

            // Show the settings form
            dr = changesForm.ShowDialog();

            if (dr == DialogResult.Abort)
            {
                cellch.Clear();
                DataChanged = false;
            }

            textBox1.Text = cellch.Count.ToString() + " changes";
            textBox1.Update();
            return;
        }


        private void button6_Click(object sender, EventArgs e)
        {
            //waiting = false;
        }

        //
        // sheet tab selection
        //
        private void combox1_selectedIndexChange(object sender, EventArgs e)
        {
            checkedListBox1.Items.Clear();
            if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == 1)
            {
                string[] checklist = new string[] { "N/A", "Bus", "F", "DNC", "dF", "dN", "blank", "No", "O", "In", "X" };
                for (int i = 0; i < checklist.Length; i++)
                {
                    checkedListBox1.Items.Add(checklist[i]);
                }
                checkedListBox1.Visible = true;
                label1.Visible = true;
                textBox3.Visible = true;
                textBox5.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                label1.Visible = true;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                button8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                label10.Visible = false;
                textBox6.Visible = false;
                textBox7.Visible = false;
                textBox8.Visible = false;
                button9.Visible = false;
                button10.Visible = false;
                label9.Visible = false;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox2.Visible = false;                  // move
                checkBox3.Visible = false;                  // delete
                checkBox4.Visible = false;
                button11.Visible = false;                   // suggest next

                checkedListBox1.CheckOnClick = true;
            }

            if (comboBox1.SelectedIndex == 0)               // imported names
            {
                button10.Visible = true;
                checkedListBox1.Visible = true;
                label1.Visible = true;
                textBox3.Visible = true;
                textBox5.Visible = true;
                label5.Visible = true;
                label4.Visible = true;
                button9.Visible = false;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox4.Visible = false;
                textBox9.Visible = true;
                textBox10.Visible = true;
                label10.Visible = true;
                checkedListBox1.CheckOnClick = true;
                checkedListBox1.SetItemChecked(6, true);                  //"blank";
                checkBox1.Checked = true;
            }

            if (comboBox1.SelectedIndex == 1)           // Field Service
            {
                button10.Visible = true;
                checkBox2.Checked = true;                                   // move but dont delete
                checkBox3.Checked = false;
                checkBox4.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                label10.Visible = false;
                checkedListBox1.SetItemChecked(0, true);                    // "N/A";
                checkedListBox1.SetItemChecked(6, true);                    // "<blank>"
                checkedListBox1.SetItemChecked(9, true);                    // "In"
                checkBox1.Checked = true;
            }
            if (comboBox1.SelectedIndex == 2)
            {
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
            }

            if (comboBox1.SelectedIndex == 3)
            {
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
            }

            if (comboBox1.SelectedIndex == 4)
            {
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
            }

            if (comboBox1.SelectedIndex == 5)
            {
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
            }
            button2.BackColor = System.Drawing.Color.LightGreen;
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
            var rSelected = new List<DataGridViewRow>();
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
                seeVisibleRow(dataGridView1, rowIndex);
            }
            else return;
            //
            //
            //
        }

        //
        // hit the search button on searchbox enter key
        //
        void TextBox4_GotFocus(object sender, EventArgs e)
        {
            this.AcceptButton = button6;
        }

        //
        // Prev. Search result
        //
        private void button13_Click(object sender, EventArgs e)
        {
            var rSelected = new List<DataGridViewRow>();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                rSelected.Add(r);

            if (curFound == 0) curFound = firstFound;
            for (int i = curFound - 1; i > 0; --i)
            {
                if (dataGridView1.Rows[i].Visible == false) continue;
                if (dataGridView1.Rows[i].Selected)
                {
                    curFound = i;
                    //dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                    //Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
                    dataGridView1.CurrentCell = dataGridView1[0, i];
                    seeVisibleRow(dataGridView1, i);
                    break;
                }
            }
            foreach (DataGridViewRow rs in rSelected)
                dataGridView1.Rows[rs.Index].Selected = true;

        }

        //
        // Find NExt
        //
        private void button12_Click(object sender, EventArgs e)
        {
            var rSelected = new List<DataGridViewRow>();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;
            foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                rSelected.Add(r);

            if (curFound == 0) curFound = firstFound;
            for (int i = curFound + 1; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Visible == false) continue;
                if (dataGridView1.Rows[i].Selected)
                {
                    curFound = i;
                    dataGridView1.CurrentCell = dataGridView1[0, i];
                    seeVisibleRow(dataGridView1, i);
                    break;
                }
            }
            foreach (DataGridViewRow rs in rSelected)
                dataGridView1.Rows[rs.Index].Selected = true;

        }

        //
        // find DGV column index by name
        //
        public int GetindexOf(DataGridView dgv, string name)
        {
            int colInd= dgv.Columns[name].Index;
            return (colInd);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //
        // search datagrid for value
        private void button6_Click_1(object sender, EventArgs e)
        {
            var rSelected = new List<DataGridViewRow>();
            string searchValue = "";
            searchValue = textBox4.Text;
            curFound = 0;
            foundCnt = 0;

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
                            if (firstFound == 0) firstFound = row.Index;
                            foundCnt++;
                            break;
                        }
                    }

                    continue;
                }
                System.Windows.Forms.MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));

                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    rSelected.Add(r);
                dataGridView1.CurrentCell = dataGridView1[visibleColumnIndex, firstFound];
                seeVisibleRow(dataGridView1, firstFound);
                foreach (DataGridViewRow rs in rSelected)
                    dataGridView1.Rows[rs.Index].Selected = true;

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));
            }
        }
        //
        // Clear search box and Selected Rows
        //
        private void button7_Click(object sender, EventArgs e)
        {
            foundCnt = 0;
            curFound = 0;
            textBox4.Clear();
            dataGridView1.ClearSelection();
        }

        //
        //
        //
        public void SetResultTime(string result)
        {
            DateTime today = DateTime.Today;
            int c;
            //if (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value != result)
            //{
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Selected = true;
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = result;
                c = cellch.Count;
                cellch.Add(new List<String>());
                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                cellch[c].Add(result);
            //}
            //if (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value != today.ToString("yyyy-MM-dd"))
            //{
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value = today.ToString("yyyy-MM-dd");
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Selected = true;
                c = cellch.Count;
                cellch.Add(new List<String>());
                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                cellch[c].Add(today.ToString("yyyy-MM-dd"));
            //}

            //if (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Value != "0");
            //{
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Value = "0";
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Selected = true;
                c = cellch.Count;
                cellch.Add(new List<String>());
                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                cellch[c].Add("0");
            //}

            //if (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Value != " ") ;
            //{
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Value = " ";
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Selected = true;
                c = cellch.Count;
                cellch.Add(new List<String>());
                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                cellch[c].Add(" ");
            //}

        }
        //
        //
        // RunImported Names
        //
        List<List<String>> names2chk = new List<List<String>>();
        private void RunImportedNames(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            int k = 0;
            int nLim = 0;
            button8.BackColor = System.Drawing.Color.Coral;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            using (var UserControl1 = new UserControl1())
            {
                startPB(System.Drawing.Color.Green);
                cellch.Clear();
                textBox1.Text = " ..logging in ";
                textBox1.Update();

                // If modifying these scopes, delete your previously saved credentials
                // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
                string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                string ApplicationName = "Google Sheets API .NET Quickstart";

                //UserCredential credential;

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

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
                progressBar1.Value = 2;
                progressBar1.Update();
                // Define request parameters.
                if (radioButton2.Checked) spreadsheetId = Properties.Settings.Default.TestSheet;
                if (radioButton2.Checked) spreadsheetId = Properties.Settings.Default.ProdSheet;
                progressBar1.Value = 3;
                progressBar1.Update();
                //String range = "Common First Names!A1:A";
                //SpreadsheetsResource.ValuesResource.GetRequest request =
                //        service.Spreadsheets.Values.Get(spreadsheetId, range);
                //ValueRange response = request.Execute();

                textBox1.Text = ".. reading data";
                textBox1.Update();
                //names2chk.Clear();
                DataChanged = true;
                int numOfSP = 0;
                textBox7.Clear();
                int numOfEN = 0;
                textBox6.Clear();
                int numOfSkip = 0;
                textBox8.Clear();
                IList<IList<Object>> values = null;

                int i;
                bool found = false;
                
                //foreach (var row in values)
                //{
                //    i = names2chk.Count;
                //    names2chk.Add(new List<String>()); //Adds new sub List
                //    names2chk[i].Add(values[i][0].ToString()); //Add values to the sub List at index 0
                //    names2chk[i].Add(" ".ToString());
                //}
                progressBar1.Value = 4;
                progressBar1.Update();

                //dataGridView1.Visible = false;

                //
                // set up destination tabs for Imported Names
                //
                int m = 0;
                moves.Clear();
                m = moves.Count;
                moves.Add(new List<String>());
                moves[m].Add("I");
                moves[m].Add("Field Service");
                m = moves.Count;
                moves.Add(new List<String>());
                moves[m].Add("O");
                moves[m].Add("Only Spanish");

                int nRow = 0;
                int nOff = 0;
                Int32.TryParse(textBox9.Text, out nLim);
                while ((nLim >= dataGridView1.RowCount || dataGridView1.Rows[nLim].Visible == false ) && nLim > 0 )       // dont end on an invisable row
                    nLim--;
                textBox9.Text = Convert.ToString(nLim);
                textBox9.Update();
            LB1: while (nRow < nLim/*+nOff*/)        //dataGridView1.RowCount)
                {
                    if (dataGridView1.Rows[nRow].Visible == false)
                    {
                        nRow++;
                        nOff++;
                        continue;
                    }

                    dataGridView1.Rows[nRow].Selected = true;
                    int currentRow = nRow;
                    dataGridView1.Rows[currentRow].Cells[GetindexOf(dataGridView1, "ADDRESS")].Selected = true;
                    dataGridView1.CurrentRow.Selected = true;
                    break;
                }
                if (nRow >= nLim/*+nOff*/)
                {
                    updateinprogress = false;
                    dataGridView1.Visible = true;
                    goto LB2;
                }
                int l = 0;
                int c = cellch.Count;

                //foreach (var row in values)
                //{
                    string streetadr = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "ADDRESS")].Value?.ToString();
                    string names = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NAME")].Value?.ToString();
                    string phone = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "TELEPHONE")].Value?.ToString();
                    if ((string.IsNullOrEmpty(names) || string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(streetadr)))
                    {
                    if (nRow < nLim)
                    {
                        //numOfSkip++;
                        nRow++;
                    }
                        goto LB1;
                    }

                    string phone1 = Regex.Replace(phone, "[^0-9]", "");
                    if(phone != phone1)
                    {
                        dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "TELEPHONE")].Selected = true;
                        dataGridView1.CurrentRow.Cells[0].Value = phone1;
                        cellch.Add(new List<String>());
                        cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[c].Add(phone1);
                    }
                    

                        if (dataGridView1.CurrentRow.Visible)
                        {
                            dataGridView1.CurrentRow.Selected = true;
                            Thread.Sleep(nRow / 20);
                            dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "ADDRESS")].Selected = true;
                            dataGridView1.Update();
                            Thread.Sleep(100);                // give form a chance to update
                            streetadr = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "ADDRESS")].Value?.ToString();
                            string cityadr = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "CITY")].Value?.ToString();
                            if (string.IsNullOrEmpty(cityadr)) cityadr = "Quito";
                            RegexOptions options = RegexOptions.None;               // remove multiple spaces
                            Regex regex = new Regex("[ ]{2,}", options);
                            streetadr = regex.Replace(streetadr, " ");
                            if (streetadr.Length > 120)                             // limit url size
                                streetadr = streetadr.Substring(0, 120);
                            streetadr = RestSharp.Extensions.MonoHttp.HttpUtility.UrlEncode(streetadr); // encode for specail characters
                            vs = getGPSfromAddr(streetadr, cityadr);
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value = today.ToString("yyyy-MM-dd");
                            if (vs == null || vs[0] == "E" || vs[0] == "X")
                            {
                                found = false;

                            }
                            else
                            {
                                found = true;
                            }
                        }

                    string t1 = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value?.ToString();

                    if (found)
                    {
                        if (vs[0] == "In" || vs[0] == "G")
                        {
                            SetResultTime("In");
                            numOfEN++;
                            textBox6.Text = string.Format("found {0}", numOfEN);
                            textBox6.Update();

                            if (vs != null)
                            {
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Selected = true;
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = vs[1];
                                c = cellch.Count;
                                cellch.Add(new List<String>());
                                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                                cellch[c].Add(vs[1]);
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Selected = true;
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Value = vs[2];
                                c = cellch.Count;
                                cellch.Add(new List<String>());
                                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                                cellch[c].Add(vs[2]);
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Selected = true;
                                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Value = vs[3];
                                c = cellch.Count;
                                cellch.Add(new List<String>());
                                cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                                cellch[c].Add(vs[3]);
                            }
                        }
                    }
                    else
                    {
                    if (vs != null && (vs[0] == "X" || vs[0] == "O" || vs[0] == "E"))
                    {
                        if (vs[0] == "E") SetResultTime("O");
                        else SetResultTime(vs[0]);
                        if (vs[0] == "O" || vs[0] == "E")
                        {
                            numOfSP++;
                            textBox7.Text = string.Format("found {0}", numOfSP);
                            textBox7.Update();
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = vs[1];
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add(vs[1]);
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Value = vs[2];
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add(vs[2]);
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Value = vs[3];
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add(vs[3]);
                        }
                        //}
                        else
                        {
                            numOfSkip++;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = vs[0];
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add(vs[0]);
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Value = "N/A";
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add("N/A");
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Selected = true;
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Value = "N/A";
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add("N/A");
                        }

                    }
                    else
                    if (vs == null) numOfSkip++;
                    } 
                    textBox8.Text = string.Format("skipped {0}", numOfSkip);
                    textBox8.Update();
                    textBox1.Text = cellch.Count.ToString() + " changes";
                    textBox1.Update();
                    if (checkBox1.Checked) SaveSheetChanges(null);
                    nRow++;

                    int remainder;
                    k++;
                    Math.DivRem(k, nLim, out remainder);
                    if (nLim > 100)
                        progress = (remainder / (nLim / 100));
                    else
                        progress = remainder * ((100 - 1) / nLim);
                    if (progress < 100) progressBar1.Value = progress;
                    progressBar1.Update();

                    goto LB1;
                }
                updateinprogress = false;
                dataGridView1.Visible = true;
                for (int y = 0; y < nLim; y++)
                    if (dataGridView1.Rows[y].Visible == true)
                    {
                        dataGridView1.Rows[y].Cells[0].Selected = true;
                        dataGridView1.Update();
                        break;
                    }

            //}
        LB2: if (comboBox1.SelectedIndex == 0)
            {
                progressBar1.Value = 100;
                progressBar1.Update();
                label9.Visible = true;
                button9.Visible = true;
                checkBox2.Visible = true;
                checkBox3.Visible = true;
                button8.BackColor = System.Drawing.Color.LightGray;
                button9.BackColor = System.Drawing.Color.LightGreen;
            }
            Tabfocus = "Imported Names";
            return;
        }

        //
        // get GPS coordinates from address and city
        //
        public string[] getGPSfromAddr(string address, string city)
        {
            bool inTerritory = false;
            string strx = "";
            string stry = "";
            int pos = -1;
            int pos1 = -1;
            string NewText = "";
            HttpWebRequestHandler hTTPrequest = new HttpWebRequestHandler();
            //http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=Guamani&postalCode=-&addressLine={addressLine}&userLocation=-&userIp={-}&usermapView={usermapView}&includeNeighborhood=includeNeighborhood&maxResults={maxResults}&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS
            string webAdr = @"http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=" + city + "&postalCode=-&addressLine=" + address + "&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS";
            if (!string.IsNullOrEmpty(address))
            {
                var webReply = hTTPrequest.GetReply(webAdr);
                pos = webReply.IndexOf("\"coordinates\":");
                pos1 = webReply.IndexOf("\"confidence\":");
                if (pos1 > -1)
                    NewText = webReply.Substring(pos1 + 14, 1);
                if (pos > -1)
                {
                    string pwebReply = webReply.Substring(pos, 56);
                    pos = pwebReply.IndexOf(',');
                    strx = pwebReply.Substring(15, pos - 15);

                    int pos2 = pwebReply.IndexOf(']');
                    stry = pwebReply.Substring(pos + 1, (pos2 - pos) - 1);
                }
                else return (null);

            }
            else return (null);
            pos = -1;

            if (!string.IsNullOrEmpty(strx) || !string.IsNullOrEmpty(stry))
            {
                float x = float.Parse(strx);
                float y = float.Parse(stry);

                GPSgeofence gPSgeofence = new GPSgeofence();
                GPSgeofence fence = gPSgeofence;
                fence.ReadGPSfence();
                inTerritory = fence.PointInPolygon(x, y);
                if(x == notfoundlat && y == notfoundlon)
                {
                    string[] reText = { "X", NewText, strx, stry };
                    return (reText);
                }

                if (inTerritory)
                {
                    string[] retText = { "G", NewText, strx, stry };
                    return (retText);
                }
                else
                {
                    string[] retText = { "E", NewText, strx, stry };
                    return (retText);
                }
            }
            else
                return (null);
        }

        //
        // Get all known phone numbers
        //
        private void getAllPhoneNumbers(SheetsService sheetService, string spreadsheet)
        {
            //-----------------------------------------------------------
            //
            // get phone numbers in Only Spanish
            //
            string range = "Only Spanish!A5:A";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetService.Spreadsheets.Values.Get(spreadsheet, range);

            ValueRange respnse = request.Execute();
            values = respnse.Values;  // list of phone numbers and source
            lastSProw = respnse.Values.Count;
            textBox1.Text = ".. reading data";
            textBox1.Update();
            int p = 2;

            names2chk.Clear();                              // clear phone array
            DataChanged = true;

            int i;
            int j = 0;
            string forchk = "";
            foreach (var row in values)
            {
                i = names2chk.Count;

                if (values[j++].Count == 0) continue;
                else forchk = values[j - 1][0].ToString();
                names2chk.Add(new List<String>());      //Adds new sub List
                names2chk[i].Add(forchk);               //Add values to the sub List at index 0
                names2chk[i].Add("Only Spanish".ToString());
                if (names2chk.Distinct().Count() != names2chk.Count())
                {
                    System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                }

            }
            p = p + 2;
            progressBar1.Value = p;
            progressBar1.Update();

            //------------------------------------------------------------
            //
            // get phone numbers in Field Service
            //
            j = 0;
            range = "Field Service!A5:A";
            request = sheetService.Spreadsheets.Values.Get(spreadsheetId, range);
            respnse = request.Execute();
            textBox1.Text = ".. reading data";
            lastFSrow = respnse.Values.Count;
            textBox1.Update();
            values = respnse.Values;
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
                    System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                }
            }
            p = p + 2;
            progressBar1.Value = p;
            progressBar1.Update();

            //------------------------------------------------------------
            //
            // get phone numbers in Confirmed English
            //
            j = 0;
            range = "Confirmed English!A5:A";
            request = sheetService.Spreadsheets.Values.Get(spreadsheetId, range);
            respnse = request.Execute();
            textBox1.Text = ".. reading data";
            if (respnse.Values == null) lastCErow = 0;
            else lastCErow = respnse.Values.Count;
            textBox1.Update();
            values = respnse.Values;
            forchk = "";
            if (lastCErow > 0)
            {
                foreach (var row in values)
                {
                    i = names2chk.Count;

                    if (values[j++].Count == 0) continue;
                    else forchk = values[j - 1][0].ToString();
                    names2chk.Add(new List<String>());
                    names2chk[i].Add(forchk); //Add values to the sub List at index 0
                    names2chk[i].Add("Confirmed English".ToString());
                    if (names2chk.Distinct().Count() != names2chk.Count())
                    {
                        System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                    }
                }
            }
            p = p + 2;
            progressBar1.Value = p;
            progressBar1.Update();
            //------------------------------------------------------------
            //
            // get phone numbers in Contacted 5 times letters
            //
            j = 0;
            range = "Contacted 5 times letters!A5:A";
            request = sheetService.Spreadsheets.Values.Get(spreadsheetId, range);
            respnse = request.Execute();
            textBox1.Text = ".. reading data";
            lastC5row = respnse.Values.Count;
            textBox1.Update();
            values = respnse.Values;
            forchk = "";
            foreach (var row in values)
            {
                i = names2chk.Count;

                if (values[j++].Count == 0) continue;
                else forchk = values[j - 1][0].ToString();
                names2chk.Add(new List<String>());
                names2chk[i].Add(forchk); //Add values to the sub List at index 0
                names2chk[i].Add("Contacted 5 times letters".ToString());
                if (names2chk.Distinct().Count() != names2chk.Count())
                {
                    System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                }
            }
            p = p + 2;
            progressBar1.Value = p;
            progressBar1.Update();
        }


        //--------------------------------------------------------------------------
        //
        //
        //  copy a range of new names to either E or SP lists
        //  as part of EOD activities
        //
        private void copyrange(object sender, EventArgs e)
        {
            numOfSP = 0;
            numOfpE = 0;
            numOfSkip = 0;
            DateTime today = DateTime.Today;

            label6.Visible = true;
            textBox6.Visible = true;
            textBox6.Update();
            label7.Visible = true;
            textBox7.Visible = true;
            textBox7.Update();
            label8.Visible = true;
            textBox8.Visible = true;
            textBox8.Update();
            checkBox4.Visible = false;
            button9.BackColor = System.Drawing.Color.Coral;
            checkBox2.Checked = true;
            checkBox2.Update();
            checkBox3.Checked = true;
            checkBox3.Update();
            sSaveRow4Del.Clear();
            Thread.Sleep(5);                // give form a chance to update

            using (var UserControl1 = new UserControl1())
            {
                startPB(System.Drawing.Color.Yellow);
                progressBar1.Value = 2;
                progressBar1.Update();
                cellch.Clear();
                textBox1.Text = " ..logging in ";
                textBox1.Update();
                sSaveRow4Del.Clear();

                // If modifying these scopes, delete your previously saved credentials
                // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
                string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
                string ApplicationName = "Google Sheets API .NET Quickstart";

                //UserCredential credential;

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

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
                progressBar1.Value = 3;
                progressBar1.Update();
                // Create Google Sheets API service.
                var sheetsService = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define request parameters.
                string spreadsheetId1 = Properties.Settings.Default.TestSheet;     // test sheet
                string spreadsheetId2 = Properties.Settings.Default.ProdSheet;      // live sheet  /edit#gid=2145664999"

                //
                // process RadioButtons on Form1

                if (radioButton1.Checked) spreadsheetId = spreadsheetId1;
                if (radioButton2.Checked) spreadsheetId = spreadsheetId2;
                textBox1.Text = ".. get all phone #";
                textBox1.Update();
                getAllPhoneNumbers(sheetsService, spreadsheetId);

                progressBar1.Value = 5;
                progressBar1.Update();
                dataGridView1.Visible = false;

                //-------------------------------------------------------
                //
                // prepare move information to "Only Spansh" or "Confired English
                // as part of "Imported Names" EOD
                //
                //
                int nRow = 0;
                int n = -1;
                int o = -1;
                foreach (DataGridViewRow row in dataGridView1.Rows)  // find first viable row
                {
                    if (row.Cells[0].Visible == false) continue;
                    nRow = row.Index;
                    break;
                }

                List<List<string>> svalue = new List<List<string>>();
                string tmp = "";
                sTabName selectFrom = (sTabName)comboBox1.SelectedItem;

                if (comboBox1.SelectedIndex == 0)           // EOD for Imported Names
                {
                    // setup for moves to other tabs based on 'Results'
                    //moves.Add(new List<String>());
                    //moves[m].Add("pE");
                    //moves[m].Add("Field Service");
                    int m = 0;
                    moves.Clear();
                    m = moves.Count;
                    moves.Add(new List<String>());
                    moves[m].Add("I");
                    moves[m].Add("Field Service");
                    m = moves.Count;
                    moves.Add(new List<String>());
                    moves[m].Add("O");
                    moves[m].Add("Only Spanish");

                    nRow = 0;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Visible == false) continue;

                        int remainder;
                        Math.DivRem(row.Index, dataGridView1.Rows.Count, out remainder);
                        if (dataGridView1.Rows.Count > 100)
                            progress = (remainder / (dataGridView1.Rows.Count / 100));
                        else
                            progress = remainder * ((100 - 1) / dataGridView1.Rows.Count);
                        if (progress < 99) progressBar1.Value = progress;
                        progressBar1.Update();

                        int t = 0;
                        if (o != row.Index)
                        {
                            o = row.Index;
                            int a = sSaveRow4Del.Count;
                            sSaveRow4Del.Add(string.Format("{0}", o));
                        }
                        if (!string.IsNullOrEmpty(row.Cells[4].Value?.ToString()))
                        {
                            if (row.Cells[4].Value.ToString() == "I") t = lastFSrow++; // 'pE' goes to Files Service
                            if (row.Cells[4].Value.ToString() == "O") t = lastSProw++; // 'pS' goes to Only Spanish
                        }
                        seeVisibleRow(dataGridView1, o);
                        dataGridView1.Update();

                        n = dataGridView1.CurrentCellAddress.X;         // save row for later
                        t = n + t - 4;                                      // calculate last row in target sheet

                        for (int w = 0; w < row.Cells.Count; w++)       // list every change location and data
                        {
                            int c = cellch.Count;                           // build list of changes
                            cellch.Add(new List<String>());
                            string ts = string.Format("{{X={0},Y={1}}}", w, t);
                            cellch[c].Add(ts.ToString());
                            if (string.IsNullOrEmpty(row.Cells[w].Value?.ToString())) cellch[c].Add(tmp);
                            else cellch[c].Add(row.Cells[w].Value?.ToString());
                        }
                        textBox1.Text = string.Format("{0} changes", cellch.Count); // total the changes
                        textBox1.Update();

                        if (checkBox1.Checked == true) SaveSheetChanges("Field Service"); // auto update?
                        else continue;
                    }
                    Tabfocus = "Imported Names";
                }

                if (comboBox1.SelectedIndex == 1)            // process EOD for Field Service
                {
                    //-------------------------------------------------------
                    //
                    // prepare move information to "Only Spanish", "Confirmed English" or "Other"
                    // as part of "Field Service" EOD
                    //
                    //int m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("pE");
                    //moves[m].Add("Field Service");
                    moves.Clear();
                    //int m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("E");
                    //moves[m].Add("Confirmed English");
                    int m = moves.Count;
                    moves.Add(new List<String>());
                    moves[m].Add("O");
                    moves[m].Add("Only Spanish");
                    //m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("pS");
                    //moves[m].Add("Only Spanish");
                    //m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("O");
                    //moves[m].Add("Other");
                    //m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("C5");
                    //moves[m].Add("Contacted 5 times letters");
                    m = moves.Count;
                    moves.Add(new List<String>());
                    moves[m].Add("I");
                    moves[m].Add("Field Service");
                    //m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("N/A");
                    //moves[m].Add("Field Service");
                    //m = moves.Count;
                    //moves.Add(new List<String>());
                    //moves[m].Add("skip");
                    //moves[m].Add("Field Service");

                    nRow = 0;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        int remainder;
                        Math.DivRem(row.Index, dataGridView1.Rows.Count, out remainder);
                        if (dataGridView1.Rows.Count > 100)
                            progress = (remainder / (dataGridView1.Rows.Count / 100));
                        else
                            progress = remainder * ((100 - 1) / dataGridView1.Rows.Count);
                        if (progress < 99) progressBar1.Value = progress;
                        progressBar1.Update();

                        if (row.Visible == false) continue;
                        if (o != row.Index)
                        {
                            o = row.Index;
                            int a = sSaveRow4Del.Count;
                            sSaveRow4Del.Add(string.Format("{0}", o));
                        }
                        int t = 0;
                        int f = row.Index;
                        if (o != row.Index)
                        {
                            o = row.Index;
                            int a = sSaveRow4Del.Count;
                            sSaveRow4Del.Add(string.Format("{0}", o));
                        }
                        if (string.IsNullOrEmpty(row.Cells[4].Value?.ToString()))
                            row.Cells[4].Value = "N/A";
                        if (row.Cells[4].Value.ToString() == "I") t = lastFSrow++;
                        if (row.Cells[4].Value.ToString() == "O") t = lastSProw++;
                        //if (row.Cells[4].Value.ToString() == "pS") t = lastSProw++;
                        //if (row.Cells[4].Value.ToString() == "I") t = f;        // we don't move or add these, just delete it!
                        //if (row.Cells[4].Value.ToString() == "N/A")             // see if N/A was attempted more than 5 times
                        {
                            short numtrys = 0;
                            t = lastC5row++;
                            if (string.IsNullOrEmpty(row.Cells[6].Value?.ToString())) row.Cells[6].Value = 0;
                            if (Int16.TryParse(row.Cells[6].Value?.ToString(), out numtrys))
                                if (numtrys > 5)
                                {
                                    row.Cells[4].Value = "C5";
                                    sSaveRow4Del.Remove(string.Format("{0}", o));
                                    sSaveRow4Del.Add("0");
                                }
                                else
                                    continue;
                        }

                        dataGridView1.Update();

                        int c = cellch.Count;                           // build list of changes
                        cellch.Add(new List<String>());
                        n = dataGridView1.CurrentCellAddress.X;         // save row for later
                        t = n + t + 1;                                      // calculate last row in target sheet

                        for (int w = 0; w < row.Cells.Count; w++)       // list every change location and data
                        {
                            string ts = string.Format("{{X={0},Y={1}}}", w, t);
                            cellch[c].Add(ts.ToString());
                            if (string.IsNullOrEmpty(row.Cells[w].Value?.ToString())) cellch[c].Add(tmp);
                            else cellch[c].Add(row.Cells[w].Value?.ToString());
                        }
                        textBox1.Text = string.Format("{0} changes", cellch.Count); // total the changes
                        textBox1.Update();


                        if (checkBox1.Checked == true) SaveSheetChanges("Field Service");
                        else continue;
                    }
                    Tabfocus = "Field Service";
                }
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                    if (dataGridView1.Rows[y].Visible == true)
                    {
                        dataGridView1.Rows[y].Cells[0].Selected = true;
                        break;
                    }
                checkBox2.Visible = true;
                checkBox3.Visible = true;
                progressBar1.Value = 100;
                progressBar1.Update();
                label9.Visible = true;
                button9.BackColor = System.Drawing.Color.LightGray;
                button4.BackColor = System.Drawing.Color.LightGreen;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
                checkBox2.Update();
                checkBox3.Update();
                dataGridView1.Visible = true;
                dataGridView1.Update();
                return;
            }
        }

        //
        // button10 - Prep Filter for EOD
        //
        private void button10_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                gEODhit = true;
                checkBox2.Checked = true;                                   // move and delete
                checkBox3.Checked = true;
                checkedListBox1.SetItemChecked(0, false);                  // "N/A"
                checkedListBox1.SetItemChecked(1, false);                  // "B"
                checkedListBox1.SetItemChecked(2, false);                  // "I"
                checkedListBox1.SetItemChecked(3, false);                  //" DNC"
                checkedListBox1.SetItemChecked(4, false);                  // "SP"
                checkedListBox1.SetItemChecked(5, false);                  // "E"
                checkedListBox1.SetItemChecked(6, true);                   //"blank"
                checkedListBox1.SetItemChecked(7, true);                   // "pE"
                checkedListBox1.SetItemChecked(8, true);                   // "pS"
            }
            if (comboBox1.SelectedIndex == 1)
            {
                gEODhit = true;
                checkBox2.Checked = true;                                   // move but dont delete
                checkBox3.Checked = false;
                checkedListBox1.SetItemChecked(0, true);                   // "N/A"
                checkedListBox1.SetItemChecked(1, false);                   // "B"
                checkedListBox1.SetItemChecked(2, true);                    // "I"
                checkedListBox1.SetItemChecked(3, false);                   //" DNC"
                checkedListBox1.SetItemChecked(4, true);                    // "SP"
                checkedListBox1.SetItemChecked(5, true);                    // "E"
                checkedListBox1.SetItemChecked(6, false);                   //"blank"
                checkedListBox1.SetItemChecked(7, false);                   // "pE"
                checkedListBox1.SetItemChecked(8, true);                    // "pS"
            }
            button2.BackColor = System.Drawing.Color.LightGreen;
        }

        //
        // suggest Next number
        //
        public decimal _attempt = 0;

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            pastTime = DateTimePicker.MinimumDateTime;
            string[] days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
            foreach (string dow in days)
                form3.checkedListBox2.Items.Add(dow);
            string whichM = DateTime.Now.ToString("tt");
            if (whichM == "AM") form3.radioButton1.Checked = true;
            else form3.radioButton2.Checked = true;

            int day1 = (int)DateTime.Now.DayOfWeek;
            form3.checkedListBox2.SetItemChecked(day1, true);

            button11.BackColor = System.Drawing.Color.LightGray;
            button11.Visible = false;
            button11.Update();

            if (!checkBox4.Checked)
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                    dataGridView1.Rows[y].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;

            if (_setAM) form3.radioButton1.Checked = true;
            else form3.radioButton1.Checked = false;
            if (_setPM) form3.radioButton2.Checked = true;
            else form3.radioButton2.Checked = false;

            //lst = new List<string>(form3.checkedListBox2.CheckedItems.Cast<string>());
            for (int count = 0; count < form3.checkedListBox2.Items.Count; count++)
            {
                if (lst.Contains(form3.checkedListBox2.Items[count].ToString()))
                    form3.checkedListBox2.SetItemChecked(count, true);
                else
                    form3.checkedListBox2.SetItemChecked(count, false);
            }
            if (_attempt == 0) _attempt = 5;
            if (lastDate == DateTime.MinValue)
            {
                lastDate = DateTime.Now;
                pastTime = DateTime.Now;
            }
            form3.monthCalendar1.SelectionStart = lastDate;
            form3.monthCalendar1.SelectionEnd = pastTime;
            form3.monthCalendar1.SetDate(lastDate);
            form3.numericUpDown1.Value = _attempt;

            if (checkBox4.Checked == true)
            {

                dr = form3.ShowDialog();            // bring up the form
            }

            if (DialogResult.Cancel == dr)
            {
                checkBox4.Checked = false;
                button11.Visible = false;
                dataGridView1.MultiSelect = false;

            }
            if (DialogResult.OK == dr)
            {
                if (form3.setMinstate) setMinDate = true;
                if (form3.radioButton1.Checked) _setAM = true;
                if (form3.radioButton2.Checked) _setPM = true;
                lst = new List<string>(form3.checkedListBox2.CheckedItems.Cast<string>());
                _attempt = form3.numericUpDown1.Value;
                if (setMinDate == false)
                {
                    pastTime = form3.monthCalendar1.SelectionStart;
                    lastDate = form3.monthCalendar1.SelectionEnd;
                }
                else
                {
                    pastTime = DateTime.MinValue;
                    lastDate = DateTime.MinValue;
                }

                button11.Visible = true;
                if (checkBox4.Checked) button11.BackColor = System.Drawing.Color.LightGreen;
                else button11.BackColor = System.Drawing.Color.LightGray;
                button11.Update();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            int sel_rows = 0;
            button11.Visible = true;
            bool _am = false;
            bool _pm = false;
            button11.BackColor = System.Drawing.Color.Coral;
            button11.Update();
            DateTime now = DateTime.Now;
            DateTime pastDate = pastTime;
            DateTime date = new DateTime();
            DayOfWeek whatDay = 0;
            int s = -1;

            dataGridView1.ClearSelection();
            cellch.Clear();
            DataChanged = false;
            button2.BackColor = System.Drawing.Color.LightGray;
            button4.BackColor = System.Drawing.Color.LightGray;
            startPB(System.Drawing.Color.Aqua);
            try
            {
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    int remainder;
                    Math.DivRem(y, dataGridView1.Rows.Count, out remainder);
                    if (dataGridView1.Rows.Count > 100)
                        progress = remainder / (dataGridView1.Rows.Count/100);
                    else
                        progress = remainder * ((100 - 1) / dataGridView1.Rows.Count);
                    if (progress < 100) progressBar1.Value = progress;
                    progressBar1.Update();

                    bool hasDate = false;
                    bool didMatch = false;
                    _am = false;
                    _pm = false;

                    if (dataGridView1.Rows[y].Visible == false)
                        continue;

                    string text = dataGridView1.Rows[y].Cells[5].Value?.ToString();
                    try
                    {
                        date = DateTime.Parse(text);
                        hasDate = true;
                        lastDate = date;
                        var isAorPm = date.ToString("tt");
                        if (isAorPm.ToString() == "AM") _am = true;
                        else _am = false;
                        if (isAorPm.ToString() == "PM") _pm = true;
                        else _pm = false;
                        whatDay = date.DayOfWeek;
                    }
                    catch (Exception ex) { }

                    if (hasDate && (setMinDate == false) || string.IsNullOrEmpty(dataGridView1.Rows[y].Cells[4].Value?.ToString()))
                    {
                        TimeSpan pastDays = now - pastTime;
                        TimeSpan diff = now - lastDate;
                        if (diff.Days > pastDays.Days || string.IsNullOrEmpty(dataGridView1.Rows[y].Cells[4].Value?.ToString()))
                            didMatch = true;
                        else
                            didMatch = false;
                    }
                    //else didMatch = true;

                    if ((_setAM || _setPM) && didMatch)
                    {
                        if ((_setAM && _pm) || (_setPM && _am))
                            didMatch = true;
                        else
                            didMatch = false;
                    }
                    //else didMatch = true;
                    if ((_attempt > 0) && didMatch)
                    {
                        if (!string.IsNullOrEmpty(dataGridView1.Rows[y].Cells[6].Value?.ToString()))
                        {
                            if (Convert.ToInt32(Convert.ToInt32(dataGridView1.Rows[y].Cells[6].Value.ToString())) > _attempt)
                                didMatch = false;
                            else
                                didMatch = true;
                        }
                    }

                    if ((whatDay != null) && didMatch)
                    {
                        bool found = false;

                        foreach (string wday in lst)
                        {
                            if (wday != whatDay.ToString()) found = true;
                            break;
                        }
                        if (found) didMatch = true;
                        else didMatch = false;
                    }
                    //else didMatch = true;

                    if (didMatch)
                    {
                        dataGridView1.Rows[y].DefaultCellStyle.BackColor = System.Drawing.Color.Aquamarine;
                        if (s < 1) s = y;
                        sel_rows++;
                    }

                }
            }
            catch { }
            progressBar1.Value = 100;

            DialogResult result = System.Windows.Forms.MessageBox.Show(string.Format(" there are {0} rows recommended. Do you want to accept and work these?", sel_rows),
                                                                 "Important Question", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                button11.BackColor = System.Drawing.Color.Aquamarine;
                button11.Update();
                int y;

                for (y = 0; y < dataGridView1.Rows.Count; y++)
                {
                    if (dataGridView1.Rows[y].DefaultCellStyle.BackColor != System.Drawing.Color.Aquamarine)
                    {
                        dataGridView1.Rows[y].Selected = true;
                        break;
                    }
                }
                dataGridView1.CurrentCell = dataGridView1[0, s];
                seeVisibleRow(dataGridView1, s);
            }
            if (result == DialogResult.No)
            {
                for (int y = 0; y < dataGridView1.Rows.Count; y++)
                    dataGridView1.Rows[y].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;
                button11.BackColor = System.Drawing.Color.LightGray;
                button11.Update();
                checkBox4.Checked = false;
            }

            dataGridView1.Update();
            return;
        }
        
        //
        // see VisableRow
        //
        private static void seeVisibleRow(DataGridView view, int rowToShow)
        {
            int i = 0;
            var countVisible = view.DisplayedRowCount(false);

            for (i = rowToShow; i > 0; --i)
                if (view.Rows[i].Visible) break;

            view.FirstDisplayedScrollingRowIndex =  i;
        }
        //private void SetLanguageDictionary()
        //{
        //    ResourceDictionary dict = new ResourceDictionary();
        //    switch (Thread.CurrentThread.CurrentCulture.ToString())
        //    {
        //        case "en-US":
        //            dict.Source = new Uri("..\\Resources\\StringResources.en-US.xaml",
        //                          UriKind.Relative);
        //            break;
        //        case "sp-EC":
        //            dict.Source = new Uri("..\\Resources\\StringResources.fr-CA.xaml",
        //                               UriKind.Relative);
        //            break;
        //        default:
        //            dict.Source = new Uri("..\\Resources\\StringResources.en-US.xaml",
        //                              UriKind.Relative);
        //            break;
        //    }
        //    //Resources.MergedDictionaries.Add(dict);
        //    MergedDictionaries.Add(dict);
        //}
        //
        private void btnCalculate_Click(object sender, EventArgs e)
        {
            //int noItem;
            //decimal priceItem, price;
            //noItem = Convert.ToInt16(txtNoItems.Text);
            //priceItem = Convert.ToDecimal(txtPriceItem.Text);
            //price = priceItem * noItem;
            //lblDisplay.Text = price.ToString("C", CultureInfo.CurrentUICulture); //the "C" is used to display the local currency based on the language and culture selected
        }

        //
        private void ChangeLanguage(string lang) //A function called to change the language
        {
            foreach (Control c in this.Controls)
            {
                ComponentResourceManager resources = new ComponentResourceManager(typeof(Form1));
                resources.ApplyResources(c, c.Name, new CultureInfo(lang));
            }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton6.Checked)
                    {

                    //ChangeLanguage to American English
                    language = "en-US";

                    // Save user choice in settings
                    Properties.Settings.Default.Language = "en-US";
                    Properties.Settings.Default.Save();

                    ChangeLanguage(language);
                    }
           }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
                if(radioButton5.Checked)
                {
                    //ChangeLanguage to Ecuadorian Spanish
                    language = "es-EC";

                    // Save user choice in settings
                    Properties.Settings.Default.Language = "es-EC";
                    Properties.Settings.Default.Save();

                    ChangeLanguage(language);
                }
            }
    }
}





