namespace WFAGoolgeSheet
{
    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Services;
    using Google.Apis.Sheets.v4;
    using Google.Apis.Sheets.v4.Data;
    using Google.Apis.Util.Store;
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Windows;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="Form1" />.
    /// </summary>
    public partial class Form1 : Form
    {
        /// <summary>
        /// Defines the language.
        /// </summary>
        public string language = Properties.Settings.Default.Language;

        /// <summary>
        /// Gets or sets the dataGridView.
        /// </summary>
        public DataGridView dataGridView { get; set; }

        //private  FormMediator _formMediator;
        /// <summary>
        /// Gets the MergedDictionaries.
        /// </summary>
        public IList<ResourceDictionary> MergedDictionaries { get; }

        /// <summary>
        /// Gets or sets the Source.
        /// </summary>
        public System.Uri Source { get; set; }

        /// <summary>
        /// Defines the <see cref="sTabName" />.
        /// </summary>
        public class sTabName
        {
            /// <summary>
            /// Gets or sets the langName.
            /// </summary>
            public string langName { get; set; }

            /// <summary>
            /// Gets or sets the tabname.
            /// </summary>
            public string tabname { get; set; }

            /// <summary>
            /// Gets or sets the range.
            /// </summary>
            public string range { get; set; }

            /// <summary>
            /// Gets or sets the keyfield.
            /// </summary>
            public string keyfield { get; set; }

            /// <summary>
            /// The ToString.
            /// </summary>
            /// <returns>The <see cref="string"/>.</returns>
            public override string ToString()
            {
                return this.langName;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// </summary>
        public Form1()
        {
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(language);

            InitializeComponent();

            this.Font = System.Drawing.SystemFonts.IconTitleFont;
            SystemEvents.UserPreferenceChanged += new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);


            BuildComboBoxItems(1);
        }

        /// <summary>
        /// Gets the myDG.
        /// </summary>
        public DataGridView myDG
        {
            get { return dataGridView1; }
        }

        /// <summary>
        /// Gets the Resources.
        /// </summary>
        public object Resources { get; private set; }

        //---------------------------------------------------------------------------//
        // global flags and data
        //
        /// <summary>
        /// Defines the updateinprogress.
        /// </summary>
        internal bool updateinprogress = false;

        /// <summary>
        /// Defines the DataChanged.
        /// </summary>
        internal bool DataChanged = false;

        /// <summary>
        /// Defines the isProcessRunning.
        /// </summary>
        internal bool isProcessRunning = false;

        /// <summary>
        /// Defines the Tabfocus.
        /// </summary>
        internal string Tabfocus = null;

        /// <summary>
        /// Defines the gEODhit.
        /// </summary>
        internal bool gEODhit = false;

        /// <summary>
        /// Defines the rowOffset.
        /// </summary>
        internal int rowOffset = 0;

        /// <summary>
        /// Defines the firstrow.
        /// </summary>
        internal int firstrow = 0;

        /// <summary>
        /// Defines the skiprow.
        /// </summary>
        internal int skiprow = 0;

        /// <summary>
        /// Defines the lastFSrow.
        /// </summary>
        internal int lastFSrow = 0;

        /// <summary>
        /// Defines the lastSProw.
        /// </summary>
        internal int lastSProw = 0;

        /// <summary>
        /// Defines the lastCErow.
        /// </summary>
        internal int lastCErow = 0;

        /// <summary>
        /// Defines the lastC5row.
        /// </summary>
        internal int lastC5row = 0;

        /// <summary>
        /// Defines the progress.
        /// </summary>
        internal int progress = 0;

        /// <summary>
        /// Defines the rcount.
        /// </summary>
        internal int rcount = -1;

        /// <summary>
        /// Defines the chgCount.
        /// </summary>
        internal int chgCount = 0;

        /// <summary>
        /// Defines the foundCnt.
        /// </summary>
        internal int foundCnt = 0;

        /// <summary>
        /// Defines the firstFound.
        /// </summary>
        internal int firstFound = 0;

        /// <summary>
        /// Defines the curFound.
        /// </summary>
        internal int curFound = 0;

        /// <summary>
        /// Defines the myVar.
        /// </summary>
        public static string myVar = null;// passing data between forms

        /// <summary>
        /// Defines the lockMethod2.
        /// </summary>
        public static object lockMethod2 = new object();

        /// <summary>
        /// Defines the dr.
        /// </summary>
        public DialogResult dr = new DialogResult();

        /// <summary>
        /// Defines the sSaveRow4Del.
        /// </summary>
        internal List<String> sSaveRow4Del = new List<String>();

        /// <summary>
        /// Defines the notfoundlat.
        /// </summary>
        public float notfoundlat = float.Parse(" -1.4487420320510864");

        /// <summary>
        /// Defines the notfoundlon.
        /// </summary>
        public float notfoundlon = float.Parse("-78.376708984375");

        /// <summary>
        /// Defines the done.
        /// </summary>
        internal bool done = true;

        /// <summary>
        /// Defines the totalRow.
        /// </summary>
        internal int totalRow = 0;

        /// <summary>
        /// Defines the totalChgs.
        /// </summary>
        internal int totalChgs = 0;

        /// <summary>
        /// Defines the numOfSP.
        /// </summary>
        internal int numOfSP = 0;

        /// <summary>
        /// Defines the numOfpE.
        /// </summary>
        internal int numOfpE = 0;

        /// <summary>
        /// Defines the numOfSkip.
        /// </summary>
        internal int numOfSkip = 0;

        /// <summary>
        /// Defines the adder.
        /// </summary>
        internal int adder = 0;

        /// <summary>
        /// Defines the setMinDate.
        /// </summary>
        public bool setMinDate = false;

        /// <summary>
        /// Defines the _setPM.
        /// </summary>
        public bool _setPM = false;

        /// <summary>
        /// Defines the _setAM.
        /// </summary>
        public bool _setAM = false;

        /// <summary>
        /// Defines the checkedRadio.
        /// </summary>
        public string checkedRadio = "";

        /// <summary>
        /// Defines the resul.
        /// </summary>
        public string resul = "";

        /// <summary>
        /// Defines the attempt.
        /// </summary>
        public string attempt = "";

        /// <summary>
        /// Defines the notes.
        /// </summary>
        public string notes = "";

        /// <summary>
        /// Defines the adjGPS.
        /// </summary>
        public bool adjGPS = false;

        /// <summary>
        /// Defines the confid.
        /// </summary>
        public string confid = "";

        /// <summary>
        /// Defines the lat.
        /// </summary>
        public string lat = "";

        /// <summary>
        /// Defines the lon.
        /// </summary>
        public string lon = "";

        /// <summary>
        /// Defines the lastDate.
        /// </summary>
        internal DateTime lastDate = new DateTime();

        /// <summary>
        /// Defines the pastTime.
        /// </summary>
        internal DateTime pastTime = new DateTime();

        /// <summary>
        /// Defines the lst.
        /// </summary>
        internal List<string> lst = new List<string>();// for days of the week

        /// <summary>
        /// Defines the SecondFromTop.
        /// </summary>
        public int SecondFromTop;

        /// <summary>
        /// Defines the SecondFormLeft.
        /// </summary>
        public int SecondFormLeft;

        /// <summary>
        /// Defines the SecondFormSize.
        /// </summary>
        public System.Drawing.Size SecondFormSize = System.Drawing.Size.Empty;

        /// <summary>
        /// Defines the ForthFromTop.
        /// </summary>
        public int ForthFromTop;

        /// <summary>
        /// Defines the ForthFormLeft.
        /// </summary>
        public int ForthFormLeft;

        /// <summary>
        /// Defines the ForthFormSize.
        /// </summary>
        public System.Drawing.Size ForthFormSize = System.Drawing.Size.Empty;

        /// <summary>
        /// Defines the MapFormTop.
        /// </summary>
        public int MapFormTop;

        /// <summary>
        /// Defines the MapFormLeft.
        /// </summary>
        public int MapFormLeft;

        /// <summary>
        /// Defines the MapFormSize.
        /// </summary>
        public System.Drawing.Size MapFormSize;

        /// <summary>
        /// Defines the vs.
        /// </summary>
        internal string[] vs = null;

        /// <summary>
        /// Defines the ttemp.
        /// </summary>
        internal static string ttemp = "";

        /// <summary>
        /// Defines the trans.
        /// </summary>
        internal static string trans = "";

        //
        // list for sheets row moves
        //
        /// <summary>
        /// Defines the moves.
        /// </summary>
        internal List<List<String>> moves = new List<List<String>>();

        //
        // list of all datagridview changes
        //
        /// <summary>
        /// Defines the cellch.
        /// </summary>
        public List<List<String>> cellch = new List<List<String>>();//Creates new nested List

        //
        // global phone number list
        //
        /// <summary>
        /// Defines the values.
        /// </summary>
        internal IList<IList<object>> values = new List<IList<object>>();

        //
        // Google parameters
        /// <summary>
        /// Defines the credential.
        /// </summary>
        public UserCredential credential;// Sheet Service

        /// <summary>
        /// Defines the ApplicationName.
        /// </summary>
        public string ApplicationName = null;

        /// <summary>
        /// Defines the spreadsheetId.
        /// </summary>
        public String spreadsheetId = null;// Spreadsheet ID

        /// <summary>
        /// The BuildComboBoxItems.
        /// </summary>
        /// <param name="defIndex">The defIndex<see cref="Int32"/>.</param>
        public void BuildComboBoxItems(Int32 defIndex)
        {
            comboBox1.Items.Clear();
            if (radioButton6.Checked)
            {
                comboBox1.Items.Add(new sTabName { langName = "Imported Names", tabname = "Imported Names", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Field Service", tabname = "Field Service", range = "!A1:K", keyfield = "Field Service" });
                comboBox1.Items.Add(new sTabName { langName = "Confirmed Deaf", tabname = "Confirmed Deaf", range = "!A1:K", keyfield = "Telephone" });
                comboBox1.Items.Add(new sTabName { langName = "Contacted 5 times letters", tabname = "Contacted 5 times letters", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "NoGPSMap", tabname = "NoGPSMap", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Out of Territory", tabname = "Out Terr", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "GeoFence Data", tabname = "GeoFence Data", range = "!A1:D", keyfield = "territory" });
                comboBox1.Items.Add(new sTabName { langName = "Names of Deaf Lost", tabname = "Names of Deaf Lost", range = "!A1:D", keyfield = "TELEPHONE" });
            }
            else
            {
                comboBox1.Items.Add(new sTabName { langName = "Nombres importados", tabname = "Imported Names", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Servicio de campo", tabname = "Field Service", range = "!A1:K", keyfield = "Field Service" });
                comboBox1.Items.Add(new sTabName { langName = "Sordos confirmados", tabname = "Confirmed Deaf", range = "!A1:K", keyfield = "Telephone" });
                comboBox1.Items.Add(new sTabName { langName = "Contactado 5 veces cartas", tabname = "Contacted 5 times letters", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Sin mapa GPS", tabname = "NoGPSMap", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Fuera territorio", tabname = "Out Terr", range = "!A1:K", keyfield = "TELEPHONE" });
                comboBox1.Items.Add(new sTabName { langName = "Datos de GeoFence", tabname = "GeoFence Data", range = "!A1:D", keyfield = "territory" });
                comboBox1.Items.Add(new sTabName { langName = "Nombres de sordos perdidos", tabname = "Names of Deaf Lost", range = "!A1:D", keyfield = "TELEPHONE" });
                //comboBox1.Items.Add(new sTabName { tabname = "Common First Names", range = "!A1:G", keyfield = "TELEPHONE" });
            }
            comboBox1.SelectedIndex = defIndex;
        }

        /// <summary>
        /// The SystemEvents_UserPreferenceChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="UserPreferenceChangedEventArgs"/>.</param>
        internal void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Window)
            {
                this.Font = System.Drawing.SystemFonts.IconTitleFont;
            }
        }

        /// <summary>
        /// The Form1_FormClosing.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="FormClosingEventArgs"/>.</param>
        internal void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
        }

        //---------------------------------------------------------------------------
        //
        // close application and exit
        //
        /// <summary>
        /// The button1_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = DialogResult.No;
                if (radioButton6.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                if (radioButton5.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show("Hay cambios no guardados \n ¿Desea salir y perder estos cambios?",
                    "Preguntas importantes",
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
        /// <summary>
        /// The button2_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button2_Click(object sender, EventArgs e)
        {
            GPSgeofence gPSgeofence = new GPSgeofence();
            GPSgeofence fence = gPSgeofence;
            fence.ReadGPSfence();
            textBox11.Text = string.Format("GeoFence {0} - {1} pts.", fence.fenceName, fence.polyCorners);
            textBox11.Update();

            int NumofRec = 0;
            int r1 = -1;
            if (string.IsNullOrEmpty(textBox3.Text)) firstrow = 0;  // setup x of n grouping
            else firstrow = Convert.ToInt32(textBox3.Text);
            if (string.IsNullOrEmpty(textBox5.Text)) skiprow = 0;
            else skiprow = Convert.ToInt32(textBox5.Text);
            firstrow = skiprow + 1 - firstrow;                       // invert it!
            fence.ReadGPSfence();
            if (checkedListBox1.CheckedItems.Count < 1 && comboBox1.SelectedIndex < 2)
            {
                if (radioButton6.Checked)
                    System.Windows.Forms.MessageBox.Show(" You have no Filters selected");
                if (radioButton5.Checked)
                    System.Windows.Forms.MessageBox.Show(" No tiene filtros seleccionados");
                return;
            }
            if (comboBox1.SelectedIndex == -1)
            {
                if (radioButton6.Checked)
                    System.Windows.Forms.MessageBox.Show(" No Sheet Tab selected");
                if (radioButton5.Checked)
                    System.Windows.Forms.MessageBox.Show("No se ha seleccionado la ficha de la hoja");
                return;
            }

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                if (radioButton6.Checked)
                    System.Windows.Forms.MessageBox.Show(" No Sheet [test/live] selected");
                if (radioButton5.Checked)
                    System.Windows.Forms.MessageBox.Show(" No se ha seleccionado la prueba de la hoja / en vivo");
                return;
            }
            if (comboBox1.SelectedIndex == 1)
            {
                if (firstrow == 0 || skiprow == 0 || firstrow > skiprow)
                {
                    if (radioButton6.Checked)
                        System.Windows.Forms.MessageBox.Show(" group range is bad ");
                    if (radioButton5.Checked)
                        System.Windows.Forms.MessageBox.Show(" el rango del grupo es malo ");
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
                    if (radioButton6.Checked)
                        System.Windows.Forms.MessageBox.Show(" no filters selected ");
                    if (radioButton5.Checked)
                        System.Windows.Forms.MessageBox.Show(" no hay filtros seleccionados");
                    return;
                }
            }
            if (DataChanged)
            {
                DialogResult result1 = DialogResult.No;
                if (radioButton6.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                if (radioButton5.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show("Hay cambios no guardados \n ¿Desea salir y perder estos cambios?",
                    "Preguntas importantes",
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
                // calculate RGB color values
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
                        if (radioButton6.Checked)
                            textBox1.Text = "adding .. " + NumofRec.ToString();
                        if (radioButton5.Checked)
                            textBox1.Text = "añadir .. " + NumofRec.ToString();
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
                            // Light Orange (~12) = Exists in "Contacted 5 times letters"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#fffbbc04"))
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Contacted 5 times letters\"";

                            // Pink   = Exists in "Only Spanish"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ffea9999"))
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Out Terr\"";

                            // Red (~1)   = Exists in "This Sheet"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ffea4335"))
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"This Sheet\"";

                            // Purple = Exists in "Confirmed Deaf"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ff8e7cc3"))
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Confirmed Deaf\"";

                            // Dark Orange (Gold) (~15) = Exists in "Field Service"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ffff6d01"))
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Field Service\"";

                            // Dark Grey (30)  = Exists in "Imported Names"
                            if (ccolor == System.Drawing.ColorTranslator.FromHtml("#ff7f6000"))
                            {
                                dataGridView1.Rows[x].Cells[0].ToolTipText = "Exists in \"Imported Names\"";
                                ccolor = System.Drawing.Color.SandyBrown; ;             // this RGB color was too dark
                            }

                            // handle green bar option in Google Sheets and change to white
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
                    dataGridView1.Rows[y].Cells[1].Selected = true;
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
        /// <summary>
        /// The button3_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button3_Click(object sender, EventArgs e)
        {

            if (DataChanged)
            {
                DialogResult result1 = DialogResult.No;
                if (radioButton6.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                if (radioButton5.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show("Hay cambios no guardados \n ¿Desea salir y perder estos cambios?",
                    "Preguntas importantes",
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
            checkedListBox1.Visible = true;
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
        /// <summary>
        /// The button5_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button5_Click(object sender, EventArgs e)
        {
            if (DataChanged)
            {
                DialogResult result1 = DialogResult.No;
                if (radioButton6.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show(" There are unsaved changes \n Do you want to exit and loose these changes?",
                    "Important Question",
                    MessageBoxButtons.YesNo);
                if (radioButton5.Checked)
                    result1 = System.Windows.Forms.MessageBox.Show("Hay cambios no guardados \n ¿Desea salir y perder estos cambios?",
                    "Preguntas importantes",
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
        /// <summary>
        /// The DataGridView1_Cell_Mouse_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewCellEventArgs"/>.</param>
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

                    //string cellResult = null;
                    //string cellNote = null;

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
                            if (intIndex + 1 > dataGridView1.RowCount)
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
                                form2.textBox9.Text = dataGridView1.CurrentRow.Index.ToString();

                                form2.textBox7.Text = dataGridView1.CurrentRow.Cells[3].Value?.ToString();
                                string tmp = "";
                                string tmp0 = "";
                                if (string.IsNullOrEmpty(form2.textBox1.Text)) tmp = " no phone ";
                                else
                                {
                                    ToolTip myToolTip = new ToolTip();
                                    tmp = form2.textBox1.Text;
                                    form2.textBox1.BackColor = dataGridView1.Rows[nRow].Cells[0].Style.BackColor;
                                    myToolTip.SetToolTip(form2.textBox1, dataGridView1.CurrentRow.Cells[0].ToolTipText);
                                }
                                form2.textBox1.Update();
                                if ((dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value?.ToString() == "g") ||
                                    (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value?.ToString() == "H"))
                                {
                                    form2.textBox10.Text = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Latitude")].Value?.ToString();
                                    form2.textBox11.Text = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Longitude")].Value?.ToString();
                                    form2.textBox10.Update();
                                    form2.textBox11.Update();
                                }

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
                                    Translator translation = new Translator();
                                    for (int j = 0; j < lines.Count(); j++)
                                    {
                                        ttemp = form2.textBox6.Lines[j];
                                        firstpos = ttemp.Contains(today.ToString("yyyy-MM-dd"));
                                        if (firstpos) break;
                                    }
                                    trans = translation.TranslationAsync(form2.textBox6.Text).ToString();    //new - translation added

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
        /// <summary>
        /// The saveFSdata.
        /// </summary>
        /// <param name="nRow">The nRow<see cref="int"/>.</param>
        private void saveFSdata(int nRow)
        {
            Form2 form2 = new Form2();
            string cellResult = "";

            dataGridView1.Rows[nRow].Selected = true;
            dataGridView1.Rows[nRow].Cells[0].Selected = true;

            //
            // process RadioButtons on Form2
            //
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

            if (adjGPS)
            {
                dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Selected = true;
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
        /// <summary>
        /// The dataGridView1_ColumnHeaderMouseClick.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewCellMouseEventArgs"/>.</param>
        private void dataGridView1_ColumnHeaderMouseClick(
    object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn newColumn = dataGridView1.Columns[e.ColumnIndex];
        }

        //---------------------------------------------------------------------------
        //
        // Cell Changes are tracked  and exit is prevented until saved
        //
        /// <summary>
        /// The DataGridView1_CellValueChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewCellEventArgs"/>.</param>
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
        /// <summary>
        /// The DataGridView1_RowsRemoved.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewRowsRemovedEventArgs"/>.</param>
        private void DataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

            UpdateSheet();
        }

        //---------------------------------------------------------------------------
        //
        // user row deletion is not permitted
        //
        /// <summary>
        /// The DataGridView1_UserDeletingRow.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewRowCancelEventArgs"/>.</param>
        private void DataGridView1_UserDeletingRow(object sender,
    DataGridViewRowCancelEventArgs e)
        {
            if (radioButton6.Checked)
                System.Windows.Forms.MessageBox.Show("Cannot delete a record!");
            if (radioButton5.Checked)
                System.Windows.Forms.MessageBox.Show("No se puede eliminar un registro.");
            e.Cancel = true;
        }

        //---------------------------------------------------------------------------
        //
        // UpdateSheet - writes back all changed to the Sheet by calling SaveSheetChanges
        //
        /// <summary>
        /// The UpdateSheet.
        /// </summary>
        private void UpdateSheet()
        {
            if (!updateinprogress && checkBox1.Checked)
            {
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
        /// <summary>
        /// Defines the begTimInc.
        /// </summary>
        internal DateTime begTimInc = DateTime.MinValue;

        /// <summary>
        /// Defines the nowTimInc.
        /// </summary>
        internal DateTime nowTimInc = DateTime.MinValue;

        /// <summary>
        /// Defines the maxChgs.
        /// </summary>
        internal int maxChgs = Properties.Settings.Default.maxChg;

        /// <summary>
        /// Defines the perSec.
        /// </summary>
        internal int perSec = Properties.Settings.Default.perSec;

        /// <summary>
        /// Defines the rows2del.
        /// </summary>
        internal int rows2del = 0;

        /// <summary>
        /// The SaveSheetChanges.
        /// </summary>
        /// <param name="Tabname">The Tabname<see cref="string"/>.</param>
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
            if (!checkBox1.Checked)   // dont show on auto save
            {
                progressBar1.Value = 2;
                progressBar1.Update();
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //
            // initialize process variables
            //
            if (!checkBox1.Checked)   // dont show on auto save
            {
                progressBar1.Value = 4;
                progressBar1.Update();
            }

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
            //int index2 = 0;
            //var dPoint = "";
            //int dpCount = -1;
            //string DataRow = "";

            string temp0 = "";
            string temp1 = "";
            string temp2 = "";
            string temp3 = "";
            string temp4 = "";
            bool toggle = false;
            bool modstr = false;
            bool conseq = false;

            //Copy items from list1 to list2
            //
            List<List<String>> cellch2 = new List<List<String>>();

            cellch2.Clear();

            foreach (var dPoint in cellch)
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
                    if (oldc != -1) temp2 = temp2 + "," + temp1;
                    oldc = newc;
                    modstr = true;
                }
                else
                {
                    temp3 = temp4;
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

                if (!String.IsNullOrEmpty(oRow))                // increment to next row
                {
                    int number = Convert.ToInt32(oRow);
                    number = number + rowOffset + 1;
                    sRow = number.ToString();
                }
                string stmp = "";
                if (Tabname == "Field Service") stmp = "FS";
                //if (Tabname == "Out Terr") stmp = "OT";
                if (Tabname == "Imported Names") stmp = "IM";
                if (Tabname == "Out Terr") stmp = "O";
                if (Tabname == "Contacted 5 times letters") stmp = "5X";
                if (Tabname == "Confirmed Deaf") stmp = "CD";
                textBox2.Text = string.Format(" " + stmp + " row {0}", sRow);
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
                if (!checkBox1.Checked)
                {
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
            }

            //
            // Clean up and finish
            //
            updateinprogress = false;
            DataChanged = false;
            cellch2.Clear();
            if (radioButton6.Checked)
                textBox1.Text = " working ";
            if (radioButton5.Checked)
                textBox1.Text = " trabajar ";
            textBox1.Update();
            if (!(checkBox1.Checked && comboBox1.SelectedIndex == 0))
            {
                if (radioButton6.Checked)
                    textBox1.Text = " done ";
                if (radioButton5.Checked)
                    textBox1.Text = " hecho ";
                textBox1.Update();
                if (!checkBox1.Checked)
                    progressBar1.Value = 100;
                button4.BackColor = System.Drawing.Color.LightGray;
            }

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
        /// <summary>
        /// The button4_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
                SaveSheetChanges(Tabfocus);
            else
                SaveSheetChanges(null);
            cellch.Clear();
        }

        //
        // progress bar
        //
        /// <summary>
        /// The startPB.
        /// </summary>
        /// <param name="color">The color<see cref="System.Drawing.Color"/>.</param>
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

        //
        //
        //
        /// <summary>
        /// The DataGridView1_CellEndEdit.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewCellEventArgs"/>.</param>
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

        /// <summary>
        /// The dataGridView1_SelectionChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="DataGridViewRowEventArgs"/>.</param>
        private void dataGridView1_SelectionChanged(object sender, DataGridViewRowEventArgs e)
        {
            textBox2.Text = e.Row.ToString();
            textBox2.Update();
        }

        /// <summary>
        /// The listBox1_SelectedIndexChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="System.EventArgs"/>.</param>
        private void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            button2.BackColor = System.Drawing.Color.LightGreen;
        }

        /// <summary>
        /// The Textbox_DoulbClickEvent.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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

        /// <summary>
        /// The button6_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button6_Click(object sender, EventArgs e)
        {
        }

        //
        // sheet tab selection
        //
        /// <summary>
        /// The combox1_selectedIndexChange.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
                textBox9.Visible = false;
                textBox10.Visible = false;
                label10.Visible = false;
                checkedListBox1.CheckOnClick = true;
                checkedListBox1.SetItemChecked(6, true);                  //"blank";
                checkBox1.Checked = true;
                checkedListBox1.Items.Add("Ignore Dups", false);
            }

            if (comboBox1.SelectedIndex == 1)           // Field Service
            {
                checkedListBox1.Items.Remove("Ignore Dups");
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
                checkedListBox1.Items.Remove("Ignore Dups");
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                checkBox2.Visible = false;                  // move
                checkBox3.Visible = false;                  // delete
                checkBox4.Visible = false;
                button11.Visible = false;                   // suggest next
            }

            if (comboBox1.SelectedIndex == 3)
            {
                checkedListBox1.Items.Remove("Ignore Dups");
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                checkBox2.Visible = false;                  // move
                checkBox3.Visible = false;                  // delete
                checkBox4.Visible = false;
                button11.Visible = false;                   // suggest next
            }

            if (comboBox1.SelectedIndex == 4)
            {
                checkedListBox1.Items.Remove("Ignore Dups");
                button9.Visible = false;
                label1.Visible = false;
                checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                checkBox2.Visible = false;                  // move
                checkBox3.Visible = false;                  // delete
                checkBox4.Visible = false;
                button11.Visible = false;                   // suggest next
            }

            if ((comboBox1.SelectedIndex == 5) ||
                (comboBox1.SelectedIndex == 6) ||
                (comboBox1.SelectedIndex == 7))
            {
                checkedListBox1.Items.Remove("Ignore Dups");
                button9.Visible = false;
                label1.Visible = false;
                //checkBox4.Visible = false;
                checkedListBox1.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label5.Visible = false;
                label4.Visible = false;
                checkBox2.Visible = false;                  // move
                checkBox3.Visible = false;                  // delete
                checkBox4.Visible = false;
                button11.Visible = false;                   // suggest next
            }
            button2.BackColor = System.Drawing.Color.LightGreen;
        }

        //
        // update current row textBox
        //
        /// <summary>
        /// Defines the curRow.
        /// </summary>
        internal int curRow = -1;

        /// <summary>
        /// Defines the curRowUpdate.
        /// </summary>
        internal int curRowUpdate = -1;

        /// <summary>
        /// Defines the txtConsole.
        /// </summary>
        private Control txtConsole;

        /// <summary>
        /// Defines the visibleColumnIndex.
        /// </summary>
        private readonly int visibleColumnIndex;

        /// <summary>
        /// The dataGridView1_SelectionChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
        //
        // hit the search button on searchbox enter key
        //
        /// <summary>
        /// The TextBox4_GotFocus.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        internal void TextBox4_GotFocus(object sender, EventArgs e)
        {
            this.AcceptButton = button6;
        }

        //
        // Prev. Search result
        //
        /// <summary>
        /// The button13_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
        /// <summary>
        /// The button12_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
        /// <summary>
        /// The GetindexOf.
        /// </summary>
        /// <param name="dgv">The dgv<see cref="DataGridView"/>.</param>
        /// <param name="name">The name<see cref="string"/>.</param>
        /// <returns>The <see cref="int"/>.</returns>
        public int GetindexOf(DataGridView dgv, string name)
        {
            int colInd = dgv.Columns[name].Index;
            return (colInd);
        }

        /// <summary>
        /// The Form1_Load.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void Form1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Language == "en-US")
                radioButton6.Checked = true;
            else
                radioButton5.Checked = true;
        }

        //
        // search datagrid for value
        //
        /// <summary>
        /// The button6_Click_1.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button6_Click_1(object sender, EventArgs e)
        {
            var rSelected = new List<DataGridViewRow>();
            string searchValue = "";
            searchValue = textBox4.Text;
            curFound = 0;
            foundCnt = 0;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;

            dataGridView1.ClearSelection();
            dataGridView1.Update();
            dataGridView1.CurrentRow.Selected = false;

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
                if (radioButton6.Checked)
                    System.Windows.Forms.MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));
                if (radioButton5.Checked)
                    System.Windows.Forms.MessageBox.Show(string.Format(" Búsqueda encontrada {0} artículos", Convert.ToString(foundCnt)));

                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    rSelected.Add(r);
                dataGridView1.CurrentCell = dataGridView1[visibleColumnIndex, firstFound];
                dataGridView1.CurrentRow.Selected = true;
                seeVisibleRow(dataGridView1, firstFound);
                foreach (DataGridViewRow rs in rSelected)
                {
                    dataGridView1.Rows[rs.Index].Selected = true;
                }
            }
            catch (Exception ex)
            {
                if (radioButton6.Checked)
                    System.Windows.Forms.MessageBox.Show(string.Format(" Search found {0} items", Convert.ToString(foundCnt)));
                if (radioButton5.Checked)
                    System.Windows.Forms.MessageBox.Show(string.Format(" Búsqueda encontrada {0} artículos", Convert.ToString(foundCnt)));
            }
        }

        //
        // Clear search box and Selected Rows
        //
        /// <summary>
        /// The button7_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
        /// <summary>
        /// The SetResultTime.
        /// </summary>
        /// <param name="result">The result<see cref="string"/>.</param>
        public void SetResultTime(string result)
        {
            DateTime today = DateTime.Today;
            int c;

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Selected = true;
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = result;
            c = cellch.Count;
            cellch.Add(new List<String>());
            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
            cellch[c].Add(result);

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value = today.ToString("yyyy-MM-dd");
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Selected = true;
            c = cellch.Count;
            cellch.Add(new List<String>());
            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
            cellch[c].Add(today.ToString("yyyy-MM-dd"));

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Value = "0";
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "#Attempts")].Selected = true;
            c = cellch.Count;
            cellch.Add(new List<String>());
            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
            cellch[c].Add("0");

            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Value = " ";
            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NOTES")].Selected = true;
            c = cellch.Count;
            cellch.Add(new List<String>());
            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
            cellch[c].Add(" ");
        }

        //
        // remove duplicate words in string
        //
        /// <summary>
        /// The RemoveAnyDuplicates.
        /// </summary>
        /// <param name="SentanceString">The SentanceString<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public string RemoveAnyDuplicates(string SentanceString)
        {
            string withoutDuplicates = String.Join("+", SentanceString.Split('+').Distinct());
            return (withoutDuplicates);
        }

        //
        // remove duplicate words in string
        //
        /// <summary>
        /// The RemoveConsecutive.
        /// </summary>
        /// <param name="SentenceString">The SentenceString<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public string RemoveConsecutive(string SentenceString)
        {
            //string SetenceString = "red white black white green yellow red red black white";
            var v = new HashSet<string>();
            string[] data = SentenceString.Split('+');
            string retStr = "";
            HashSet<string> set = new HashSet<string>();
            for (int i = 0; i < data.Length; i++)
                set.Add(data[i]);

            string adjset = string.Join("+", set);
            for (int i = 0; i < (SentenceString.Length - adjset.Length); i++)
                retStr = retStr + SentenceString[i];

            int indexOfSubString = adjset.IndexOf(retStr);
            //remove specified substring from string
            if (indexOfSubString > 0)
                adjset = adjset.Remove(indexOfSubString, retStr.Length);
            //}
            return (adjset);
        }

        /// <summary>
        /// Defines the <see cref="Coordinates" />.
        /// </summary>
        public class Coordinates
        {
            /// <summary>
            /// Gets or sets the Longitude.
            /// </summary>
            public double Longitude { get; set; }

            /// <summary>
            /// Gets or sets the Latitude.
            /// </summary>
            public double Latitude { get; set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="Coordinates"/> class.
            /// </summary>
            /// <param name="Long">The Long<see cref="double"/>.</param>
            /// <param name="Lat">The Lat<see cref="double"/>.</param>
            public Coordinates(double Long, double Lat)
            {
                this.Longitude = Long;
                this.Latitude = Lat;
            }
        }

        /// <summary>
        /// The ExtractCoordinates.
        /// </summary>
        /// <param name="Coord">The Coord<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        private string ExtractCoordinates(string Coord)
        {
            List<Coordinates> lstOfCoordinates = new List<Coordinates>();

            Regex reg = new Regex(@"[-+]?[0-9]*\.[0-9]+.");

            //I get the matches and save them in a list
            MatchCollection collection = reg.Matches(Coord);
            List<string> lstOfMatches = (from Match match in collection
                                         select match.Value).ToList();

            string sval = "";
            int scnt = 0;
            for (int z = lstOfMatches.Count; z > 0; z--)
            {
                if (string.IsNullOrEmpty(sval))
                    sval = sval + lstOfMatches[z - 1];
                else
                    sval = sval + "," + lstOfMatches[z - 1];
                if (scnt++ > 1) break;
            }

            if (lstOfMatches.Count < 2) return (null);
            else return (sval);
        }

        //
        //
        // RunImported Names
        //
        /// <summary>
        /// Defines the names2chk.
        /// </summary>
        internal List<List<String>> names2chk = new List<List<String>>();

        /// <summary>
        /// Defines the <see cref="MyListBoxItem" />.
        /// </summary>
        public class MyListBoxItem
        {
            /// <summary>
            /// Gets or sets the Text.
            /// </summary>
            public string Text { get; set; }

            /// <summary>
            /// Gets or sets the Value.
            /// </summary>
            public string Value { get; set; }
        }

        /// <summary>
        /// The RunImportedNames.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void RunImportedNames(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            int k = 0;
            int nLim = 0;
            GPSgeofence gPSgeofence = new GPSgeofence();
            GPSgeofence fence = gPSgeofence;
            bool haveGPS = false;
            float lx = 0;
            float ly = 0;
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

                textBox1.Text = ".. reading data";
                textBox1.Update();

                DataChanged = true;
                int numOfSP = 0;
                textBox7.Clear();
                int numOfEN = 0;
                textBox6.Clear();
                int numOfSkip = 0;
                textBox8.Clear();
                //IList<IList<Object>> values = null;

                int i;
                bool found = false;

                progressBar1.Value = 4;
                progressBar1.Update();
                fence.ReadGPSfence();
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
                while ((nLim >= dataGridView1.RowCount || dataGridView1.Rows[nLim].Visible == false) && nLim > 0)       // dont end on an invisable row
                    nLim--;
                textBox9.Text = Convert.ToString(nLim);
                textBox9.Update();
            LB1: while (nRow < nLim + 1 /*+nOff*/)        //dataGridView1.RowCount)
                {
                    if (dataGridView1.Rows[nRow].Visible == false)
                    {
                        nRow++;
                        nOff++;
                        continue;
                    }

                    if ((dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "TELEPHONE")].ToolTipText.ToString() != "") && !(checkedListBox1.CheckedItems.Contains("Ignore Dups")))
                    {
                        dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "RESULTS")].Selected = true;
                        dataGridView1.CurrentCell.Value = "F";
                        SetResultTime("F");
                        dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "NOTES")].Selected = true;
                        dataGridView1.CurrentCell.Value = dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "TELEPHONE")].ToolTipText.ToString();
                        int s = cellch.Count;
                        cellch.Add(new List<string>());
                        cellch[s].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[s].Add(dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "TELEPHONE")].ToolTipText.ToString());
                        if (checkBox1.Checked) SaveSheetChanges(null);
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
                if (nRow >= nLim + 1/*+nOff*/)
                {
                    updateinprogress = false;
                    dataGridView1.Visible = true;
                    goto LB2;
                }
                int l = 0;
                int c = cellch.Count;

                string streetadr = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "ADDRESS")].Value?.ToString();
                string names = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "NAME")].Value?.ToString();
                string phone = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "TELEPHONE")].Value?.ToString();

                if ((string.IsNullOrEmpty(names) || string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(streetadr)))
                {
                    if (string.IsNullOrEmpty(streetadr))
                    {
                        dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "RESULTS")].Selected = true;
                        dataGridView1.CurrentCell.Value = "F";
                        SetResultTime("F");
                        dataGridView1.Rows[nRow].Cells[GetindexOf(dataGridView1, "NOTES")].Selected = true;
                        dataGridView1.CurrentCell.Value = "no address found";
                        int s = cellch.Count;
                        cellch.Add(new List<string>());
                        cellch[s].Add(dataGridView1.CurrentCellAddress.ToString());
                        cellch[s].Add(dataGridView1.CurrentCell.Value.ToString());
                        if (checkBox1.Checked) SaveSheetChanges(null);
                    }
                    if (nRow < nLim)
                    {
                        nRow++;
                    }
                    goto LB1;
                }

                string phone1 = Regex.Replace(phone, "[^0-9]", "");
                phone1 = phone1.TrimStart('0');
                if (phone != phone1)
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
                    string list = "";

                    streetadr = RestSharp.Extensions.MonoHttp.HttpUtility.UrlEncode(streetadr); // encode for specail characters

                    list = ExtractCoordinates(streetadr);
                    if (streetadr.Length > 180)                             // limit url size
                        streetadr = streetadr.Substring(0, 180);

                    if (!string.IsNullOrEmpty(list))
                    {
                        //GPSgeofence fence = new GPSgeofence();
                        haveGPS = true;

                        string[] loc = list.Split(',');
                        try
                        {
                            if (loc.Length > 1)
                            {
                                lx = (float)Convert.ToDouble(loc[0].TrimEnd('+'));
                                ly = (float)Convert.ToDouble(loc[1].TrimEnd('+'));
                                if (Math.Abs(ly) < Math.Abs(lx))
                                    found = fence.PointInPolygon(ly, lx);
                                else
                                    found = fence.PointInPolygon(lx, ly);
                            }
                            else
                            {
                                haveGPS = false;
                                found = false;
                                vs.Initialize();
                            }
                            streetadr = list.TrimEnd('+');
                        }
                        catch (FormatException)
                        {
                            found = false;
                            haveGPS = false;
                        }

                    }
                    int t = 3;
                    while (t-- > 0 && !haveGPS)
                    {
                        vs = getGPSfromAddr(streetadr, cityadr);

                        if (t == 2 && (vs == null || vs[0] == "X"))
                        {
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = "-";
                            dataGridView1.Update();
                            streetadr = RemoveConsecutive(streetadr);
                            continue;
                        }
                        if (t == 1 && (vs == null || vs[0] == "X"))
                        {
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = "x";
                            dataGridView1.Update();
                            streetadr = RemoveAnyDuplicates(streetadr);
                            continue;
                        }
                        break;
                    }
                    if ((dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value == "-") ||
                        (dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value == "x"))
                    {
                        dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value = "X";
                        dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = "X";
                        dataGridView1.Update();
                        numOfSkip++;
                    }
                    dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "DATE")].Value = today.ToString("yyyy-MM-dd");
                    if (vs == null || vs[0] == "E" || vs[0] == "X")
                        found = false;
                    else
                        found = true;
                    //}
                    if (haveGPS)
                    {
                        if (found)
                            vs = new string[] { "In", "g", Convert.ToString(lx), Convert.ToString(ly) };
                        else
                            vs = new string[] { "O", "g", Convert.ToString(lx), Convert.ToString(ly) };
                        haveGPS = false;
                    }
                    //}

                    string t1 = dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "RESULTS")].Value?.ToString();

                    if (found)
                    {
                        if (vs[0] == "In" || vs[0] == "G")
                        {
                            SetResultTime("In");
                            numOfEN++;
                            if (radioButton6.Checked)
                                textBox6.Text = string.Format("found {0}", numOfEN);
                            if (radioButton5.Checked)
                                textBox6.Text = string.Format("encontrado {0}", numOfEN);
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
                                if (radioButton6.Checked)
                                    textBox7.Text = string.Format("found {0}", numOfSP);
                                if (radioButton5.Checked)
                                    textBox7.Text = string.Format("encontrado {0}", numOfSP);
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
                        }
                        else
                        {
                            SetResultTime("X");
                            numOfSkip++;
                            textBox8.Update();
                            dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Selected = true;
                            //dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value = vs[0];
                            c = cellch.Count;
                            cellch.Add(new List<String>());
                            cellch[c].Add(dataGridView1.CurrentCellAddress.ToString());
                            cellch[c].Add(dataGridView1.CurrentRow.Cells[GetindexOf(dataGridView1, "Confidence")].Value.ToString());
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

                        //}
                        //else
                        //if (vs == null) numOfSkip++;
                    }
                    if (radioButton6.Checked)
                        textBox8.Text = string.Format("skipped {0}", numOfSkip);
                    if (radioButton5.Checked)
                        textBox8.Text = string.Format("omitido {0}", numOfSkip);

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

            }
        LB2: if (comboBox1.SelectedIndex == 0)
            {
                progressBar1.Value = 100;
                progressBar1.Update();
                label9.Visible = true;
                button9.Visible = true;
                checkBox2.Visible = true;
                checkBox3.Visible = true;
                if (radioButton6.Checked)
                    textBox1.Text = " done ";
                if (radioButton5.Checked)
                    textBox1.Text = " hecho ";
                textBox1.Update();

                button8.BackColor = System.Drawing.Color.LightGray;
                button9.BackColor = System.Drawing.Color.LightGreen;
            }
            Tabfocus = "Imported Names";
            return;
        }

        //
        // get GPS coordinates from address and city
        //
        /// <summary>
        /// The getGPSfromAddr.
        /// </summary>
        /// <param name="address">The address<see cref="string"/>.</param>
        /// <param name="city">The city<see cref="string"/>.</param>
        /// <returns>The <see cref="string[]"/>.</returns>
        public string[] getGPSfromAddr(string address, string city)
        {
            bool inTerritory = false;
            string strx = "";
            string stry = "";
            int pos = -1;
            int pos1 = -1;
            string NewText = "";
            GPSgeofence gPSgeofence = new GPSgeofence();
            GPSgeofence fence = gPSgeofence;
            fence.ReadGPSfence();
            HttpWebRequestHandler hTTPrequest = new HttpWebRequestHandler();
            //http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=Guamani&postalCode=-&addressLine={addressLine}&userLocation=-&userIp={-}&usermapView={usermapView}&includeNeighborhood=includeNeighborhood&maxResults={maxResults}&key=

            string webAdr = Properties.Settings.Default.WebURL;
            address = address.Replace("+", "%20");
            webAdr = webAdr.Replace("Guamani", city);
            webAdr = webAdr.Replace("{addressLine}", address);
            webAdr = webAdr + fence.bkey;
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

                inTerritory = fence.PointInPolygon(x, y);
                if (x == notfoundlat && y == notfoundlon)
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
        /// <summary>
        /// The getAllPhoneNumbers.
        /// </summary>
        /// <param name="sheetService">The sheetService<see cref="SheetsService"/>.</param>
        /// <param name="spreadsheet">The spreadsheet<see cref="string"/>.</param>
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
                    if (radioButton6.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                    if (radioButton5.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicado {0}", forchk.ToString());

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
                    if (radioButton6.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                    if (radioButton5.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicado {0}", forchk.ToString());
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
                        if (radioButton6.Checked)
                            System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                        if (radioButton5.Checked)
                            System.Windows.Forms.MessageBox.Show("duplicado {0}", forchk.ToString());
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
                    if (radioButton6.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicate {0}", forchk.ToString());
                    if (radioButton5.Checked)
                        System.Windows.Forms.MessageBox.Show("duplicado {0}", forchk.ToString());
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
        /// <summary>
        /// The copyrange.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
                // prepare move information to "Out of Terr" or "Confired Deaf
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
        /// <summary>
        /// The button10_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button10_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                checkedListBox1.Items.Remove("Ignore Dups");
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
                checkedListBox1.Items.Remove("Ignore Dups");
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
        /// <summary>
        /// Defines the _attempt.
        /// </summary>
        public decimal _attempt = 0;

        /// <summary>
        /// The checkBox4_CheckedChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            pastTime = DateTimePicker.MinimumDateTime;
            string[] days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
            if (radioButton6.Checked)
            {
                days[0] = "Sunday";
                days[1] = "Monday";
                days[2] = "Tuesday";
                days[3] = "Wednesday";
                days[4] = "Thursday";
                days[5] = "Friday";
                days[6] = "Saturday";
            }
            if (radioButton5.Checked)
            {
                days[0] = "Domingo";
                days[1] = "Lunes";
                days[2] = "Martes";
                days[3] = "Miercoles";
                days[4] = "Jueves";
                days[5] = "Viernes";
                days[6] = "Sabado";
            }

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

        /// <summary>
        /// The button11_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
                        progress = remainder / (dataGridView1.Rows.Count / 100);
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
            DialogResult result = DialogResult.No;
            if (radioButton6.Checked)
                result = System.Windows.Forms.MessageBox.Show(string.Format(" there are {0} rows recommended. Do you want to accept and work these?", sel_rows),
                                                                 "Important Question", MessageBoxButtons.YesNo);
            if (radioButton5.Checked)
                result = System.Windows.Forms.MessageBox.Show(string.Format("  hay {0} filas recomendadas. ¿Quiere aceptarlas y trabajarlas?", sel_rows),
                                                          "Pregunta importante", MessageBoxButtons.YesNo);

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
        /// <summary>
        /// The seeVisibleRow.
        /// </summary>
        /// <param name="view">The view<see cref="DataGridView"/>.</param>
        /// <param name="rowToShow">The rowToShow<see cref="int"/>.</param>
        private static void seeVisibleRow(DataGridView view, int rowToShow)
        {
            int i = 0;
            int j = 0;
            var countVisible = view.DisplayedRowCount(false);

            view.CurrentRow.Selected = false;
            for (i = rowToShow; i > 0; i--)
                if (view.Rows[i].Visible) break;
            view.Rows[i].Selected = true;
            view.Rows[i].Cells[1].Selected = true;
            while (j + 5 < i)
            {
                j = view.FirstDisplayedScrollingRowIndex;
                if (view.Rows[j++].Visible) break;
            }
            view.Update();
        }

        /// <summary>
        /// The lang.
        /// </summary>
        /// <param name="lang">The lang<see cref="string"/>.</param>
        private void lang(string lang)
        {
            string culture = lang;

            Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo(culture);
            ApplyResourceToControl(
            this,
            new ComponentResourceManager(typeof(Form1)),
            new CultureInfo(culture));
        }

        /// <summary>
        /// The ApplyResourceToControl.
        /// </summary>
        /// <param name="control">The control<see cref="Control"/>.</param>
        /// <param name="cmp">The cmp<see cref="ComponentResourceManager"/>.</param>
        /// <param name="cultureInfo">The cultureInfo<see cref="CultureInfo"/>.</param>
        private void ApplyResourceToControl(
       Control control,
       ComponentResourceManager cmp,
       CultureInfo cultureInfo)
        {
            foreach (Control child in control.Controls)
            {
                //Store current position and size of the control
                var childSize = child.Size;
                var childLoc = child.Location;
                //Apply CultureInfo to child control
                ApplyResourceToControl(child, cmp, cultureInfo);
                //Restore position and size
                child.Location = childLoc;
                child.Size = childSize;
            }
            //Do the same with the parent control
            var parentSize = control.Size;
            var parentLoc = control.Location;
            cmp.ApplyResources(control, control.Name, cultureInfo);
            control.Location = parentLoc;
            control.Size = parentSize;
        }

        /// <summary>
        /// The ChangeLanguage.
        /// </summary>
        /// <param name="lang">The lang<see cref="string"/>.</param>
        public void ChangeLanguage(string lang) //A function called to change the language
        {
            foreach (Control c in this.Controls)
            {
                ComponentResourceManager resources = new ComponentResourceManager(typeof(Form1));
                //resources.ApplyResources(c, c.Name, new CultureInfo(lang));
                ApplyResourceToControl(this, new ComponentResourceManager(typeof(Form1)), new CultureInfo(lang));
            }
        }

        /// <summary>
        /// Defines the _previousSelectedItem.
        /// </summary>
        internal int _previousSelectedItem = 1;

        /// <summary>
        /// The comboBox1_SelectionChangeCommitted.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            ComboBox xx = (ComboBox)sender;
            xx.SelectedItem = this._previousSelectedItem;
        }

        /// <summary>
        /// The radioButton6_CheckedChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                _previousSelectedItem = comboBox1.SelectedIndex;
                //ChangeLanguage to American English
                string language = "en-US";

                // Save user choice in settings
                Properties.Settings.Default.Language = "en-US";
                Properties.Settings.Default.Save();

                ChangeLanguage(language);
                comboBox1.SelectedIndex = _previousSelectedItem;
                string tstring = comboBox1.SelectedText;
                comboBox1.SelectedIndex = comboBox1.FindStringExact(tstring);
                BuildComboBoxItems(_previousSelectedItem);
                this.Update();
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }

        /// <summary>
        /// The radioButton5_CheckedChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                _previousSelectedItem = comboBox1.SelectedIndex;
                //ChangeLanguage to Ecuadorian Spanish
                string language = "es-EC";

                // Save user choice in settings
                Properties.Settings.Default.Language = "es-EC";
                Properties.Settings.Default.Save();

                ChangeLanguage(language);
                comboBox1.SelectedIndex = _previousSelectedItem;
                string tstring = comboBox1.SelectedText;
                comboBox1.SelectedIndex = comboBox1.FindStringExact(tstring);
                BuildComboBoxItems(_previousSelectedItem);
                this.Update();
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
            }
        }
    }
}
