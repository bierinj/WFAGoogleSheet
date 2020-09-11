using Google.Apis.Auth.OAuth2;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WFAGoolgeSheet
{
    public partial class UserControl1 : UserControl
    {
        private Form1 form1;
        public UserControl1()
        {
            InitializeComponent();
        }


        //---------------------------------------------------------------------------//
        // global flags and data
        //

        //
        // Google parameters
        static UserCredential credential;               // Sheet Service
        static string ApplicationName = null;
        static String spreadsheetId = null;             // Spreadsheet ID
        List<List<String>> names2chk = new List<List<String>>();

        //Form1 form1 = new Form1();
        //public void RunImportedNames()
        //{
        //    using (var UserControl1 = new UserControl1())
        //    {

        //Form1 form1 = new Form1();
        //    UserControl1 userControl1 = new UserControl1();
        //    DataChanged = false;
        //    updateinprogress = true;
        //    form1.textBox1.Text = " ..logging in ";
        //    form1.textBox1.Update();

        //    // If modifying these scopes, delete your previously saved credentials
        //    // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        //    string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        //    string ApplicationName = "Google Sheets API .NET Quickstart";

        //    //UserCredential credential;

        //    Cursor.Current = Cursors.WaitCursor;

        //    using (var stream =
        //            new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
        //    {
        //        // The file token.json stores the user's access and refresh tokens, and is created
        //        // automatically when the authorization flow completes for the first time.
        //        string credPath = "token.json";
        //        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
        //            GoogleClientSecrets.Load(stream).Secrets,
        //            Scopes,
        //            "user",
        //            CancellationToken.None,
        //            new FileDataStore(credPath, true)).Result;
        //        Console.WriteLine("Credential file saved to: " + credPath);
        //    }
        //    form1.textBox1.Text = "..connecting Sheet ";
        //    form1.textBox1.Update();
        //    // Create Google Sheets API service.
        //    var service = new SheetsService(new BaseClientService.Initializer()
        //    {
        //        HttpClientInitializer = credential,
        //        ApplicationName = ApplicationName,
        //    });

        //    // Define request parameters.
        //    //spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
        //    //spreadsheetId = "1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI";
        //    string spreadsheetId1 = "1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok";    // test sheet
        //    string spreadsheetId2 = "1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc";    // live sheet  /edit#gid=2145664999"

        //    //
        //    // process RadioButtons on Form1
        //    //var checkedRadio = new[] { groupBox1 }.SelectMany(g => g.Controls.OfType<RadioButton>()
        //    //                    .Where(r => r.Checked));
        //    //*if (form1.radioButton1.Checked) */
        //        spreadsheetId = spreadsheetId1;
        //    if (form1.radioButton2.Checked) spreadsheetId = spreadsheetId2;

        //    //int selectIndex = form1.comboBox1.SelectedIndex;
        //    //selectCar = "List of Names";
        //    String range = "Common First Names!A1:A";
        //    //String range = selectCar.tabname + selectCar.range;

        //    //UserControl..ActiveForm.Text = String.Format("Working with sheet '{0}', columns {1} ", spreadsheetId, range);
        //    SpreadsheetsResource.ValuesResource.GetRequest request =
        //            service.Spreadsheets.Values.Get(spreadsheetId, range);

        //    // Prints the names and majors of students in a sample spreadsheet:
        //    // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
        //    // https://docs.google.com/spreadsheets/d/1VARB7hr74Q89Z3g9idDs1WUniBZoZCyLzSZgtnEAJkI/edit#gid=1072334158
        //    // https://docs.google.com/spreadsheets/d/1vsKqcYoxStY9ksmEDZodkOpWf5wOcXx_FXzXeEGfpok/edit#gid=2145664999
        //    // https://docs.google.com/spreadsheets/d/1Ju9HyYti08VlqaILvQXEWyiflD9DJUzXZ9KYYt5hBzc/edit#gid=2145664999
        //    ValueRange response = request.Execute();
        //    form1.textBox1.Text = ".. reading data";
        //    form1.textBox1.Update();
        //    IList<IList<Object>> values = response.Values;
        //    int i;
        //    foreach(var row in values)
        //    {
        //        i = names2chk.Count;
        //        //if (string.IsNullOrWhiteSpace(values[][0].ToString()))
        //        //    continue;

        //        names2chk.Add(new List<String>()); //Adds new sub List
        //        names2chk[i].Add(values[i][0].ToString()); //Add values to the sub List at index 0
        //        names2chk[i].Add(" ".ToString());
        //    }
        //    //userControl1.Visable = true;
        //    //userControl1.StartPosition = FormStartPosition.Manual;
        //    userControl1.Left = SecondFormLeft;
        //    userControl1.Top = SecondFromTop;
        //                    if (form1.SecondFormSize == System.Drawing.Size.Empty)
        //                        form1.SecondFormSize = userControl1.Size;
        //    userControl1.Size = form1.SecondFormSize;;
        //    //userControl1.StartPosition = FormStartPosition.Manual;
        //    userControl1.Left = SecondFormLeft;
        //    userControl1.Top = SecondFromTop;
        //    if (form1.SecondFormSize == System.Drawing.Size.Empty)
        //        form1.SecondFormSize = userControl1.Size;
        //    userControl1.Size = form1.SecondFormSize;
        //    //userControl1.ShowDialog();


        //    int nRow = 0;
        //    while (nRow < 10 - 1)
        //    {
        //        if (form1.dataGridView1.Rows[++nRow].Visible == false)
        //            continue;
        //        form1.dataGridView1.Rows[nRow].Selected = true;
        //        int currentRow = nRow;
        //        form1.dataGridView1.Rows[currentRow].Cells[0].Selected = true;
        //        //form1.dataGridView1.CurrentCell.ro = nRow;
        //        form1.dataGridView1.CurrentRow.Selected = true;
        //        break;
        //    }

        //    int l = 0;
        //    int J = names2chk.Count;
        //    bool found = false;
        //    foreach (var row in values)
        //    {
        //        for(l = 0; l < names2chk.Count; l++)
        //            {
        //            if (names2chk[l][0].ToString() == form1.dataGridView1.CurrentRow.Cells[1].Value.ToString())
        //                found = true;
        //            break;
        //        }
        //        if (found)
        //        {
        //            form1.dataGridView1.CurrentRow.Cells[4].Value = " ";
        //        }
        //        else
        //        {
        //            form1.dataGridView1.CurrentRow.Cells[4].Value = "SP";
        //        }
        //    }

        //}
        //return;
        //}

        //private DialogResult ShowDialog()
        //{
        //    //throw new NotImplementedException();
        //}

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
