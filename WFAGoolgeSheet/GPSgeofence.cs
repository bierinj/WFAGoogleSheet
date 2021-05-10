namespace WFAGoolgeSheet
{
    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Services;
    using Google.Apis.Sheets.v4;
    using Google.Apis.Sheets.v4.Data;
    using Google.Apis.Util.Store;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Threading;

    /// <summary>
    /// Defines the <see cref="GPSgeofence" />.
    /// </summary>
    public partial class GPSgeofence : Component
    {
        /// <summary>
        /// Defines the bkey.
        /// </summary>
        public string bkey = Properties.Settings.Default.WebKey;

        /// <summary>
        /// Defines the fenceName.
        /// </summary>
        public string fenceName;

        /// <summary>
        /// Defines the polyCorners.
        /// </summary>
        public int polyCorners = 0;// how many corners the polygon has

        /// <summary>
        /// Defines the polyX.
        /// </summary>
        public List<float> polyX = new List<float>();

        /// <summary>
        /// Defines the polyY.
        /// </summary>
        public List<float> polyY = new List<float>();

        /// <summary>
        /// Defines the firstTime.
        /// </summary>
        internal static bool firstTime = true;

        /// <summary>
        /// Defines the form1.
        /// </summary>
        internal Form1 form1 = new Form1();

        /// <summary>
        /// Initializes a new instance of the <see cref="GPSgeofence"/> class.
        /// </summary>
        public GPSgeofence()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GPSgeofence"/> class.
        /// </summary>
        /// <param name="container">The container<see cref="IContainer"/>.</param>
        public GPSgeofence(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        /// <summary>
        /// Gets the mc2.
        /// </summary>
        public Form1 mc2 { get; private set; }

        /// <summary>
        /// The PointInPolygon.
        /// </summary>
        /// <param name="x">The x<see cref="float"/>.</param>
        /// <param name="y">The y<see cref="float"/>.</param>
        /// <returns>The <see cref="bool"/>.</returns>
        public bool PointInPolygon(float x, float y)
        {
            int i, j = polyCorners - 1;
            bool oddNodes = false;

            for (i = 0; i < polyCorners; i++)
            {
                if ((polyY[i] < y && polyY[j] >= y
                || polyY[j] < y && polyY[i] >= y)
                && (polyX[i] <= x || polyX[j] <= x))
                {
                    oddNodes ^= (polyX[i] + (y - polyY[i]) / (polyY[j] - polyY[i]) * (polyX[j] - polyX[i]) < x);
                }
                j = i;
            }
            return oddNodes;
        }

        /// <summary>
        /// The ReadGPSfence.
        /// </summary>
        public void ReadGPSfence()
        {

            if (Form1.GPSfenceDone) return;
            form1.startPB(System.Drawing.Color.Green);
            form1.textBox1.Text = " ..logging in ";
            form1.textBox1.Update();

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
                form1.credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            form1.textBox1.Text = "..connecting Sheet ";
            form1.textBox1.Update();
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = form1.credential,
                ApplicationName = ApplicationName,
            });
            form1.progressBar1.Value = 2;
            form1.progressBar1.Update();
            // Define request parameters.
            if (form1.radioButton1.Checked == false) form1.spreadsheetId = Properties.Settings.Default.TestSheet;
            if (form1.radioButton2.Checked) form1.spreadsheetId = Properties.Settings.Default.ProdSheet;
            form1.progressBar1.Value = 3;
            form1.progressBar1.Update();
            String range = "GeoFence Data!A4:D";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(form1.spreadsheetId, range);
            // use retry function as task
            ValueRange response = request.Execute();
            //ValueRange response = Program.Try<System.String>(request.Execute(), (int)2);
            polyCorners = response.Values.Count;

            form1.textBox1.Text = ".. reading data";
            form1.textBox1.Update();
            for (int i = 0; i < polyCorners; i++)
            {
                fenceName = response.Values[i][3].ToString();
                float tempx;
                float.TryParse(response.Values[i][1].ToString(), out tempx);
                polyX.Add(tempx);
                float.TryParse(response.Values[i][2].ToString(), out tempx);
                polyY.Add(tempx);
            }
            Form1.GPSfenceDone = true;
        }
    }
}
