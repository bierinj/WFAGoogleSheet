using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;
using BingMapsRESTToolkit;
using BingMapsRESTToolkit.Extensions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System.Threading;

namespace WFAGoolgeSheet
{
    public partial class GPSgeofence : Component
    {
        // static GPS fence values
        // and Bing Maps values
        public string bkey = "AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-";
        Form1 form1 = new Form1();

        //
        // Google parameters
        //static UserCredential credential;               // Sheet Service
        //static string ApplicationName = null;
        //static String spreadsheetId = null;             // Spreadsheet ID

        //private float[] polyX;
        //-0.32541f,
        //-0.30631F,
        //-0.30640F,
        //-0.33163F,
        //-0.33103F,
        //-0.36622F,
        //-0.37137F,
        //-0.35979F,
        //-0.34614F,
        //-0.32545F,
        //-0.31189F};  // horizontal coordinates of corners
        //private float[] polyY;
        //-78.58218f,
        //-78.56296F,
        //-78.55945F,
        //-78.55515F,
        //-78.55043F,
        //-78.55086F,
        //-78.55919F,
        //-78.57481F,
        //-78.58253F,
        //-78.58271F,
        //-78.56975F};                        // vertical coordinates of corners

        bool GPSfenceDone = false;
        int polyCorners = 0;            // how many corners the polygon has


        List<float> polyX = new List<float>();
        List<float> polyY = new List<float>();

        //float x, y = 0;                                   // point to be tested

        public GPSgeofence()
        {
            InitializeComponent();
        }

        public GPSgeofence(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        public void ReadGPSfence()
        {
            if (GPSfenceDone) return;
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
            String range = "GeoFence Data!A4:C";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(form1.spreadsheetId, range);
            ValueRange response = request.Execute();
            polyCorners = response.Values.Count;
            form1.textBox1.Text = ".. reading data";
            form1.textBox1.Update();
            for(int i=0; i < polyCorners; i++)
            {
                //polyX = $"{response.Values[i][1]}F";
                float tempx;
                float.TryParse(response.Values[i][1].ToString(), out tempx);
                polyX.Add(tempx);
                float.TryParse(response.Values[i][2].ToString(), out tempx);
                polyY.Add(tempx);
            }
            GPSfenceDone = true;
        }
        //  Globals which should be set before calling this function:
        //
        //  int    polyCorners  =  how many corners the polygon has
        //  float  polyX[]      =  horizontal coordinates of corners
        //  float  polyY[]      =  vertical coordinates of corners
        //  float  x, y         =  point to be tested
        //
        //  (Globals are used in this example for purposes of speed.  Change as
        //  desired.)
        //
        //  The function will return YES if the point x,y is inside the polygon, or
        //  NO if it is not.  If the point is exactly on the edge of the polygon,
        //  then the function may return YES or NO.
        //
        //  Note that division by zero is avoided because the division is protected
        //  by the "if" clause which surrounds it.

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
    }
}

