namespace WFAGoolgeSheet
{
    using Google.GData.Client;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="Program" />.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Defines the CloseParentheses.
        /// </summary>
        private static readonly char[] CloseParentheses = { '+', ')', ']', '}' };

        /// <summary>
        /// Defines the OpenParentheses.
        /// </summary>
        private static readonly char[] OpenParentheses = { '+', '(', '[', '{' };

        /// <summary>
        /// Defines the formisup.
        /// </summary>
        public static bool formisup = false;

        /// <summary>
        /// Defines the form1.
        /// </summary>
        internal static Form1 form1 = new Form1();

        /// <summary>
        /// Gets the WFAgoolgeSheet.
        /// </summary>
        public static object WFAgoolgeSheet { get; private set; }

        /// <summary>
        /// The Check.
        /// </summary>
        /// <param name="input">The input<see cref="string"/>.</param>
        /// <returns>The <see cref="bool"/>.</returns>
        public static bool Check(string input)
        {
            // Indices of the currently open parentheses:
            Stack<int> parentheses = new Stack<int>();

            foreach (char chr in input)
            {
                int index;

                // Check if the 'chr' is an open parenthesis, and get its index:
                if ((index = Array.IndexOf(OpenParentheses, chr)) != -1)
                {
                    parentheses.Push(index);  // Add index to stach
                }
                // Check if the 'chr' is a close parenthesis, and get its index:
                else if ((index = Array.IndexOf(CloseParentheses, chr)) != -1)
                {
                    // Return 'false' if the stack is empty or if the currently
                    // open parenthesis is not paired with the 'chr':
                    if (parentheses.Count == 0 || parentheses.Pop() != index)
                        return false;
                }
            }
            // Return 'true' if there is no open parentheses, and 'false' - otherwise:
            return parentheses.Count == 0;
        }

        /// <summary>
        /// The ColumnAdress.
        /// </summary>
        /// <param name="col">The col<see cref="int"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string ColumnAdress(int col)
        {
            col++;
            if (col <= 26)
            {
                return Convert.ToChar(col + 64).ToString();
            }
            int div = col / 26;
            int mod = col % 26;
            if (mod == 0) { mod = 26; div--; }
            return ColumnAdress(div) + ColumnAdress(mod);
        }

        /// <summary>
        /// The EnclosedStrings.
        /// </summary>
        /// <param name="s">The s<see cref="string"/>.</param>
        /// <param name="begin">The begin<see cref="string"/>.</param>
        /// <param name="end">The end<see cref="string"/>.</param>
        /// <returns>The <see cref="IEnumerable{string}"/>.</returns>
        public static IEnumerable<string> EnclosedStrings(
    this string s,
    string begin,
    string end)
        {
            int beginPos = s.IndexOf(begin, 0);
            while (beginPos >= 0)
            {
                int start = beginPos + begin.Length;
                int stop = s.IndexOf(end, start);
                if (stop < 0)
                    yield break;
                yield return s.Substring(start, stop - start);
                beginPos = s.IndexOf(begin, stop + end.Length);
            }
        }

        /// <summary>
        /// Extract only the hex digits from a string.
        /// </summary>
        /// <param name="input">The input<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string ExtractHexDigits(string input)
        {
            // remove any characters that are not digits (like #)
            Regex isHexDigit
               = new Regex("[abcdefABCDEF\\d]+", RegexOptions.Compiled);
            string newnum = "";
            foreach (char c in input)
            {
                if (isHexDigit.IsMatch(c.ToString()))
                    newnum += c.ToString();
            }
            return newnum;
        }

        /// <summary>
        /// The GoogleMapUrl.
        /// </summary>
        /// <param name="query">The query<see cref="string"/>.</param>
        /// <param name="map_type">The map_type<see cref="string"/>.</param>
        /// <param name="zoom">The zoom<see cref="int"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string GoogleMapUrl(string query, string map_type, int zoom)
        {
            // Start with the base map URL.
            // http://maps.google.com/maps/api/geocode/xml?address=Humberto+Albornoz+1339+Y+Lizarazu+-+Humberto+Albornoz&sensor=false
            string url = "http://maps.google.com/maps?";
            //string url = "http://maps.google.com/maps/api/geocode/xml?address=";

            // Add the query.
            url += "q=" + HttpUtility.UrlEncode(query, Encoding.UTF8);

            // Add the type.
            map_type = GoogleMapTypeCode(map_type);
            if (map_type != null) url += "&t=" + map_type;

            // Add the zoom level.
            if (zoom > 0) url += "&z=" + zoom.ToString();
            GPSgeofence fence = new GPSgeofence();
            bool p = fence.PointInPolygon(
                (float)-0.31920, (float)-78.56841);
            if (p)
                Console.WriteLine("Point is in");
            else
                Console.WriteLine("Point is out");
            return url;
        }

        /// <summary>
        /// The GTranslate.
        /// </summary>
        /// <param name="sent">The sent<see cref="string"/>.</param>
        /// <param name="idioma">The idioma<see cref="string"/>.</param>
        /// <param name="detectonly">The detectonly<see cref="bool"/>.</param>
        /// <returns>The <see cref="String"/>.</returns>
        public static async Task<string> gTranslate(string sent, string idioma, bool detectonly)
        {
            string RetSentance = "";
            string subscriptionKey = Properties.Settings.Default.AzureKey;
            string endpoint = Properties.Settings.Default.AzureEndPt;

            /// <summary>
            /// Defines the location.
            /// </summary>
            string location = Properties.Settings.Default.AzureLoc;

            if (!form1.checkBox5.Checked || (string.IsNullOrEmpty(subscriptionKey)) ||
                string.IsNullOrEmpty(endpoint) ||
                string.IsNullOrEmpty(location)) return ("");
            RetSentance = sent.Replace("(en):", "");
            RetSentance = sent.Replace("(es):", "");
            form1.textBox1.Text = string.Format(" translate {0}", idioma);
            form1.textBox1.Refresh();
            Thread.Sleep(100);
            // Input and output languages are defined as parameters.
            // Build the request.
            string route = null;
            if (detectonly)
                route = "detect?api-version=3.0";
            else
            {
                if (idioma == "en")
                    route = @"translate?api-version=3.0&from=en&to=es";
                else
                    route = @"translate?api-version=3.0&from=es&to=en";
            }

            string textToTranslate = RetSentance;
            textToTranslate = textToTranslate.Where(c => !char.IsPunctuation(c)).Aggregate("", (current, c) => current + c);
            object[] body = new object[] { new { Text = textToTranslate } };
            string requestBody = JsonConvert.SerializeObject(body);
            try
            {

                using (var client = new HttpClient())
                using (var request = new HttpRequestMessage())
                {
                    // Build the request.
                    request.Method = HttpMethod.Post;
                    request.RequestUri = new Uri(endpoint + route);
                    request.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                    request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
                    request.Headers.Add("Ocp-Apim-Subscription-Region", location);

                    // Send the request and get response.
                    var response = await client.SendAsync(request).ConfigureAwait(continueOnCapturedContext: false);
                    //System.IO.Stream apiresponse = await client.GetStreamAsync(request.ToString())/*.ConfigureAwait(false)*/;
                    // Read response as a string.

                    if (response.IsSuccessStatusCode)
                        RetSentance = await response.Content.ReadAsStringAsync();
                    if (detectonly)
                    {
                        string[] retstr = RetSentance.Split(':');
                        retstr = retstr[1].Split(',');
                        retstr[0] = retstr[0].Replace("\\\"", "");
                        retstr = retstr[0].Split('"');
                        RetSentance = retstr[1];
                    }
                    else
                    {
                        string[] retstr = RetSentance.Split(':');
                        retstr = retstr[2].Split(',');
                        retstr[0] = retstr[0].Replace("\\\"", "");
                        retstr = retstr[0].Split('"');
                        RetSentance = retstr[1];
                    }
                    if (sent.Contains("(en):")) RetSentance = "(en):" + RetSentance;
                    if (sent.Contains("(es):")) RetSentance = "(es):" + RetSentance;
                    return (RetSentance);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return ("");
        }

        /// <summary>
        /// Convert a hex string to a .NET Color object.
        /// </summary>
        /// <param name="hexColor">a hex string: "FFFFFF", "#000000".</param>
        /// <returns>The <see cref="Color"/>.</returns>
        public static Color HexStringToColor(string hexColor)
        {
            string hc = ExtractHexDigits(hexColor);
            if (hc.Length != 8)
            {
                // you can choose whether to throw an exception
                //throw new ArgumentException("hexColor is not exactly 6 digits.");
                return Color.Empty;
            }
            string r = hc.Substring(2, 2);
            string g = hc.Substring(4, 2);
            string b = hc.Substring(6, 2);
            Color color = Color.Empty;
            try
            {
                int ri
                   = Int32.Parse(r, System.Globalization.NumberStyles.HexNumber);
                int gi
                   = Int32.Parse(g, System.Globalization.NumberStyles.HexNumber);
                int bi
                   = Int32.Parse(b, System.Globalization.NumberStyles.HexNumber);
                color = Color.FromArgb(ri, gi, bi);
            }
            catch
            {
                // you can choose whether to throw an exception
                //throw new ArgumentException("Conversion failed.");
                return Color.Empty;
            }
            return color;
        }

        /// <summary>
        /// The ReplaceFirstOccurrence.
        /// </summary>
        /// <param name="Source">The Source<see cref="string"/>.</param>
        /// <param name="Find">The Find<see cref="string"/>.</param>
        /// <param name="Replace">The Replace<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string ReplaceFirstOccurrence(string Source, string Find, string Replace)
        {
            //int Place = Source.IndexOf(Find);
            //string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            //return result;
            Regex r = new Regex(Find, RegexOptions.IgnoreCase);

            return r.Replace(Source, Replace, 1);
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]


        internal static void Main()
        {
            //Application.SetCompatibleTextRenderingDefault(false);
            Application.EnableVisualStyles();
            //Application.Run(new CustomRowSelectionPainting());
            Application.Run(new Form1());
        }

        /// <summary>
        /// The Try.
        /// </summary>
        /// <typeparam name="T">.</typeparam>
        /// <param name="valueRange">The valueRange<see cref="Google.Apis.Sheets.v4.Data.ValueRange"/>.</param>
        /// <param name="v">The v<see cref="int"/>.</param>
        /// <returns>The <see cref="Google.Apis.Sheets.v4.Data.ValueRange"/>.</returns>
        internal static Google.Apis.Sheets.v4.Data.ValueRange Try<T>(Google.Apis.Sheets.v4.Data.ValueRange valueRange, int v)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// The GoogleMapTypeCode.
        /// </summary>
        /// <param name="map_type">The map_type<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        private static string GoogleMapTypeCode(string map_type)
        {
            // Insert the proper type.
            switch (map_type)
            {
                case "Map":
                    return "m";
                case "Satellite":
                    return "k";
                case "Hybrid":
                    return "h";
                case "Terrain":
                    return "p";
                case "Google Earth":
                    return "e";
                default:
                    return null;
            }
        }

        /// <summary>
        /// Defines the <see cref="CustomRowSelectionPainting" />.
        /// </summary>
        private class CustomRowSelectionPainting : Form
        {
        }
    }
}
