using Google.GData.Client;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Net;
using System.Web.UI;
using System.Linq;

namespace WFAGoolgeSheet
{
    static class Program
    {

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]


        static void Main()
        {
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        //
        // remove + from changes array
        //
        // Arrays should contain paired parentheses in the same order:
        private static readonly char[] OpenParentheses = { '+', '(', '[', '{' };
        private static readonly char[] CloseParentheses = { '+', ')', ']', '}' };
        public static bool formisup = false;
        public static object WFAgoolgeSheet { get; private set; }

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

        //
        // map column to spreadsheet letter
        //    
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
        /// Convert a hex string to a .NET Color object.
        /// </summary>
        /// <param name="hexColor">a hex string: "FFFFFF", "#000000"</param>
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

        internal static Google.Apis.Sheets.v4.Data.ValueRange Try<T>(Google.Apis.Sheets.v4.Data.ValueRange valueRange, int v)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Extract only the hex digits from a string.
        /// </summary>
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

        //
        // Return a Google map URL.
        //
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

        // Return a Google map type code.
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

        //
        // translate a textbox
        //
        public static String GTranslate(String sentance)
        {
            string RetSentance="";
            var toLanguage = "sp";//Spanish
            var fromLanguage = "en";//English
            var punctuation = sentance.Where(Char.IsPunctuation).Distinct().ToArray();
            var words = sentance.Split().Select(x => x.Trim(punctuation));
            foreach(var word in words)
            {
                var url = $"https://translate.googleapis.com/translate_a/single?client=gtx&sl={fromLanguage}&tl={toLanguage}&dt=t&q={HttpUtility.UrlEncode(word)}";
                var webClient = new WebClient
                {
                    Encoding = System.Text.Encoding.UTF8
                };
                var result = webClient.DownloadString(url);
                try
                {
                    result = result.Substring(4, result.IndexOf("\"", 4, StringComparison.Ordinal) - 4);
                    RetSentance = RetSentance + " "+ result;
                }
                catch
                {
                    return "";
                }
            }
            return (RetSentance);
        }

        //Get the HtmlAgilityPack here: http://www.codeplex.com/htmlagilitypack
        public static async Task<T> Try<T>(this Func<T> func, int retries)
        {
            var i = 0;
            do
            {
                
                try
                {
                    return await Task.Run(func);
                    //Task t1 = Task.Run(() => Method1());
                }
                catch (TaskCanceledException exception)
                {
                    Console.WriteLine(exception.Message);
                }
            } while (i++ < retries);
            return default(T);
        }

    }
}



