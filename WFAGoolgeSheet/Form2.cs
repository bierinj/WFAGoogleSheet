
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using Microsoft.Win32;

namespace WFAGoolgeSheet
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();

            this.Font = System.Drawing.SystemFonts.IconTitleFont;
            SystemEvents.UserPreferenceChanged += new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
            this.FormClosing += new FormClosingEventHandler(Form2_FormClosing);

            formIsUp = true;
        }

        bool formIsUp = false;
        bool preventExit = false;
        bool inTerritory = false;
        string strx = "";
        string stry = "";

        //
        // OK button processing
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Form1 f1 = (Form1)Application.OpenForms["Form1"];
                Form2 f2 = (Form2)Application.OpenForms["Form2"];
                Form4 f4 = (Form4)Application.OpenForms["Form4"];

                f1.attempt = textBox5.Text;
                f1.notes = textBox6.Text;

                preventExit = false;
                formIsUp = false;
                f2.Close(); f4.Close();
                //this.Close();
            }
            catch (NullReferenceException ne)
            {
                //One of the forms is not opened
                Console.WriteLine(ne.Message);
            }
        }

        private void Form2_FormClosing(Object sender, FormClosingEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);

            if (preventExit) e.Cancel = true;
            else e.Cancel = false;
            formIsUp = false;
        }
        void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Window)
            {
                this.Font = System.Drawing.SystemFonts.IconTitleFont;
            }
        }

        //
        // Exit button processing
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Form2 f2 = (Form2)Application.OpenForms["Form2"];
                Form4 f4 = (Form4)Application.OpenForms["Form4"];
                preventExit = false;
                formIsUp = false;
                f2.Close(); f4.Close();
            }
            catch (NullReferenceException ne)
            {
                //One of the forms is not opened
                Console.WriteLine(ne.Message);
            }

        }

        //
        // Skip / Next - record processing
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Form2 f2 = (Form2)Application.OpenForms["Form2"];
                Form4 f4 = (Form4)Application.OpenForms["Form4"];

                formIsUp = false;
                preventExit = false;
                f2.Close(); f4.Close();
            }
            catch (NullReferenceException ne)
            {
                //One of the forms is not opened
                Console.WriteLine(ne.Message);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //
        // Skip back button
        //
        private void button4_Click(object sender, EventArgs e)
        {
            Form1 form1 = (Form1)Application.OpenForms["Form1"];
            try
            {
                Form2 f2 = (Form2)Application.OpenForms["Form2"];
                Form4 f4 = (Form4)Application.OpenForms["Form4"];
                preventExit = false;
                formIsUp  = false;
                f2.Close(); f4.Close();
            }
            catch (NullReferenceException ne)
            {
                //One of the forms is not opened
                Console.WriteLine(ne.Message);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form4 f4 = (Form4)Application.OpenForms["Form4"];
            Form1 form1 = new Form1();
            form1.button4.BackColor = System.Drawing.Color.LightGray;
            form1.button4.Update();
            if (form1.radioButton5.Checked)
            {
                int i = form1.comboBox1.SelectedIndex;
                //ChangeLanguage to Ecuadorian Spanish
                string language = "es-EC";

                // Save user choice in settings
                Properties.Settings.Default.Language = "es-EC";
                Properties.Settings.Default.Save();

                form1.ChangeLanguage(language);
                form1.comboBox1.SelectedIndex = i;
                form1.comboBox1.Refresh();
            }
            if (form1.radioButton6.Checked)
            {
                int i = form1.comboBox1.SelectedIndex;
                //ChangeLanguage to American English
                string language = "en-US";

                // Save user choice in settings
                Properties.Settings.Default.Language = "en-US";
                Properties.Settings.Default.Save();

                form1.ChangeLanguage(language);
                form1.comboBox1.SelectedIndex = i;
                form1.comboBox1.Refresh();
            }
            if (checkBox3.Checked)
                button5.PerformClick();
        }

        //
        // map via Google
        //
        public void textBox3_TextChanged(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            Form4 form4 = new Form4();
            double lat = 0;
            double lon = 0;

            if (!string.IsNullOrEmpty(textBox3.Text))
            {
                string streetadr = textBox3.Text;
                RegexOptions options = RegexOptions.None;               // remove multiple spaces
                Regex regex = new Regex("[ ]{2,}", options);
                streetadr = regex.Replace(streetadr, " ");
                if (streetadr.Length > 120)                             // limit url size
                    streetadr = streetadr.Substring(0, 120);
                streetadr = RestSharp.Extensions.MonoHttp.HttpUtility.UrlEncode(streetadr); // encode for specail characters
                string[] vs = form1.getGPSfromAddr(streetadr, "Quito");
                //form4.Size = form1.SecondFormSize;
                if(vs != null && vs[0] == "X")
                    return;
                if (!string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrEmpty(textBox11.Text))
                {
                    lat = Convert.ToDouble(textBox10.Text);
                    lon = Convert.ToDouble(textBox11.Text);
                }
                else
                {
                    if (vs != null)
                    {
                        lat = Convert.ToDouble(vs[2]);
                        lon = Convert.ToDouble(vs[3]);
                    }
                    else return;
                }

                form4.LoadIntoMap(lat, lon);

                form4.Show();

            }
            
            else
                textBox6.AppendText(" Address Is Blank");
            //textBox6.Text = Program.GTranslate(textBox6.Text);
            return;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            if (!form1.checkBox4.Checked && checkBox2.Checked)
                button2.PerformClick();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            Form4 f4 = new Form4();

            int pos = -1;
            int pos1 = -1;
            string NewText = "";
            if (string.IsNullOrEmpty(textBox10.Text) || string.IsNullOrEmpty(textBox11.Text))
            {
                HttpWebRequestHandler hTTPrequest = new HttpWebRequestHandler();
                //http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=Guamani&postalCode=-&addressLine={addressLine}&userLocation=-&userIp={-}&usermapView={usermapView}&includeNeighborhood=includeNeighborhood&maxResults={maxResults}&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS
                string webAdr = @"http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=-&postalCode=-&addressLine=" + textBox3.Text + "&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS";
                if (!string.IsNullOrEmpty(textBox3.Text))
                {
                    var webReply = hTTPrequest.GetReply(webAdr);
                    pos = webReply.IndexOf("\"coordinates\":");
                    pos1 = webReply.IndexOf("\"confidence\":");
                    if (pos1 > -1)
                        NewText = "confidence is " + webReply.Substring(pos1 + 14, 6) + "  ";
                    if (pos > -1)
                    {
                        string pwebReply = webReply.Substring(pos, 56);
                        pos = pwebReply.IndexOf(',');
                        textBox10.Text = pwebReply.Substring(15, pos - 15);
                        textBox10.Refresh();
                        int pos2 = pwebReply.IndexOf(']');
                        textBox11.Text = pwebReply.Substring(pos + 1, (pos2 - pos) - 1);
                        textBox11.Refresh();
                    }
                }
            }
            
            strx = textBox10.Text;
            stry = textBox11.Text;
            if (!string.IsNullOrEmpty(strx) && !string.IsNullOrEmpty(strx))
            {
                float x = float.Parse(strx);
                float y = float.Parse(stry);

                GPSgeofence gPSgeofence = new GPSgeofence();
                GPSgeofence fence = gPSgeofence;
                fence.ReadGPSfence();
                inTerritory = fence.PointInPolygon(x, y);

                if (inTerritory)
                    NewText = NewText + strx + " " + stry + " - location in Territory ";
                else
                {
                    if (x == form1.notfoundlat && y == form1.notfoundlon)
                        NewText = "\n GPS location not found ";
                    else
                        NewText = NewText + strx + " " + stry + " - location not in Territory - ";
                }

            }
            else
                NewText = "Address or Co-ordinates are blank!";
            textBox6.AppendText(Environment.NewLine + NewText);
            if (checkBox4.Checked)
            {
                textBox3_TextChanged(sender, e);
                f4.WindowState = FormWindowState.Normal;
                f4.BringToFront();
                f4.Focus();
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        { 
            if (checkBox3.Checked)
                button5.PerformClick();
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
                button5.PerformClick();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form1 f1 = (Form1)Application.OpenForms["Form1"];
            if (!string.IsNullOrEmpty(strx) && !string.IsNullOrEmpty(stry))
            {
                int pos = textBox6.Text.IndexOf("confidence is ");
                f1.confid = textBox6.Text.Substring(pos + 14, 1);
                float x = float.Parse(strx);
                float y = float.Parse(stry);
                if (x == f1.notfoundlat && y == f1.notfoundlon)
                {
                    f1.confid = "X";
                }
                else
                {
                    if (inTerritory) f1.resul = "In";
                    else f1.resul = "O";
                }

                if (f1.confid == "X") f1.lat = "N/A";
                else f1.lat = textBox10.Text;

                if (f1.confid == "X") f1.lon = "N/A";
                else f1.lon = textBox11.Text;
                f1.adjGPS = true;
            }
            else
                textBox6.AppendText("\n GPS NOT saved - Latitude or Longitude are blank");
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked && checkBox4.Focused) textBox3_TextChanged(sender, e);
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Validated(object sender, EventArgs e)
        {
            Form1 f1 = (Form1)Application.OpenForms["Form1"];
            GroupBox g = sender as GroupBox;
            var a = from RadioButton r in g.Controls where r.Checked == true select r.Text;
            f1.checkedRadio = a.First();
        }
    }
}
