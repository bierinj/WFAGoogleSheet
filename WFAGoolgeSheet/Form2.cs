
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WFAGoolgeSheet
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
        }

        bool preventExit = false;

        //
        // OK button processing
        private void button1_Click(object sender, EventArgs e)
        {
            Form1.myVar = textBox6.Text;
            this.Close();
        }

        private void Form2_FormClosing(Object sender, FormClosingEventArgs e)
        {
            if (preventExit) e.Cancel = true;
            else e.Cancel = false;
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
                Form1 form1 = (Form1)Application.OpenForms["Form1"];
                Form2 f2 = (Form2)Application.OpenForms["Form2"];
                Form4 f4 = (Form4)Application.OpenForms["Form4"];

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
                Program.formisup = false;
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
            Form1 form1 = new Form1();
            form1.button4.BackColor = System.Drawing.Color.LightGray;
            form1.button4.Update();
            if (checkBox3.Checked)
                button5.PerformClick();
            if (checkBox4.Checked)
                textBox3_TextChanged( sender, e);
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
                if(vs[0] == "X")
                {
                    //textBox6.AppendText(" GPS location not found");
                    return;
                }
                lat = Convert.ToDouble(vs[2]);
                lon = Convert.ToDouble(vs[3]);
                form4.LoadIntoMap(lat, lon);

                /*DialogResult dr =*/
                form4.Show();        // bring up the form
           
            }
            else
                textBox6.AppendText(" Address Is Blank");
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
            bool inTerritory = false;
            string strx = "";
            string stry = "";
            int pos = -1;
            int pos1 = -1;
            string NewText = "";
            HttpWebRequestHandler hTTPrequest = new HttpWebRequestHandler();
            //http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=Guamani&postalCode=-&addressLine={addressLine}&userLocation=-&userIp={-}&usermapView={usermapView}&includeNeighborhood=includeNeighborhood&maxResults={maxResults}&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS
            string webAdr = @"http://dev.virtualearth.net/REST/v1/Locations?countryRegion=Ecuador&adminDistrict=Quito&locality=-&postalCode=-&addressLine="+ textBox3.Text +"&key=AhbjdGZqctwmlxK6GXWgkfE5CL7J2c5OWuTCk7WaAy-xVXphOgT2_AWrLL-L90OS";
            if(!string.IsNullOrEmpty(textBox3.Text))
            {
                var webReply = hTTPrequest.GetReply(webAdr);
                pos = webReply.IndexOf("\"coordinates\":");
                pos1 = webReply.IndexOf("\"confidence\":");
                if (pos1 > -1)
                    NewText = "confidence is " + webReply.Substring(pos1+14, 6) +"  ";
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
            
            pos = -1;
            pos = textBox12.Text.IndexOf(",");
            if (!String.IsNullOrEmpty(textBox12.Text)&&(pos != -1))
            {
                strx = textBox12.Text.Substring(0, pos - 1);
                stry = textBox12.Text.Substring(pos+1, (textBox12.Text.Length-pos)-1);
                textBox10.Text = strx;
                textBox10.Refresh();
                textBox11.Text = stry;
                textBox11.Refresh();
                textBox12.Text = "";
                textBox12.Refresh();
            }
            else
            {
                strx = textBox10.Text;
                stry = textBox11.Text;
            }

            if (!string.IsNullOrEmpty(strx) || !string.IsNullOrEmpty(strx))
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
                    if(x == form1.notfoundlat && y == form1.notfoundlon)
                        NewText = "\n GPS location not found ";
                    else
                        NewText = NewText + strx + " " + stry + " - location not in Territory - ";
                }

            }
            else
                NewText = "Address or Co-ordinates are blank!";
            textBox6.AppendText(Environment.NewLine + NewText);
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
    }
}
