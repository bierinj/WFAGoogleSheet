
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace WFAGoolgeSheet
{
    public partial class Form2 : Form
    {
        private Form1 form1;
        public Form2()
        {
            InitializeComponent();
        }

        bool preventExit = false;

        private void label4_Click(object sender, EventArgs e)
        { }

        //
        // OK button processing
        private void button1_Click(object sender, EventArgs e)
        {
            Form1.myVar = textBox6.Text;

            //Thread.Sleep(200);
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
            preventExit = false;
            this.Close();
        }

        //
        // Skip / Next - record processing
        private void button3_Click(object sender, EventArgs e)
        {
            preventExit = false;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        //
        // Skip back button
        //
        private void button4_Click(object sender, EventArgs e)
        {
            preventExit = false;
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.button4.BackColor = System.Drawing.Color.LightGray;
            form1.button4.Update();
        }

        //
        // map via Google
        //
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            // Get the Google Maps URL with defult zoom.
            string url = Program.GoogleMapUrl(textBox3.Text, "h", 2);

            // Display the URL in the default browser.
            Process.Start(url);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            if (!form1.checkBox4.Checked && checkBox2.Checked)
                button2.PerformClick();
        }
    }
}
