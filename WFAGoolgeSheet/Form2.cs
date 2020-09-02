
using System;
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
        {}

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

        }
    }
}
