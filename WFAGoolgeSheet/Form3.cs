using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WFAGoolgeSheet
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();

            this.Font = System.Drawing.SystemFonts.IconTitleFont;
            SystemEvents.UserPreferenceChanged += new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
            this.FormClosing += new FormClosingEventHandler(Form3_FormClosing);
        }

        void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Window)
            {
                this.Font = System.Drawing.SystemFonts.IconTitleFont;
            }
        }

        void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
        }

        public bool setMinstate = false;
        //
        // change default format for datepicker
        //
        //public void SetMyCustomFormat()
        //{
        //    // Set the Format type and the CustomFormat string.
        //    monthCalendar1.Format = DateTimePickerFormat.Custom;
        //    monthCalendar1.CustomFormat = "MM/yyyy";
        //}
        //
        // clear datePicker
        //
        private void button3_Click(object sender, EventArgs e)
        {
            setMinstate = true;
            monthCalendar1.SelectionStart = DateTime.Now;
            monthCalendar1.SelectionEnd = DateTime.Now;
        }
        // hide if set to MinimumDateTime
        private void monthCalendar1_ValueChanged(object sender, EventArgs e)
        {
            setMinstate = false;
        }
        //
        // Clear AM/PM
        //
        private void button4_Click(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }

        //
        // Clear Day of Week
        //
        private void button5_Click(object sender, EventArgs e)
        {
            foreach (int i in checkedListBox2.CheckedIndices)
            {
                checkedListBox2.SetItemCheckState(i, CheckState.Unchecked);
            }
        }

        //
        // Invert Day of Week
        //
        private void button6_Click(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (checkedListBox2.GetItemChecked(i))
                    checkedListBox2.SetItemChecked(i, false);
                else checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
