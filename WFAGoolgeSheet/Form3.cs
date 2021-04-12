namespace WFAGoolgeSheet
{
    using Microsoft.Win32;
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="Form3" />.
    /// </summary>
    public partial class Form3 : Form
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Form3"/> class.
        /// </summary>
        public Form3()
        {
            InitializeComponent();

            this.Font = System.Drawing.SystemFonts.IconTitleFont;
            SystemEvents.UserPreferenceChanged += new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
            this.FormClosing += new FormClosingEventHandler(Form3_FormClosing);
        }

        /// <summary>
        /// The SystemEvents_UserPreferenceChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="UserPreferenceChangedEventArgs"/>.</param>
        internal void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Window)
            {
                this.Font = System.Drawing.SystemFonts.IconTitleFont;
            }
        }

        /// <summary>
        /// The Form3_FormClosing.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="FormClosingEventArgs"/>.</param>
        internal void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            SystemEvents.UserPreferenceChanged -= new UserPreferenceChangedEventHandler(SystemEvents_UserPreferenceChanged);
        }

        /// <summary>
        /// Defines the setMinstate.
        /// </summary>
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
        /// <summary>
        /// The button3_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button3_Click(object sender, EventArgs e)
        {
            setMinstate = true;
            monthCalendar1.SelectionStart = DateTime.Now;
            monthCalendar1.SelectionEnd = DateTime.Now;
        }

        // hide if set to MinimumDateTime
        /// <summary>
        /// The monthCalendar1_ValueChanged.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void monthCalendar1_ValueChanged(object sender, EventArgs e)
        {
            setMinstate = false;
        }

        //
        // Clear AM/PM
        //
        /// <summary>
        /// The button4_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button4_Click(object sender, EventArgs e)
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
        }

        //
        // Clear Day of Week
        //
        /// <summary>
        /// The button5_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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
        /// <summary>
        /// The button6_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
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

        /// <summary>
        /// The button1_Click.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="EventArgs"/>.</param>
        private void button1_Click(object sender, EventArgs e)
        {
        }
    }
}
