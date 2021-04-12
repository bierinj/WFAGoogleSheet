namespace WFAGoolgeSheet
{
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// Defines the <see cref="FormMediator" />.
    /// </summary>
    public class FormMediator
    {
        /// <summary>
        /// Defines the _subForm.
        /// </summary>
        private Form4 _subForm;

        /// <summary>
        /// Defines the _mainForm.
        /// </summary>
        private Form1 _mainForm;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormMediator"/> class.
        /// </summary>
        /// <param name="mainForm">The mainForm<see cref="Form1"/>.</param>
        /// <param name="subForm">The subForm<see cref="Form4"/>.</param>
        public FormMediator(Form1 mainForm, Form4 subForm)
        {
            if (mainForm == null)
                throw new ArgumentNullException("mainForm");

            if (subForm == null)
                throw new ArgumentNullException("subForm");

            _mainForm = mainForm;
            _subForm = subForm;

            _subForm.FormClosed += MainForm_FormClosed;
        }

        /// <summary>
        /// The MainForm_FormClosed.
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/>.</param>
        /// <param name="e">The e<see cref="FormClosedEventArgs"/>.</param>
        internal void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                _mainForm.Refresh();
            }
            catch (NullReferenceException ex)
            {
                throw new InvalidOperationException("Unable to close the Main Form because the FormMediator no longer has a reference to it.", ex);
            }
        }
    }
}
