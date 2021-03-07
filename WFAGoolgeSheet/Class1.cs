using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WFAGoolgeSheet
{
    public class FormMediator
    {
        private Form4 _subForm;
        private Form1 _mainForm;

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

        void MainForm_FormClosed(object sender, FormClosedEventArgs e)
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
