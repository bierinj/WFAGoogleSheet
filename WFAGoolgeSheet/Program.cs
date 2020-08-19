using System;
using System.Collections.Generic;
using System.Windows.Forms;


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
        private static readonly char[] OpenParentheses = { '+','(', '[', '{' };
        private static readonly char[] CloseParentheses = { '+',')', ']', '}' };

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
    }
}


