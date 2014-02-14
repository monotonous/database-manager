/**
 * Program.cs
 * Author: Joshua Parker
 * 
 * The main program file.
 * Shows the splash screen then the main form.
 */

using System;
using System.Windows.Forms;

namespace CS280A2
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
            Application.SetCompatibleTextRenderingDefault(false);
           
            // shows splash screen
            new SplashForm().ShowDialog();
           
            Application.Run(new Form1());
        }
    }
}
