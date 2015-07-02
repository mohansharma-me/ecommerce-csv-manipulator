using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CSVManip
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                Exception excep = e.ExceptionObject as Exception;
                if (excep != null)
                {
                    try
                    {
                        MessageBox.Show("Unhandeled error occured, please try agian." + Environment.NewLine + "Error message:" + excep, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.IO.File.AppendAllText("logout.log", Environment.NewLine + "Unhandeled error:" + excep + Environment.NewLine);
                    }
                    catch (Exception) { }
                }
            };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmMain());
        }
    }
}
