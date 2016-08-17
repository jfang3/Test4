using System;
using System.Collections.Generic;
using System.Linq;
// using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Application.EnableVisualStyles();
            // Application.SetCompatibleTextRenderingDefault(false);
            // Application.Run(new OptimizationDemo());
            Application.Run(new Form1());
            // Application.Run(new Login());
            /*
            var exitCode = 0;
            Mutex mutex = new Mutex(false, @"Global\"+appGuid);
            try
            {
                if (!mutex.WaitOne(0, false))
                {
                    MessageBox.Show("Instance already running"); return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
                GC.KeepAlive(mutex);
            }
            catch(SystemException systemEx)
            {
                exitCode = -systemEx.HResult; MessageBox.Show(exitCode.ToString());
            }
            catch(Exception ex)
            {
                Environment.ExitCode = int.MinValue;
                MessageBox.Show(Environment.ExitCode.ToString());
            }
        }
        private static string appGuid = "c0a76b5a-12ab-45c5-b9d9-d693faa6e7b9";*/
        }
    }
}
