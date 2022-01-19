using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZuluAnalyzer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            using (Mutex mutex = new Mutex(false, Constant.APP_GUID))
            {
                if (!mutex.WaitOne(0, false))
                {
                    MessageBox.Show("Already Existing!");
                    return;
                }
                System.IO.Directory.CreateDirectory("Downloads");
                System.IO.Directory.CreateDirectory("Logs");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
        }
    }
}
