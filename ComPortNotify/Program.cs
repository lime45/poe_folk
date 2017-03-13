using System;
using System.Windows.Forms;

namespace ComPortNotify
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            NAppUpdate.Framework.UpdateManager.Instance.UpdateSource = new NAppUpdate.Framework.Sources.SimpleWebSource("http://www.tylercrumpton.com/Random/poe/feed.xml");
            NAppUpdate.Framework.UpdateManager.Instance.ReinstateIfRestarted();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MyApplicationContext());
        }
    }
}
