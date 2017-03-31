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
			NAppUpdate.Framework.UpdateManager.Instance.UpdateSource = new NAppUpdate.Framework.Sources.SimpleWebSource("https://www.dropbox.com/s/4czsu1rljsj7lec/feed.xml?dl=1");
			NAppUpdate.Framework.UpdateManager.Instance.Config.UpdateExecutableName = "COM Port Notifier Updater.exe";
			NAppUpdate.Framework.UpdateManager.Instance.Config.UpdateProcessName = "COM Port Notifier Updater";
			NAppUpdate.Framework.UpdateManager.Instance.ReinstateIfRestarted();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MyApplicationContext());
		}
    }
}
