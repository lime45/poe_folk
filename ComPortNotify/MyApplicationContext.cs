using System;
using System.Linq;
using System.IO.Ports;
using System.Windows.Forms;

namespace ComPortNotify
{
    class MyApplicationContext : Form
    {
        private const int WM_DEVICECHANGE = 0x219;
        private const int DBT_DEVICEARRIVAL = 0x8000;
        private const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
        private static string[] portarray = SerialPort.GetPortNames();
        private string[] new_portarray;
        private string added_port, removed_port;
        private NotifyIcon TrayIcon;
        private ContextMenuStrip TrayIconContextMenu;
        private System.ComponentModel.IContainer components;
        private ToolStripMenuItem CloseMenuItem, StartOnBoot;

        public MyApplicationContext()
        {
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
            InitializeComponent();

            if (MessageBox.Show("Start Monitoring for Serial Port Events?", 
                    "Com Port Notifier", MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question) == DialogResult.Yes)
            {
                TrayIcon.Visible = true;
            }
            else
            {
                Environment.Exit(0);
            }

        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.TrayIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.TrayIconContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.CloseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.StartOnBoot = new System.Windows.Forms.ToolStripMenuItem();
            this.TrayIconContextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // TrayIcon
            // 
            this.TrayIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.TrayIcon.BalloonTipTitle = "COM Port Notifier";
            this.TrayIcon.BalloonTipClicked += new EventHandler(notifyIcon_BalloonTipClicked);

            this.TrayIcon.ContextMenuStrip = this.TrayIconContextMenu;
            this.TrayIcon.Icon = global::ComPortNotify.Properties.Resources.serial_port;
            this.TrayIcon.Text = "COM Port Notifier";
            // 
            // TrayIconContextMenu
            // 
            this.TrayIconContextMenu.Items.AddRange(
                    new System.Windows.Forms.ToolStripItem[] { this.CloseMenuItem });
            this.TrayIconContextMenu.Items.AddRange(
                    new System.Windows.Forms.ToolStripItem[] { this.StartOnBoot });
            this.TrayIconContextMenu.Name = "TrayIconContextMenu";
            this.TrayIconContextMenu.Size = new System.Drawing.Size(153, 70);
            // 
            // CloseMenuItem
            // 
            this.CloseMenuItem.Name = "CloseMenuItem";
            this.CloseMenuItem.Size = new System.Drawing.Size(152, 22);
            this.CloseMenuItem.Text = "Close COM Port Notifier";
            this.CloseMenuItem.Click += new System.EventHandler(this.CloseMenuItem_Click);
            // 
            // StartOnBoot
            // 
            this.StartOnBoot.Name = "StartOnBoot";
            this.StartOnBoot.Size = new System.Drawing.Size(152, 22);
            this.StartOnBoot.Text = "Start COM Port Notifier on Boot";
            this.StartOnBoot.Click += new System.EventHandler(this.StartOnBoot_Click);
            // 
            // MyApplicationContext
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "MyApplicationContext";
            this.Load += new System.EventHandler(this.MyApplicationContext_Load);
            this.TrayIconContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

            TrayIconContextMenu.ResumeLayout(false);
            TrayIcon.ContextMenuStrip = TrayIconContextMenu;
        }
        protected override void SetVisibleCore(bool value)
        {
            if (!this.IsHandleCreated)
            {
                value = false;
                CreateHandle();
            }
            base.SetVisibleCore(value);
        }
        // override form message handler
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            switch (m.Msg)
            {
                case WM_DEVICECHANGE:
                    switch ((int)m.WParam)
                    {
                        case DBT_DEVICEARRIVAL:
                            new_portarray = SerialPort.GetPortNames();
                            added_port = new_portarray.Except(portarray).FirstOrDefault();
                            portarray = new_portarray;
                            if (added_port != null)
                                TrayIcon.ShowBalloonTip(500, added_port, " plugged in.", ToolTipIcon.Info);
                            break;

                        case DBT_DEVICEREMOVECOMPLETE:
                            new_portarray = SerialPort.GetPortNames();
                            removed_port = portarray.Except(new_portarray).FirstOrDefault();
                            portarray = new_portarray;
                            if (removed_port != null)
                                TrayIcon.ShowBalloonTip(500, removed_port, "removed.", ToolTipIcon.Info);
                            break;
                    }
                    break;
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            //Cleanup so that the icon will be removed when the application is closed
            TrayIcon.Visible = false;
        }

        private void MyApplicationContext_Load(object sender, EventArgs e)
        {

        }
        private void StartOnBoot_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Start Notifier on startup?",
                    "Start Notifier on startup?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                RegisterInStartup(true);
            }
            else
            {
                RegisterInStartup(false);
            }
        }
         private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you really want to exit?",
                    "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        private void RegisterInStartup(bool wantsTo)
        {
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey
                    ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (wantsTo)
            {
                registryKey.SetValue("ComPortNotifier", Application.ExecutablePath);
            }
            else
            {
                registryKey.DeleteValue("ComPortNotifier");
            }
        }

    }
}
