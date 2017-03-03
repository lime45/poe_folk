using System;
using System.Drawing;
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
        private ToolStripMenuItem CloseMenuItem;

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
            TrayIcon = new NotifyIcon();

            TrayIcon.BalloonTipIcon = ToolTipIcon.Info;
            TrayIcon.BalloonTipTitle = "COM Port Notifier";
            TrayIcon.Text = "COM Port Notifier";


            //The icon is added to the project resources.
            TrayIcon.Icon = Properties.Resources.serial_port;

            //Add a context menu to the TrayIcon:
            TrayIconContextMenu = new ContextMenuStrip();
            CloseMenuItem = new ToolStripMenuItem();
            TrayIconContextMenu.SuspendLayout();

            // 
            // TrayIconContextMenu
            // 
            this.TrayIconContextMenu.Items.AddRange(new ToolStripItem[] {
            this.CloseMenuItem});
            this.TrayIconContextMenu.Name = "TrayIconContextMenu";
            this.TrayIconContextMenu.Size = new Size(153, 70);
            // 
            // CloseMenuItem
            // 
            this.CloseMenuItem.Name = "CloseMenuItem";
            this.CloseMenuItem.Size = new Size(152, 22);
            this.CloseMenuItem.Text = "Close COM Port Notifier";
            this.CloseMenuItem.Click += new EventHandler(this.CloseMenuItem_Click);

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

        private void CloseMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you really want to exit?",
                    "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
