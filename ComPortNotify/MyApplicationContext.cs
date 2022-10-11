using System;
using System.Linq;
using System.IO.Ports;
using System.Windows.Forms;
using Microsoft.Win32;
using NAppUpdate.Framework;
using NAppUpdate.Framework.Common;
using NAppUpdate.Framework.Sources;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Collections.Generic;

namespace ComPortNotify
{
    class MyApplicationContext : Form
    {
        private const int WM_DEVICECHANGE = 0x219;
        private const int DBT_DEVICEARRIVAL = 0x8000;
        private const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
        private static string[] portarray = SerialPort.GetPortNames();
        private static List<ToolStripItem> menuComItems = new List<ToolStripItem>();
        private string[] new_portarray;
        private string added_port, removed_port;
        private string com_port;
        private bool new_com_port = false;
        private string terminalCommandProcess = "C:\\Program Files\\PuTTY\\putty.exe";
        private string terminalCommandArg = "-serial %1 -sercfg 115200,8,n,1,N";
        private NotifyIcon TrayIcon;
        private ContextMenuStrip TrayIconContextMenu;
        private System.ComponentModel.IContainer components;
        private ToolStripMenuItem CloseMenuItem, StartOnBoot, TerminalSetup, CheckForUpdates;

        public MyApplicationContext()
        {
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
            InitializeComponent();

           TrayIcon.Visible = true;
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.TrayIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.TrayIconContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.CloseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.StartOnBoot = new System.Windows.Forms.ToolStripMenuItem();
            this.TerminalSetup = new System.Windows.Forms.ToolStripMenuItem();
            this.CheckForUpdates = new System.Windows.Forms.ToolStripMenuItem();
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
                new System.Windows.Forms.ToolStripItem[] { this.CheckForUpdates });
            this.TrayIconContextMenu.Items.AddRange(
                new System.Windows.Forms.ToolStripItem[] { this.TerminalSetup });
            this.TrayIconContextMenu.Items.AddRange(
                new System.Windows.Forms.ToolStripItem[] { this.StartOnBoot });
            this.TrayIconContextMenu.Items.AddRange(
                new System.Windows.Forms.ToolStripItem[] { this.CloseMenuItem });
            this.TrayIconContextMenu.Items.Add(new ToolStripSeparator());
            // Read all COM port
            initPortMenu();

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
            // TerminalSetup
            //
            this.TerminalSetup.Name = "TerminalSetup";
            this.TerminalSetup.Size = new System.Drawing.Size(152, 22);
            this.TerminalSetup.Text = "Terminal Setup";
            this.TerminalSetup.Click += new System.EventHandler(this.TerminalSetup_Click);
            //
            // StartOnBoot
            //
            this.StartOnBoot.Name = "StartOnBoot";
            this.StartOnBoot.Size = new System.Drawing.Size(152, 22);
            this.StartOnBoot.Text = "Start COM Port Notifier on Boot";
            this.StartOnBoot.Click += new System.EventHandler(this.StartOnBoot_Click);
            //
            // CheckForUpdates
            //
            this.CheckForUpdates.Name = "CheckForUpdates";
            this.CheckForUpdates.Size = new System.Drawing.Size(152, 22);
            this.CheckForUpdates.Text = "Check for updates";
            this.CheckForUpdates.Click += new System.EventHandler(this.CheckForUpdates_Click);
            //
            // MyApplicationContext
            //
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "MyApplicationContext";
            this.Load += new System.EventHandler(this.MyApplicationContext_Load);
            this.TrayIconContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        private void initPortMenu()
        {
            foreach (ToolStripItem item in menuComItems)
            {
                TrayIconContextMenu.Items.Remove(item);
            }
            foreach (string port in portarray)
            {
                menuComItems.Add(this.TrayIconContextMenu.Items.Add(port, null, new System.EventHandler((sender, e) => PortItem_Click(sender, e, port))));
            }
        }
        void notifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            if(new_com_port)
                Process.Start(terminalCommandProcess, terminalCommandArg.Replace("%1", com_port));
            Clipboard.SetText(com_port);
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
                            {
                                TrayIcon.ShowBalloonTip(500, added_port, " plugged in.", ToolTipIcon.Info);
                                com_port = added_port;
                                new_com_port = true;
                                initPortMenu();
                            }
                            break;

                        case DBT_DEVICEREMOVECOMPLETE:
                            new_portarray = SerialPort.GetPortNames();
                            removed_port = portarray.Except(new_portarray).FirstOrDefault();
                            portarray = new_portarray;
                            if (removed_port != null)
                            {
                                TrayIcon.ShowBalloonTip(500, removed_port, "removed.", ToolTipIcon.Info);
                                com_port = removed_port;
                                new_com_port = false;
                                initPortMenu();
                            }
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
        private void TerminalSetup_Click(object sender, EventArgs e)
        {
            string inputProcess = Interaction.InputBox("Launcher command process", "Terminal Setup", terminalCommandProcess);
            if(!inputProcess.Equals(""))
                terminalCommandProcess = inputProcess;
            string inputArg = Interaction.InputBox("Launcher command arguments\n%1 = COMX", "Terminal Setup", terminalCommandArg);
            if (!inputArg.Equals(""))
                terminalCommandArg = inputArg;
        }
        private void PortItem_Click(object sender, EventArgs e, string port)
        {
            Process.Start(terminalCommandProcess, terminalCommandArg.Replace("%1", port));
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
        private void CheckForUpdates_Click(object sender, EventArgs e)
        {
			// Get a local pointer to the UpdateManager instance
			NAppUpdate.Framework.UpdateManager updManager = NAppUpdate.Framework.UpdateManager.Instance;

            if (updManager.State != NAppUpdate.Framework.UpdateManager.UpdateProcessState.NotChecked)
			{
				updManager.CleanUp();
			}

            updManager.CheckForUpdates();
            if (updManager.UpdatesAvailable != 0)
            {
                DialogResult dr = MessageBox.Show(
                    string.Format("Updates are available to your software ({0} total). Do you want to download and prepare them now? You can always do this at a later time.",
                    updManager.UpdatesAvailable),
                    "Software updates available",
                     MessageBoxButtons.YesNo);

                if (dr == DialogResult.Yes)
                {
                    //NAppUpdate.Framework.UpdateManager.Instance.PrepareUpdates();
                    updManager.BeginPrepareUpdates(OnPrepareUpdatesCompleted, null);
                }
            }
            else
            {
                MessageBox.Show("Your software is up to date");
            }
        }
        private void OnPrepareUpdatesCompleted(IAsyncResult asyncResult)
        {
			try
			{
				((UpdateProcessAsyncResult)asyncResult).EndInvoke();
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Updates preperation failed. Check the feed and try again.{0}{1}", Environment.NewLine, ex));
				return;
			}

			// Get a local pointer to the UpdateManager instance
			NAppUpdate.Framework.UpdateManager updManager = NAppUpdate.Framework.UpdateManager.Instance;

			DialogResult dr = MessageBox.Show("Updates are ready to install. Do you wish to install them now?", "Software updates ready", MessageBoxButtons.YesNo);

			if (dr != DialogResult.Yes)
			{
				return;
			}
			// This is a synchronous method by design, make sure to save all user work before calling
			// it as it might restart your application
			try
			{
				updManager.ApplyUpdates(true, false, false);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Error while trying to install software updates{0}{1}", Environment.NewLine, ex));
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
