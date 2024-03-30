using DiscordRpcDemo;
using System;
using System.Linq;
using System.Windows.Forms;
using Process = System.Diagnostics.Process;


namespace SQLDPresence
{
    public partial class Form1 : Form
    {
        private DiscordRpc.EventHandlers handlers = default;
        private DiscordRpc.RichPresence presence;

        private DateTime startTime      = DateTime.Now;
        private bool     isSsmsFirstRun = true;


        public Form1()
        {
            InitializeComponent();
            Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true).SetValue("MBVRK.SsmsRpc", Application.ExecutablePath);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            CheckSsms();
        }


        /// <summary>
        /// Initializes the presence
        /// </summary>
        private void InitializePresence()
        {
            DiscordRpc.Initialize("1223365681077293156", ref handlers, true, null);
            presence = new DiscordRpc.RichPresence();
        }


        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            UpdatePresenceTime();
            DiscordRpc.UpdatePresence(ref presence);
        }


        /// <summary>
        /// Sets the presence image and text ( large image and details )
        /// </summary>
        private void SetPresenceImageAndText()
        {
            presence.largeImageText = "SQL Server";
            presence.largeImageKey  = "sql_server_logo";

            const string ssmsFullName = "SQL Server Management Studio";

            presence.largeImageText = ssmsFullName;
            presence.details        = ssmsFullName;
        }


        /// <summary>
        /// Updates the presence time
        /// </summary>
        private void UpdatePresenceTime()
        {
            presence.state = (DateTime.Now - startTime).ToString(@"hh\:mm\:ss");
        }

        /// <summary>
        /// Checks if SSMS ( SQL Server Management Studio ) is running
        /// </summary>
        private void CheckSsms()
        {
            var isSsmsOpen = IsSsmsRunning();
            if (isSsmsOpen)
            {
                if ( isSsmsFirstRun )
                {
                    InitializePresence();
                    SetPresenceImageAndText();

                    startTime      = DateTime.Now;
                    isSsmsFirstRun = false;
                }

                UpdatePresence();
            }
            else
            {
                DiscordRpc.Shutdown();
                isSsmsFirstRun = true;
            }

        }


        /// <summary>
        /// Checks if SSMS ( SQL Server Management Studio ) is running
        /// </summary>
        /// <returns>
        /// <c>true</c> if SSMS is running; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsSsmsRunning()
        {
            return Process.GetProcesses().Any(process => process.ProcessName.Equals("Ssms", StringComparison.OrdinalIgnoreCase));
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            CheckSsms();
        }
    }
}
