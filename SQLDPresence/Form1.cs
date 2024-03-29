using DiscordRpcDemo;

using EnvDTE;

using EnvDTE80;



using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Process = System.Diagnostics.Process;

namespace SQLDPresence
{
    public partial class Form1 : Form
    {
        private DiscordRpc.EventHandlers handlers;
        private DiscordRpc.RichPresence presence;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            CheckSsms();
        }

        private void InitializePresence()
        {
            handlers = default(DiscordRpc.EventHandlers);
            DiscordRpc.Initialize("1223365681077293156", ref handlers, true, null);
            handlers = default(DiscordRpc.EventHandlers);
            DiscordRpc.Initialize("1223365681077293156", ref handlers, true, null);
            //this.presence.details = "SQL Server";
            presence.state = "Running";
            presence.largeImageKey = "sql_server_logo";
            //presence.smallImageKey = "8422322";
            presence.largeImageText = "SQL Server";
            //presence.smallImageText = "SQL File or Query";
            //DiscordRpc.UpdatePresence(ref presence);
        }

        private void CheckSsms()
        {
            bool isSsmsOpen = IsSsmsRunning();
            if (isSsmsOpen)
            {
                InitializePresence();

                DTE2 dte = GetDte();
                if (dte != null)
                {
                    Window activeWindow = dte.ActiveWindow;
                    if (activeWindow != null && activeWindow.Type == vsWindowType.vsWindowTypeDocument)
                    {
                        Document activeDocument = activeWindow.Document;
                        if (activeDocument != null)
                        {
                            this.presence.smallImageKey = "8422322";

                            if (activeDocument.FullName.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                            {
                                // User is inside a SQL file
                                this.presence.details = "Editing SQL: " + activeDocument.Name;
                            }
                            else if (activeDocument.FullName.EndsWith(".query", StringComparison.OrdinalIgnoreCase))
                            {
                                // User is inside a query file
                                this.presence.details = "Editing Query: " + activeDocument.Name;
                            }
                        }
                    }
                }
            }
            else
            {
                DiscordRpc.Shutdown();
            }
            DiscordRpc.UpdatePresence(ref this.presence);
        }



        bool IsSsmsRunning()
        {
            Process[] processes = Process.GetProcessesByName("Ssms");
            return processes.Length > 0;
        }

        private DTE2 GetDte()
        {
            //DTE2 dte = null;
            //Type typeDTE = Type.GetTypeFromProgID("VisualStudio.DTE.17.0", true);
            //object objDTE = Activator.CreateInstance(typeDTE, true);
            //dte = (DTE2)objDTE;
            //return dte;


            object runningObject = null;
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;
            IMoniker[] moniker = new IMoniker[1];
            IntPtr pfetched = IntPtr.Zero;

            GetRunningObjectTable(0, out rot);
            rot.EnumRunning(out enumMoniker);

            while (enumMoniker.Next(1, moniker, pfetched) == 0)
            {
                IMoniker runningMoniker = moniker[0];
                string name = null;

                if (runningMoniker != null)
                {
                    runningMoniker.GetDisplayName(null, null, out name);

                    if (name.Contains("VisualStudio.DTE"))
                    {
                        rot.GetObject(runningMoniker, out runningObject);
                        break;
                    }
                }
            }

            return runningObject as DTE2;
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

        private void timer1_Tick(object sender, EventArgs e)
        {
            CheckSsms();
        }
    }
}
