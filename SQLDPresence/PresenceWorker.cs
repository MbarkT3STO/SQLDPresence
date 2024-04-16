using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

using MBDRPC.Core;
using MBDRPC.Helpers;

namespace SQLDPresence
{
    public class PresenceWorker
    {
        private Presence presence = new Presence();
        private bool isFirstRun = true;
        private DateTime startTime;
        private string processName;
        public Timer Timer;



        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, string lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetClassName(IntPtr hWnd, [Out] StringBuilder lpClassName, int nMaxCount);



        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
        {
            Timer = new Timer(_ => CheckSSMS(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckSSMS()
        {
            if (IsDiscordRunning())
            {
                if (IsSsmsRunning())
                {
                    if (isFirstRun)
                    {
                        presence.InitializePresence("1223365681077293156");
                        presence.UpdateLargeImage("mssql_logo", "Microsoft SQL Server Management Studio");
                        startTime = RunningAppChecker.GetProcessStartTime("ssms");
                        isFirstRun = false;
                    }

                    UpdatePresence();
                }
                else
                {
                    presence.ShutDown();
                    isFirstRun = true;
                }
            }
        }



        /// <summary>
        /// Checks if Discord is running
        /// </summary>
        private static bool IsDiscordRunning()
        {
            return Process.GetProcessesByName("discord").Any();
        }


        /// <summary>
        /// Checks if SSMS ( SQL Server Management Studio ) is running
        /// </summary>
        private static bool IsSsmsRunning()
        {
            return RunningAppChecker.IsAppRunning("ssms");
        }


        /// <summary>
        /// Checks if any Microsoft SQL Server Management Studio Window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft SQL Server Management Studio is running
            var processes = Process.GetProcessesByName("ssms")
                                   .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle));

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open queries/windows in Microsoft SQL Server Management Studio
        /// </summary>
        private static string[] GetSsmsOpenWindowNames()
        {
            // Retrieve the names of all open queries/windows in Microsoft SQL Server Management Studio
            var windowNames = Process.GetProcessesByName("ssms")
                                     .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle))
                                     .Select(p => p.MainWindowTitle)
                                     //.Select( p => p.MainWindowTitle.Replace( " - Excel" , "" ))
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft SQL Server Management Studio home/connection screen window is active
        /// </summary>
        private static bool IsConnectScreenActive()
        {
            var handle = FindWindow(null, "Connect to Server");
            return handle != IntPtr.Zero;
        }

        /// <summary>
        /// Checks if the Microsoft SQL Server Management Studio home/connection screen window is active
        /// </summary>
        private static bool IsConnectScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals("Connect to Server");
        }


        /// <summary>
        /// Checks if any Query Editor window is open and gets the name of the SQL query
        /// </summary>
        /// <returns></returns>
        private static (bool IsEditingQuery, bool IsQueryDesigner, string QueryName) IsUserEditingQuery()
        {
            // SSMS is running, check for the active window
            IntPtr foregroundWindow = GetForegroundWindow();

            if (foregroundWindow == IntPtr.Zero) return (false, false, string.Empty);

            // Get the window title
            const int nChars = 256;
            var windowTitle = new string(' ', nChars);
            GetWindowText(foregroundWindow, windowTitle, nChars);


            // Check if the window belongs to SSMS query editor
            if (windowTitle.Contains("Query Designer"))
            {
                return (true, true, string.Empty);
            }

            if (windowTitle.Contains(".sql"))
            {
                // Remove from ' - ' to the end from the window title
                var queryName = windowTitle.Substring(0, windowTitle.IndexOf(" - ", StringComparison.Ordinal));

                return (true, false, queryName);
            }

            return (false, false, string.Empty);

        }

        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any window is open
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetSsmsOpenWindowNames();
                var isConnectScreenActive = IsConnectScreenActive();
                var (isEditingQuery, isQueryDesigner, queryName) = IsUserEditingQuery();

                if (isConnectScreenActive)
                {
                    presence.UpdateDetails("Connect Screen");
                }
                else switch (isEditingQuery)
                    {
                        case true when isQueryDesigner:
                            presence.UpdateDetails("Query Designer");
                            break;
                        case true when !isQueryDesigner:
                            presence.UpdateDetails($"Editing: {queryName}");
                            break;
                        default:
                            presence.UpdateDetails(string.Empty);
                            break;
                    }
            }
            else
            {
                presence.UpdateDetails("Home Screen");
            }

            UpdatePresenceTime();
            presence.UpdatePresence();
        }


        /// <summary>
        /// Updates the presence time
        /// </summary>
        private void UpdatePresenceTime()
        {
            var elapsedTime = (DateTime.Now - startTime).ToString(@"hh\:mm\:ss");
            presence.UpdateState(elapsedTime);
        }
    }
}