using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SQLDPresence
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        public static extern bool FreeConsole();


        static void Main(string[] args)
        {
            var worker = new PresenceWorker();

            try
            {
                // Hide the console window
                FreeConsole();

                // Register the app to be auto startup
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    key?.SetValue("MBVRK.SQLServerDRPC", System.Reflection.Assembly.GetExecutingAssembly().Location);
                }


                worker.Start();

                // sleep thread forever
                Thread.Sleep(Timeout.Infinite);
            }
            catch (Exception)
            {
                // Dispose the timer
                worker.Timer.Dispose();

                // Kill the application
                Environment.Exit(0);
            }
        }
    }
}