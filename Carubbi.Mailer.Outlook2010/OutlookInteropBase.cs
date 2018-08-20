using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Threading;

namespace Carubbi.Mailer.Outlook2010
{
    public abstract class OutlookInteropBase
    {
        private const string OUTLOOK_EXE_FILENAME = "OUTLOOK.EXE";
        private const string OUTLOOK_PROCESS_PATTERN = "outlook";
        private const string OUTLOOK_REGISTRY_KEY = "Software\\microsoft\\windows\\currentversion\\app paths\\OUTLOOK.EXE";
        private const string OUTLOOK_REGISTRY_VALUE = "Path";
        private const string OUTLOOK_NOT_FOUND_MESSAGE = "O Outlook não foi encontrado nesta máquina. Favor instalar o aplicativo";
        protected Application _myApp;
        protected NameSpace _mapiNameSpace;
        protected MAPIFolder _mapiFolder;

       
        protected bool OutlookIsRunning
        {
            get
            {
                foreach (Process otlk in Process.GetProcesses())
                {
                    if (otlk.ProcessName.ToLower().Contains(OUTLOOK_PROCESS_PATTERN))
                    {
                        return true;
                    }
                }
                return false;
            }
        }


        protected void LaunchOutlook()
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(OUTLOOK_REGISTRY_KEY);
            string path = (string)key.GetValue(OUTLOOK_REGISTRY_VALUE);
            if (path != null)
            {
                var p = Process.Start(OUTLOOK_EXE_FILENAME);
                p.WaitForInputIdle();
            }
            else
                throw new ApplicationException(OUTLOOK_NOT_FOUND_MESSAGE);
        }
    }
}
