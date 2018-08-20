using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.Linq;

namespace Carubbi.Mailer.Outlook2007
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

        protected bool OutlookIsRunning { get; } = Process.GetProcesses().Any(otlk => otlk.ProcessName.ToLower().Contains(OUTLOOK_PROCESS_PATTERN));

        protected void LaunchOutlook()
        {
            var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(OUTLOOK_REGISTRY_KEY);
            if (key == null) return;
            var path = (string)key.GetValue(OUTLOOK_REGISTRY_VALUE);
            if (path != null)
            {
                var p = Process.Start(OUTLOOK_EXE_FILENAME);
                p?.WaitForInputIdle();
            }
            else
                throw new ApplicationException(OUTLOOK_NOT_FOUND_MESSAGE);
        }
    }
}
