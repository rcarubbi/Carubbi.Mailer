using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Carubbi.Mailer.Interfaces;
using Carubbi.Utils.Persistence;
using Microsoft.Office.Interop.Outlook;


namespace Carubbi.Mailer.Outlook2007
{
    public class OutlookInteropSender : OutlookInteropBase, IMailSender
    {
        public string Username { get; set; }
        public string Password { get; set; }

        #region IMailSender Members


        public void Send(System.Net.Mail.MailMessage message)
        {
            if (!OutlookIsRunning)
            {
                LaunchOutlook();
            }
            Recipients oRecips = null;
            Recipient oRecip = null;
            MailItem oMsg = null;

            try
            {
                _myApp = new Application();

                oMsg = (MailItem)_myApp.CreateItem(OlItemType.olMailItem);

                oMsg.HTMLBody = message.Body;
                oMsg.Subject = message.Subject;
                oRecips = (Recipients)oMsg.Recipients;


                foreach (var email in message.To)
                {
                    oRecip = (Recipient)oRecips.Add(email.Address);
                }

                foreach (var email in message.CC)
                {
                    oMsg.CC += string.Concat(email, ";");
                }

                List<string> filenames = Attach(message.Attachments, oMsg);
                oRecip.Resolve();
                (oMsg as _MailItem).Send();

                _mapiNameSpace = _myApp.GetNamespace("MAPI");
           

                DeleteTempFiles(filenames);
                Thread.Sleep(5000);
            }
            finally
            {
                if (oRecip != null) Marshal.ReleaseComObject(oRecip);
                if (oRecips != null) Marshal.ReleaseComObject(oRecips);
                if (oMsg != null) Marshal.ReleaseComObject(oMsg);
                if (_mapiNameSpace != null) Marshal.ReleaseComObject(_mapiNameSpace);
                if (_myApp != null) Marshal.ReleaseComObject(_myApp);
            }

        }

        private void DeleteTempFiles(List<string> filenames)
        {
            foreach (var file in filenames)
            {
                File.Delete(file);
            }
        }

        private static List<string> Attach(System.Net.Mail.AttachmentCollection attachments, MailItem oMsg)
        {
            int attachmentIndex = 0;
            List<string> filenames = new List<string>();
            foreach (var attachment in attachments)
            {
                filenames.Add(Path.GetTempFileName());
                File.WriteAllBytes(filenames.Last(), attachment.ContentStream.ToByteArray());
                oMsg.Attachments.Add(filenames.Last(), OlAttachmentType.olByValue, attachmentIndex++, attachment.Name);
            }

            return filenames;
        }


        #endregion

        public void Dispose()
        {
            // TODO: Implementar posteriormente de acordo com os padrões utilizando _disposing
            Dispose(true);
            GC.SuppressFinalize(this);
        }

          ~OutlookInteropSender()
        {
            Dispose(false);
        }


        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            { 
                
            }
        }




        public bool UseSSL
        {
            get
            {
                throw new NotSupportedException();
            }
            set
            {
                throw new NotSupportedException();
            }

        }

        public string Host
        {
            get
            {
                throw new NotSupportedException();
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public int PortNumber
        {
            get
            {
                throw new NotSupportedException();
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public bool UseDefaultCredentials
        {
            get
            {
                throw new NotSupportedException();
            }

            set
            {
                throw new NotSupportedException();
            }
        }
    }
}