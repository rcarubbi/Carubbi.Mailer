using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Carubbi.Mailer.DTOs;
using Carubbi.Mailer.Interfaces;
using Microsoft.Office.Interop.Outlook;
using System.Threading;

namespace Carubbi.Mailer.Outlook2007
{
    public class OutlookInteropReceiver : OutlookInteropBase, IMailReceiver
    {
        public string Username { get; set; }
        public string Password { get; set; }

        private void InitializeObjects()
        {
            _myApp = new Microsoft.Office.Interop.Outlook.Application();
            _mapiNameSpace = _myApp.GetNamespace("MAPI");
            _mapiFolder = _mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            
        } 

      
        private void DisposeObjects()
        {
            GC.ReRegisterForFinalize(_mapiFolder);
            GC.ReRegisterForFinalize(_mapiNameSpace);
            GC.ReRegisterForFinalize(_myApp);

           
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(_mapiFolder);
            Marshal.FinalReleaseComObject(_mapiNameSpace);
            Marshal.FinalReleaseComObject(_myApp);
            _mapiFolder = null;
            _mapiNameSpace = null;
            _myApp = null;

        }

        #region Membros de IMailReceiver

     

        public IEnumerable<MailMessage> GetPendingMessages()
        {
            int readMails = 0;
            if (!OutlookIsRunning)
            {
                LaunchOutlook();
            }
            int emailsCount = 0;
            InitializeObjects();
            
            do
            {

                var items = _mapiFolder.Items;
                emailsCount = items.Count;

                foreach (Object it in items)
                {
                    if (it is MailItem)
                    {
                        MailItem item = (MailItem)it;
                        readMails++;
                        yield return ParseMessage(item);
                        if (OnMessageRead != null)
                        {
                            OnMessageReadEventArgs e = new OnMessageReadEventArgs();
                            OnMessageRead(this, e);
                            if (e.Cancel)
                                continue;
                        }
                        item.Delete();
                        Thread.Sleep(1000);
                        if (readMails == 10)
                            break;
                    }
                }

            } while (emailsCount > 0 && (readMails > 0 && readMails < 10));

            DisposeObjects();
            
        }

        public int GetPendingMessagesCount()
        {
            InitializeObjects();
            int count = _mapiFolder.Items.Count;

            DisposeObjects();
            return count > 10 ? 10 : count;
        }

        #endregion
 
        private MailMessage ParseMessage(MailItem item)
        {
            
            var mailMessage = new MailMessage(item.SenderEmailAddress, _myApp.Session.CurrentUser.Address, item.Subject, item.Body);

            if (item.Attachments.Count > 0)
            {
                System.Exception lastException = null;
                foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in item.Attachments)
                {
                    try
                    {
                        var tempFile = Path.GetTempFileName();
                        attachment.SaveAsFile(tempFile);
                        MemoryStream ms = new MemoryStream(File.ReadAllBytes(tempFile));
                        mailMessage.Attachments.Add(new System.Net.Mail.Attachment(ms, attachment.FileName));
                        File.Delete(tempFile);
                    }
                    catch (System.Exception ex)
                    {
                        lastException = ex;
                    }
                }

                if (lastException != null)
                    throw lastException;
            }

            return mailMessage;
        }

        public event EventHandler<OnMessageReadEventArgs> OnMessageRead;


        public void Dispose()
        {
            // TODO: Implementar posteriormente de acordo com os padrões utilizando _disposing
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OutlookInteropReceiver()
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
    }
}
