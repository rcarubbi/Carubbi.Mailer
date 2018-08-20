using Carubbi.Mailer.DTOs;
using Carubbi.Mailer.Interfaces;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Threading;
using Carubbi.Extensions;
using Attachment = Microsoft.Office.Interop.Outlook.Attachment;
using Exception = System.Exception;

namespace Carubbi.Mailer.Outlook2010
{
    public class OutlookInteropReceiver : OutlookInteropBase, IMailReceiver
    {
        public string Username { get; set; }
        public string Password { get; set; }

        private void InitializeObjects()
        {
            MyApp = new Application();
            MapiNameSpace = MyApp.GetNamespace("MAPI");
            MapiFolder = MapiNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

        }

        private void DisposeObjects()
        {
            GC.ReRegisterForFinalize(MapiFolder);
            GC.ReRegisterForFinalize(MapiNameSpace);
            GC.ReRegisterForFinalize(MyApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(MapiFolder);
            Marshal.FinalReleaseComObject(MapiNameSpace);
            Marshal.FinalReleaseComObject(MyApp);
            MapiFolder = null;
            MapiNameSpace = null;
            MyApp = null;
        }

        #region Membros de IMailReceiver

        public string Address
        {
            get
            {
                if (!string.IsNullOrEmpty(_address)) return _address;

                _address = MyApp.Session.CurrentUser.Address;

                if (!_address.IsValidEmail())
                {
                    _address = GetAddress(MyApp.Session.CurrentUser.AddressEntry);
                }
                return _address;
            }

        }

        private string _address = string.Empty;

        public IEnumerable<MailMessage> GetPendingMessages()
        {
            var readMails = 0;

            if (!OutlookIsRunning)
            {
                LaunchOutlook();
            }

            var emailsCount = 0;

            InitializeObjects();

            do
            {
                var items = MapiFolder.Items;
                emailsCount = items.Count;

                foreach (var it in items)
                {
                    if (!(it is MailItem item)) continue;
                    if (!GetSenderSMTPAddress(item).IsValidEmail())
                        continue;

                    readMails++;
                    yield return ParseMessage(item);

                    if (OnMessageRead != null)
                    {
                        var e = new OnMessageReadEventArgs();
                        OnMessageRead(this, e);
                        if (e.Cancel)
                            continue;
                    }

                    item.Delete();
                    Thread.Sleep(1000);
                    if (readMails == 10)
                        break;
                }
            } while (emailsCount > 0 && (readMails > 0 && readMails < 10));

            DisposeObjects();
        }

        private string GetAddress(AddressEntry ae)
        {
            const string PR_SMTP_ADDRESS = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            if (ae == null) return null;

            //Now we have an AddressEntry representing the Sender
            if (ae.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                || ae.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                //Use the ExchangeUser object PrimarySMTPAddress
                var exchUser = ae.GetExchangeUser();
                return exchUser?.PrimarySmtpAddress;
            }

            try
            {
                return ae.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetSenderSMTPAddress(_MailItem mail)
        {
            if (mail == null)
            {
                throw new ArgumentNullException();
            }

            if (mail.SenderEmailType != "EX") return mail.SenderEmailAddress;

            var sender = mail.Sender;

            return GetAddress(sender);
        }

        public int GetPendingMessagesCount()
        {
            InitializeObjects();

            var count = MapiFolder.Items.Count;

            DisposeObjects();

            return count > 10 ? 10 : count;
        }

        #endregion

        private MailMessage ParseMessage(MailItem item)
        {
            var mailMessage = new MailMessage(GetSenderSMTPAddress(item), Address, item.Subject, item.Body);

            if (item.Attachments.Count <= 0) return mailMessage;

            foreach (Attachment attachment in item.Attachments)
            {
                try
                {
                    var tempFile = Path.GetTempFileName();
                    attachment.SaveAsFile(tempFile);
                    var ms = new MemoryStream(File.ReadAllBytes(tempFile));
                    mailMessage.Attachments.Add(new System.Net.Mail.Attachment(ms, attachment.FileName));
                    File.Delete(tempFile);
                }
                catch (Exception)
                {
                    // ignored
                }
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

        public bool UseSsl
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public string Host
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public int PortNumber
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
    }
}
