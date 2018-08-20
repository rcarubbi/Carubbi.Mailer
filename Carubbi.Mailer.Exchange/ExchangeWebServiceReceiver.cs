using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using Carubbi.Utils.Data;
using Carubbi.Mailer.Interfaces;
using Carubbi.Mailer.DTOs;
using ex = Microsoft.Exchange.WebServices.Data;
using System.Linq;
namespace Carubbi.Mailer.Exchange
{
    public class ExchangeWebServiceReceiver : ExchangeServiceBase, IMailReceiver
    {
        #region Membros de IMailReceiver


        public IEnumerable<MailMessage> GetPendingMessages()
        {
            ex.FolderId inbox = new ex.FolderId(ex.WellKnownFolderName.Inbox, new ex.Mailbox(_config["NOME_CAIXA"]));
            ex.ItemView itemView = new ex.ItemView(_config["QTD_EMAILS_RECUPERAR"].To<int>(10));
            itemView.PropertySet = new ex.PropertySet(ex.BasePropertySet.IdOnly, ex.ItemSchema.Subject, ex.ItemSchema.DateTimeReceived);

            itemView.OrderBy.Add(ex.ItemSchema.DateTimeReceived, ex.SortDirection.Ascending);

            ex.FindItemsResults<ex.Item> findResults = GetExchangeService(this.Username, this.Password).FindItems(inbox, itemView);

            ex.ServiceResponseCollection<ex.GetItemResponse> items = GetExchangeService(this.Username, this.Password).BindToItems(findResults.Select(item => item.Id), new ex.PropertySet(ex.BasePropertySet.FirstClassProperties,
                ex.EmailMessageSchema.From,
                ex.EmailMessageSchema.ToRecipients,
                ex.EmailMessageSchema.Attachments,
                ex.EmailMessageSchema.CcRecipients,
                ex.EmailMessageSchema.BccRecipients,
                ex.EmailMessageSchema.Body,
                ex.EmailMessageSchema.DateTimeCreated,
                ex.EmailMessageSchema.DateTimeReceived,
                ex.EmailMessageSchema.DateTimeSent,
                ex.EmailMessageSchema.DisplayCc,
                ex.EmailMessageSchema.DisplayTo,
                ex.EmailMessageSchema.Subject));

            foreach (ex.GetItemResponse item in items)
            {
                yield return ParseItem(item);
                OnMessageReadEventArgs ea = new OnMessageReadEventArgs();
                if (OnMessageRead != null)
                {
                    OnMessageRead(this, ea);
                    if (ea.Cancel)
                        continue;
                }
                GetExchangeService(this.Username, this.Password).DeleteItems(new ex.ItemId[] { item.Item.Id }, ex.DeleteMode.MoveToDeletedItems, null, null);

            }

        }

        private MailMessage ParseItem(ex.GetItemResponse item)
        {
            MailMessage mailMessage = new MailMessage();
            var exchangeSenderEmail = (Microsoft.Exchange.WebServices.Data.EmailAddress)item.Item[ex.EmailMessageSchema.From];
            mailMessage.From = new MailAddress(exchangeSenderEmail.Address, exchangeSenderEmail.Name);

            foreach (var emailAddress in ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item.Item[ex.EmailMessageSchema.ToRecipients]))
            {
                mailMessage.To.Add(new MailAddress(emailAddress.Address, emailAddress.Name));
            }


            mailMessage.Subject = item.Item.Subject;
            mailMessage.Body = item.Item.Body.ToString();
            var attachmentCollection = item.Item[ex.EmailMessageSchema.Attachments];


            foreach (ex.Attachment attachment in (ex.AttachmentCollection)attachmentCollection)
            {
                if (attachment is ex.FileAttachment)
                {
                    ex.FileAttachment fileAttachment = (ex.FileAttachment)attachment;
                    fileAttachment.Load();

                    var tempFile = Path.GetTempFileName();
                    File.WriteAllBytes(tempFile, fileAttachment.Content);
                    MemoryStream ms = new MemoryStream(File.ReadAllBytes(tempFile));
                    mailMessage.Attachments.Add(new System.Net.Mail.Attachment(ms, attachment.Name ?? tempFile));
                    File.Delete(tempFile);
                }
            }

            return mailMessage;
        }


        public int GetPendingMessagesCount()
        {
            try
            {
                ex.FolderId inbox = new ex.FolderId(ex.WellKnownFolderName.Inbox, new ex.Mailbox(_config["NOME_CAIXA"]));
                ex.ItemView itemView = new ex.ItemView(_config["QTD_EMAILS_RECUPERAR"].To<int>(10));
                ex.FindItemsResults<ex.Item> items = GetExchangeService(this.Username, this.Password).FindItems(inbox, itemView);
                return items.Count();
            }
            catch
            {
                _instance = null;
                throw;
            }
        }


        public string Username { get; set; }
        public string Password { get; set; }

        public event EventHandler<OnMessageReadEventArgs> OnMessageRead;

        #endregion

        public void Dispose()
        {
            _instance = null;
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
