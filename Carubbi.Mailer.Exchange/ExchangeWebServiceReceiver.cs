using Carubbi.Extensions;
using Carubbi.Mailer.DTOs;
using Carubbi.Mailer.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using ex = Microsoft.Exchange.WebServices.Data;

namespace Carubbi.Mailer.Exchange
{
    public class ExchangeWebServiceReceiver : ExchangeServiceBase, IMailReceiver
    {
        #region Membros de IMailReceiver

        public IEnumerable<MailMessage> GetPendingMessages()
        {
            var inbox = new ex.FolderId(ex.WellKnownFolderName.Inbox, new ex.Mailbox(Config["NOME_CAIXA"]));
            var itemView = new ex.ItemView(Config["QTD_EMAILS_RECUPERAR"].To(10))
            {
                PropertySet = new ex.PropertySet(ex.BasePropertySet.IdOnly, ex.ItemSchema.Subject,
                    ex.ItemSchema.DateTimeReceived)
            };

            itemView.OrderBy.Add(ex.ItemSchema.DateTimeReceived, ex.SortDirection.Ascending);

            var findResults = GetExchangeService(Username, Password).FindItems(inbox, itemView);

            var items = GetExchangeService(Username, Password).BindToItems(findResults.Select(item => item.Id),
                new ex.PropertySet(ex.BasePropertySet.FirstClassProperties,
                ex.EmailMessageSchema.From,
                ex.EmailMessageSchema.ToRecipients,
                ex.ItemSchema.Attachments,
                ex.EmailMessageSchema.CcRecipients,
                ex.EmailMessageSchema.BccRecipients,
                ex.ItemSchema.Body,
                ex.ItemSchema.DateTimeCreated,
                ex.ItemSchema.DateTimeReceived,
                ex.ItemSchema.DateTimeSent,
                ex.ItemSchema.DisplayCc,
                ex.ItemSchema.DisplayTo,
                ex.ItemSchema.Subject));

            foreach (var item in items)
            {
                yield return ParseItem(item);
                var ea = new OnMessageReadEventArgs();
                if (OnMessageRead != null)
                {
                    OnMessageRead(this, ea);
                    if (ea.Cancel)
                        continue;
                }
                GetExchangeService(Username, Password).DeleteItems(new[] { item.Item.Id }, ex.DeleteMode.MoveToDeletedItems, null, null);

            }

        }
        private static MailMessage ParseItem(ex.GetItemResponse item)
        {
            var mailMessage = new MailMessage();
            var exchangeSenderEmail = (ex.EmailAddress)item.Item[ex.EmailMessageSchema.From];
            mailMessage.From = new MailAddress(exchangeSenderEmail.Address, exchangeSenderEmail.Name);

            foreach (var emailAddress in ((ex.EmailAddressCollection)item.Item[ex.EmailMessageSchema.ToRecipients]))
            {
                mailMessage.To.Add(new MailAddress(emailAddress.Address, emailAddress.Name));
            }

            mailMessage.Subject = item.Item.Subject;
            mailMessage.Body = item.Item.Body.ToString();
            var attachmentCollection = item.Item[ex.ItemSchema.Attachments];


            foreach (var attachment in (ex.AttachmentCollection)attachmentCollection)
            {
                if (!(attachment is ex.FileAttachment fileAttachment)) continue;
                fileAttachment.Load();

                var tempFile = Path.GetTempFileName();
                File.WriteAllBytes(tempFile, fileAttachment.Content);
                var ms = new MemoryStream(File.ReadAllBytes(tempFile));
                mailMessage.Attachments.Add(new Attachment(ms, attachment.Name ?? tempFile));
                File.Delete(tempFile);
            }

            return mailMessage;
        }

        public int GetPendingMessagesCount()
        {
            try
            {
                var inbox = new ex.FolderId(ex.WellKnownFolderName.Inbox, new ex.Mailbox(Config["NOME_CAIXA"]));
                var itemView = new ex.ItemView(Config["QTD_EMAILS_RECUPERAR"].To(10));
                var items = GetExchangeService(Username, Password).FindItems(inbox, itemView);
                return items.Count();
            }
            catch
            {
                Instance = null;
                throw;
            }
        }


        public string Username { get; set; }
        public string Password { get; set; }

        public event EventHandler<OnMessageReadEventArgs> OnMessageRead;

        #endregion

        public void Dispose()
        {
            Instance = null;
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