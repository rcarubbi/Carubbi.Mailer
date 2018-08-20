using System;
using System.Collections.Generic;
using ex = Microsoft.Exchange.WebServices.Data;

using Carubbi.Mailer.Interfaces;
using Carubbi.Utils.Persistence;
using System.Net.Mail;
namespace Carubbi.Mailer.Exchange
{
    public class ExchangeWebServiceSender : ExchangeServiceBase, IMailSender
    {
        
        public bool UseSSL
        {
           get;
           set;

        }

        public string Host
        {
           get;
           set;
        }

        public int PortNumber
        {
           get;
           set;
        }    
        public string Username { get; set; }
        public string Password { get; set; }

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

        public void Send(MailMessage message)
        {
            ex.EmailMessage exchangeMessage = new ex.EmailMessage(GetExchangeService(Username, Password));

            if (message.From != null && !string.IsNullOrEmpty(message.From.Address))
            {
                exchangeMessage.From = new ex.EmailAddress(message.From.Address);
            }
            else
            {
                exchangeMessage.From = new ex.EmailAddress(_config["NOME_CAIXA"]);
            }

            exchangeMessage.Subject = message.Subject;
            exchangeMessage.Body = new ex.MessageBody(message.IsBodyHtml ? ex.BodyType.HTML : ex.BodyType.Text, message.Body);

            foreach (var destinatario in message.To)
            {
                exchangeMessage.ToRecipients.Add(destinatario.Address);
            }

            foreach (var copia in message.CC)
            {
                exchangeMessage.CcRecipients.Add(copia.Address);
            }

            foreach (var copiaOculta in message.Bcc)
            {
                exchangeMessage.BccRecipients.Add(copiaOculta.Address);
            }
           
            foreach (Attachment attachment in message.Attachments)
            {
                exchangeMessage.Attachments.AddFileAttachment(attachment.Name);
            }

            exchangeMessage.Send();
        }

        

        public void Dispose()
        {
            _instance = null;
	    _config = null;
        }
    }
}
