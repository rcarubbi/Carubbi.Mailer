using Carubbi.Mailer.Interfaces;
using System;
using System.Net.Mail;
using ex = Microsoft.Exchange.WebServices.Data;
namespace Carubbi.Mailer.Exchange
{
    public class ExchangeWebServiceSender : ExchangeServiceBase, IMailSender
    {
        public bool UseSsl { get; set; }

        public string Host { get; set; }

        public int PortNumber { get; set; }

        public string Username { get; set; }

        public string Password { get; set; }

        public bool UseDefaultCredentials
        {
            get => throw new NotSupportedException();

            set => throw new NotSupportedException();
        }

        public void Send(MailMessage message)
        {
            var exchangeMessage = new ex.EmailMessage(GetExchangeService(Username, Password));

            if (message.From != null && !string.IsNullOrEmpty(message.From.Address))
            {
                exchangeMessage.From = new ex.EmailAddress(message.From.Address);
            }
            else
            {
                exchangeMessage.From = new ex.EmailAddress(Config["NOME_CAIXA"]);
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
           
            foreach (var attachment in message.Attachments)
            {
                exchangeMessage.Attachments.AddFileAttachment(attachment.Name);
            }

            exchangeMessage.Send();
        }

        public void Dispose()
        {
            Instance = null;
	        Config = null;
        }
    }
}