using System;
using System.Net.Mail;

namespace Carubbi.Mailer.Interfaces
{
    public interface IMailSender : IDisposable
    {
        void Send(MailMessage message);

        string Username { get; set; }

        string Password { get; set; }

        bool UseSsl { get; set; }

        string Host { get; set; }

        int PortNumber { get; set; }

        bool UseDefaultCredentials { get; set; }
    }
}
