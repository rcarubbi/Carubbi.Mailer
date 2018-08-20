using System.Net.Mail;
using System.Collections.Generic;
using System;

namespace Carubbi.Mailer.Interfaces
{
    public interface IMailSender : IDisposable
    {
        void Send(MailMessage message);
        string Username { get; set; }
        string Password { get; set; }
        bool UseSSL { get; set; }
        string Host { get; set; }
        int PortNumber { get; set; }

        bool UseDefaultCredentials { get; set; }
    }
}
