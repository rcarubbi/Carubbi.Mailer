using System.Collections.Generic;
using System;
using Carubbi.Mailer.DTOs;

namespace Carubbi.Mailer.Interfaces
{
    public interface IMailReceiver : IDisposable
    {
        IEnumerable<System.Net.Mail.MailMessage> GetPendingMessages();
        int GetPendingMessagesCount();
        event EventHandler<OnMessageReadEventArgs> OnMessageRead;
        string Username { get; set; }
        string Password { get; set; }
        bool UseSSL { get; set; }
        string Host { get; set; }
        int PortNumber { get; set; }
    }
}
