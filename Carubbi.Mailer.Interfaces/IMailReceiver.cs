using Carubbi.Mailer.DTOs;
using System;
using System.Collections.Generic;

namespace Carubbi.Mailer.Interfaces
{
    public interface IMailReceiver : IDisposable
    {
        IEnumerable<System.Net.Mail.MailMessage> GetPendingMessages();
        int GetPendingMessagesCount();

        event EventHandler<OnMessageReadEventArgs> OnMessageRead;

        string Username { get; set; }

        string Password { get; set; }

        bool UseSsl { get; set; }

        string Host { get; set; }

        int PortNumber { get; set; }
    }
}
