using Carubbi.Extensions;
using Carubbi.Mailer.DTOs;
using Carubbi.Mailer.Interfaces;
using Carubbi.ServiceLocator;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using System;
using System.Collections.Generic;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace Carubbi.Mailer.Implementation
{
    public class OpenPopMailReceiver : IMailReceiver
    {
        public string Username { get; set; }

        public string Password { get; set; }

        private readonly AppSettings config;

        const int DEFAULT_TIMEOUT = 60000;

        public OpenPopMailReceiver()
        {
            config = new AppSettings("CarubbiMailer");
        }

        public virtual bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;  // force the validation of any certificate
        }

        #region IMailReceiver Members

        private const int DEFAULT_SSL_POP_PORT = 995;
        private const int DEFAULT_NON_SSL_POP_PORT = 110;

        private bool? _useSsl;

        private string _host;

        private int? _portNumber;

        public bool UseSsl
        {
            get
            { 
                if (!_useSsl.HasValue)
                {
                    _useSsl = config["EnableSSLPOP"].To(false);
                }

                return _useSsl.Value;
            }
            set => _useSsl = value;
        }



        public string Host
        {
            get
            {
                if (string.IsNullOrEmpty(_host))
                {
                    _host = config["HostPOP"];
                }
                return _host;
            }
            set => _host = value;
        }

        public int PortNumber
        {
           get
           {
               if (!_portNumber.HasValue)
               {
                   _portNumber = config["PortNumberPOP"].To<int>(UseSsl ? DEFAULT_SSL_POP_PORT : DEFAULT_NON_SSL_POP_PORT);
               }
               return _portNumber.Value;
           }
            set => _portNumber = value;
        }

        public IEnumerable<System.Net.Mail.MailMessage> GetPendingMessages()
        {
            using (var client = new Pop3Client())
            {
                client.Connect(Host,
                    PortNumber,
                    UseSsl,
                    DEFAULT_TIMEOUT,
                    DEFAULT_TIMEOUT,
                    ValidateServerCertificate);

                client.Authenticate(Username, Password);

                // Obtém o número de mensagens na caixa de entrada
                var messageCount = client.GetMessageCount();

                var allMessages = new List<OpenPop.Mime.Message>(messageCount);

                // Mensagens são numeradas a partir do número 1
                for (var i = 1; i <= messageCount; i++)
                {
                    var mensagem = client.GetMessage(i);
                    
                    var m = mensagem.ToMailMessage();
                    var headers = mensagem.Headers;
                    
                    foreach (var mailAddress in headers.Bcc)
                        m.Bcc.Add(mailAddress.MailAddress);

                    foreach (var mailAddress in headers.Cc)
                        m.CC.Add(mailAddress.MailAddress);

                    m.Headers.Add("Date", headers.Date);
                    m.From = headers.From.MailAddress;

                    yield return m;

                    if (OnMessageRead != null)
                    {
                        var e = new OnMessageReadEventArgs();
                        OnMessageRead(this, e);
                        if (e.Cancel)
                            continue;
                    }
                    try
                    {
                        client.DeleteMessage(i);
                    }
                    catch (PopServerException)
                    { 
                        //Ignore
                    }
                }
            }
        }

        public int GetPendingMessagesCount()
        {
            throw new NotImplementedException();
        }

        public event EventHandler<OnMessageReadEventArgs> OnMessageRead;

        #endregion

        public void Dispose()
        {
            
        }
    }
}