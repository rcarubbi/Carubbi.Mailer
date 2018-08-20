using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Carubbi.Mailer.DTOs;
using Carubbi.Mailer.Interfaces;
using Carubbi.Utils.Configuration;
using Carubbi.Utils.Data;
using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Mime.Header;
using OpenPop.Pop3.Exceptions;

namespace Carubbi.Mailer.Implementation
{
    public class OpenPopMailReceiver : IMailReceiver
    {
        public string Username { get; set; }
        public string Password { get; set; }
        private AppSettings config;
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

        private const int  DEFAULT_SSL_POP_PORT = 995;
        private const int DEFAULT_NON_SSL_POP_PORT = 110;


        private bool? _useSSL;
        private string _host;
        private int? _portNumber;

        public bool UseSSL
        {
            get
            { 
                if (!_useSSL.HasValue)
                {
                    _useSSL = config["EnableSSLPOP"].To<bool>(false);
                }

                return _useSSL.Value;
            }
            set
            {
                _useSSL = value;
            }
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
            set
            {
                _host = value;
            }
        }

        public int PortNumber
        {
           get
           {
               if (!_portNumber.HasValue)
               {
                   _portNumber = config["PortNumberPOP"].To<int>(UseSSL ? DEFAULT_SSL_POP_PORT : DEFAULT_NON_SSL_POP_PORT);
               }
               return _portNumber.Value;
           }
            set
            {
                _portNumber = value;
            }
        }

        public IEnumerable<System.Net.Mail.MailMessage> GetPendingMessages()
        {
            using (Pop3Client client = new Pop3Client())
            {
               
                client.Connect(Host,
                    PortNumber,
                    UseSSL,
                    DEFAULT_TIMEOUT,
                    DEFAULT_TIMEOUT,
                    ValidateServerCertificate);

                client.Authenticate(Username, Password);

                // Obtém o número de mensagens na caixa de entrada
                int messageCount = client.GetMessageCount();

                IList<OpenPop.Mime.Message> allMessages = new List<OpenPop.Mime.Message>(messageCount);

                // Mensagens são numeradas a partir do número 1
                for (int i = 1; i <= messageCount; i++)
                {
                    Message _mensagem = client.GetMessage(i);
                    
                    MailMessage m = _mensagem.ToMailMessage();
                    MessageHeader headers = _mensagem.Headers;
                    
                    foreach (var mailAddress in headers.Bcc)
                        m.Bcc.Add(mailAddress.MailAddress);

                    foreach (var mailAddress in headers.Cc)
                        m.CC.Add(mailAddress.MailAddress);

                    m.Headers.Add("Date", headers.Date);
                    m.From = headers.From.MailAddress;

                    yield return m;

                    if (OnMessageRead != null)
                    {
                        OnMessageReadEventArgs e = new OnMessageReadEventArgs();
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



