using System.Net;
using System.Net.Mail;
using Carubbi.Extensions;
using Carubbi.Mailer.Interfaces;
using Carubbi.ServiceLocator;

namespace Carubbi.Mailer.Implementation
{
    /// <summary>
    /// 
    /// </summary>
    public class SmtpSender : IMailSender
    {
        private const int DEFAULT_SSL_SMTP_PORT = 465;
        private const int DEFAULT_NON_SSL_SMTP_PORT = 25;

        private readonly AppSettings config;

        /// <summary>
        /// 
        /// </summary>
        public string Username { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public bool UseDefaultCredentials {
            get
            {
                if (!_useDefaultCredentials.HasValue)
                {
                    _useDefaultCredentials = config["UseDefaultCredentials"].To(false);
                }

                return _useDefaultCredentials != null && _useDefaultCredentials.Value;
            }
            set => _useDefaultCredentials = value;
        }

        private bool? _useSsl;
        private string _host;
        private int? _portNumber;
        private bool? _useDefaultCredentials;

        /// <summary>
        /// 
        /// </summary>
        public bool UseSsl
        {
            get
            {
                if (!_useSsl.HasValue)
                {
                    _useSsl = config["EnableSSLSMTP"].To(false);
                }

                return _useSsl != null && _useSsl.Value;
            }
            set => _useSsl = value;
        }

        /// <summary>
        /// 
        /// </summary>
        public string Host
        {
            get
            {
                if (string.IsNullOrEmpty(_host))
                {
                    _host = config["HostSMTP"];
                }
                return _host;
            }
            set => _host = value;
        }

        /// <summary>
        /// 
        /// </summary>
        public int PortNumber
        {
            get
            {
                if (!_portNumber.HasValue)
                {
                    _portNumber = config["PortNumberSMTP"].To(UseSsl ? DEFAULT_SSL_SMTP_PORT : DEFAULT_NON_SSL_SMTP_PORT);
                }

                return _portNumber ?? 0;
            }
            set => _portNumber = value;
        }

        /// <summary>
        /// 
        /// </summary>
        public SmtpSender()
        {
            config = new AppSettings("CarubbiMailer");
        }

        #region IMailSender Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public void Send(MailMessage message)
        {
            SmtpClient smtp;
            if (UseDefaultCredentials)
            {
                smtp = new SmtpClient
                {
                    UseDefaultCredentials = true,
                    Host = Host,
                    Port = PortNumber,
                    EnableSsl = UseSsl
                };
            }
            else
            {
                smtp = new SmtpClient
                {
                    Host = Host,
                    EnableSsl = UseSsl,
                    Port = PortNumber,
                    Credentials = new NetworkCredential(Username, Password),
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };
            }
            smtp.Send(message);
        }

        #endregion

        /// <inheritdoc />
        /// <summary>
        /// </summary>
        public void Dispose()
        {

        }
    }
}
