using Carubbi.Mailer.Interfaces;
using Carubbi.Utils.Configuration;
using Carubbi.Utils.Data;
using System.Net.Mail;
namespace Carubbi.Mailer.Implementation
{
    /// <summary>
    /// 
    /// </summary>
    public class SmtpSender : IMailSender
    {
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
                    _useDefaultCredentials = config["UseDefaultCredentials"].To<bool>(false);
                }

                return _useDefaultCredentials.Value;
            }
            set
            {
                _useDefaultCredentials = value;
            }
        }

        private bool? _useSSL;
        private string _host;
        private int? _portNumber;
        private bool? _useDefaultCredentials;

        /// <summary>
        /// 
        /// </summary>
        public bool UseSSL
        {
            get
            {
                if (!_useSSL.HasValue)
                {
                    _useSSL = config["EnableSSLSMTP"].To<bool>(false);
                }

                return _useSSL.Value;
            }
            set
            {
                _useSSL = value;
            }
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
            set
            {
                _host = value;
            }
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
                    _portNumber = config["PortNumberSMTP"].To<int>(UseSSL ? DEFAULT_SSL_SMTP_PORT : DEFAULT_NON_SSL_SMTP_PORT);
                }
                return _portNumber.Value;
            }
            set
            {
                _portNumber = value;
            }
        }


        private AppSettings config;

        /// <summary>
        /// 
        /// </summary>
        public SmtpSender()
        {
            config = new AppSettings("CarubbiMailer");
        }
        #region IMailSender Members


        private const int DEFAULT_SSL_SMTP_PORT = 465;
        private const int DEFAULT_NON_SSL_SMTP_PORT = 25;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public void Send(System.Net.Mail.MailMessage message)
        {
            SmtpClient smtp = null;
            if (UseDefaultCredentials)
            {
                smtp = new SmtpClient();
                smtp.UseDefaultCredentials = true;
                smtp.Host = Host;
                smtp.Port = PortNumber;
                smtp.EnableSsl = UseSSL;
            }
            else
            {
                smtp = new SmtpClient
                {
                    Host = this.Host,
                    EnableSsl = this.UseSSL,
                    Port = this.PortNumber,
                    Credentials = new System.Net.NetworkCredential(Username, Password),
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                };
            }
            smtp.Send(message);
            smtp = null;

        }

        #endregion


        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {

        }
    }
}
