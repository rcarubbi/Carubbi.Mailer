using System.Net;
using Carubbi.Utils.Configuration;
using Microsoft.Exchange.WebServices.Data;
using System;

namespace Carubbi.Mailer.Exchange
{
    public abstract class ExchangeServiceBase
    {
        protected AppSettings _config;

        public ExchangeServiceBase()
        {
            _config = new AppSettings("CarubbiMailer");
        }

        protected ExchangeService _instance = null;
        
      

        protected ExchangeService GetExchangeService(string username, string password)
        {
                if (_instance == null)
                {
                    SetupExchangeClient(username, password);
                }

                return _instance;
        }

        protected void SetupExchangeClient(string username, string password)
        {
            ServicePointManager.ServerCertificateValidationCallback =
                    delegate(
                        object sender,
                        System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };

            InitializeClient(username, password);
        }

        protected void InitializeClient(string username, string password)
        {

            string ewsURL = null;

            ewsURL = _config["URL_SERVICO_EXCHANGE"];

            _instance = new ExchangeService((ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), _config["VERSAO_EXCHANGE"]), TimeZoneInfo.Local);
            if (!string.IsNullOrEmpty(_config["WEB_PROXY"]))
            {
                WebProxy wp = new WebProxy(_config["WEB_PROXY"]);
                _instance.WebProxy = wp;

            }
            _instance.Credentials = new NetworkCredential(username, password);
            _instance.Url = new Uri(ewsURL);
        }
    }

}
